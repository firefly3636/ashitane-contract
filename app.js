// あした音 契約書類生成 - Cloudflare Pages用（静的版）
const ORIGINAL_NAME = "萩原 ヨシ子";
const ORIGINAL_ADDRESS = "〒277-0071　千葉県柏市豊住3丁目16-4";

// 郵便番号→住所
async function lookupAddress() {
  const raw = document.getElementById("postalInput").value.replace(/[－‐ー−\-\s]/g, "");
  if (!/^\d{7}$/.test(raw)) {
    showToast("郵便番号は7桁で入力してください", "error");
    return;
  }
  const btn = document.getElementById("postalBtn");
  btn.disabled = true;
  btn.textContent = "検索中...";
  try {
    const res = await fetch("https://zipcloud.ibsnet.co.jp/api/search?zipcode=" + raw);
    const data = await res.json();
    if (data.results && data.results.length > 0) {
      const r = data.results[0];
      document.getElementById("addressInput").value = r.address1 + r.address2 + r.address3;
      showToast("住所を取得しました", "success");
    } else {
      showToast("該当する住所が見つかりませんでした", "error");
    }
  } catch {
    showToast("住所の取得に失敗しました", "error");
  }
  btn.disabled = false;
  btn.textContent = "住所検索";
}

document.getElementById("postalBtn").addEventListener("click", lookupAddress);
document.getElementById("postalInput").addEventListener("keydown", function(e) {
  if (e.key === "Enter") { e.preventDefault(); lookupAddress(); }
});

// docx置換（run分割対応）
function escapeXml(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function replaceInDocxXml(xml, oldText, newText) {
  const pRe = /<w:p[\s>][\s\S]*?<\/w:p>/g;
  return xml.replace(pRe, function(pXml) {
    const tRe = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
    const parts = [];
    let m;
    while ((m = tRe.exec(pXml)) !== null) {
      parts.push({ full: m[0], text: m[1], start: m.index, end: m.index + m[0].length });
    }
    if (parts.length === 0) return pXml;
    const combined = parts.map(function(p) { return p.text; }).join("");
    if (combined.indexOf(oldText) === -1) return pXml;

    const replaced = combined.split(oldText).join(newText);
    let result = pXml;
    for (let i = parts.length - 1; i >= 0; i--) {
      const txt = i === 0 ? replaced : "";
      const newTag = '<w:t xml:space="preserve">' + escapeXml(txt) + '</w:t>';
      result = result.slice(0, parts[i].start) + newTag + result.slice(parts[i].end);
    }
    return result;
  });
}

async function processDocx(url, name, addressFull) {
  const res = await fetch(url);
  if (!res.ok) throw new Error("テンプレート読込失敗");
  const zip = await JSZip.loadAsync(await res.arrayBuffer());
  const xmlPaths = ["word/document.xml", "word/header1.xml", "word/header2.xml", "word/header3.xml", "word/footer1.xml", "word/footer2.xml", "word/footer3.xml"];
  for (const p of xmlPaths) {
    const f = zip.file(p);
    if (!f) continue;
    let xml = await f.async("string");
    xml = replaceInDocxXml(xml, ORIGINAL_NAME, name);
    xml = replaceInDocxXml(xml, ORIGINAL_ADDRESS, addressFull);
    zip.file(p, xml);
  }
  return await zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    compression: "DEFLATE"
  });
}

// xlsx置換
function escapeXmlXlsx(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

async function processXlsx(url, name, number) {
  const res = await fetch(url);
  if (!res.ok) throw new Error("テンプレート読込失敗");
  const zip = await JSZip.loadAsync(await res.arrayBuffer());
  const ss = zip.file("xl/sharedStrings.xml");
  if (ss) {
    let xml = await ss.async("string");
    xml = xml.split(ORIGINAL_NAME).join(escapeXmlXlsx(name));
    if (number != null && number !== "") {
      xml = xml.replace(/(<t(?:\s[^>]*)?>)(№[\s\u3000]*)\d+(<\/t>)/g, "$1$2" + number + "$3");
      xml = xml.replace(/(<t(?:\s[^>]*)?>)([\s\u3000]*)\d{4,6}(<\/t>)/g, "$1$2" + number + "$3");
    }
    zip.file("xl/sharedStrings.xml", xml);
  }
  const sheets = [];
  zip.forEach(function(path) {
    if (/^xl\/worksheets\/sheet\d+\.xml$/.test(path)) sheets.push(path);
  });
  for (const path of sheets) {
    let xml = await zip.file(path).async("string");
    xml = xml.split(ORIGINAL_NAME).join(escapeXmlXlsx(name));
    zip.file(path, xml);
  }
  return await zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    compression: "DEFLATE"
  });
}

// メイン
document.getElementById("genForm").addEventListener("submit", async function(e) {
  e.preventDefault();
  const name = document.getElementById("nameInput").value.trim();
  const postalRaw = document.getElementById("postalInput").value.trim();
  const address = document.getElementById("addressInput").value.trim();
  const number = document.getElementById("numberInput").value.trim();

  if (!name || !address) {
    showToast("名前と住所は必須です", "error");
    return;
  }

  const postalClean = postalRaw.replace(/[－‐ー−\-\s]/g, "");
  const postalFmt = (postalClean.length === 7 && /^\d+$/.test(postalClean))
    ? postalClean.slice(0, 3) + "-" + postalClean.slice(3) : (postalRaw || "277-0071");
  const addressFull = "〒" + postalFmt + "　" + address;

  const btn = document.getElementById("submitBtn");
  const btnText = document.getElementById("btnText");
  btn.disabled = true;
  btnText.textContent = "生成中...";

  try {
    const base = "templates/";
    const [doc1, doc2, xlsx] = await Promise.all([
      processDocx(base + encodeURIComponent("契約書テンプレート.docx"), name, addressFull),
      processDocx(base + encodeURIComponent("重要事項説明書テンプレート.docx"), name, addressFull),
      processXlsx(base + encodeURIComponent("表紙テンプレート.xlsx"), name, number || null)
    ]);

    const outZip = new JSZip();
    outZip.file("契約書_" + name + ".docx", doc1);
    outZip.file("重要事項説明書_" + name + ".docx", doc2);
    outZip.file("表紙_" + name + ".xlsx", xlsx);

    const blob = await outZip.generateAsync({ type: "blob" });
    saveAs(blob, "契約書類_" + name + ".zip");
    showToast("書類を生成しました", "success");
  } catch (err) {
    console.error(err);
    showToast("エラー: " + err.message, "error");
  } finally {
    btn.disabled = false;
    btnText.textContent = "書類を生成してダウンロード";
  }
});

function showToast(msg, type) {
  const t = document.getElementById("toast");
  t.textContent = msg;
  t.className = "toast" + (type ? " " + type : "");
  t.style.display = "block";
  setTimeout(function() { t.style.display = "none"; }, 4000);
}
