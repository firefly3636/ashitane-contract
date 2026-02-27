// ===== 定数：テンプレート内の置換対象テキスト =====
const ORIGINAL_NAME = "萩原 ヨシ子";
const ORIGINAL_ADDRESS = "千葉県柏市豊住3丁目16-4";
const ORIGINAL_POSTAL = "277-0071";

// ===== 郵便番号 → 住所自動入力 =====
async function lookupAddress() {
  const postalInput = document.getElementById("postalInput");
  const addressInput = document.getElementById("addressInput");
  const btn = document.getElementById("postalBtn");

  const raw = postalInput.value.replace(/[－‐ー−\-\s]/g, "");

  if (!/^\d{7}$/.test(raw)) {
    showToast("郵便番号は7桁の数字で入力してください（例: 277-0071）", "error");
    return;
  }

  btn.disabled = true;
  btn.textContent = "検索中...";

  try {
    const res = await fetch(
      `https://zipcloud.ibsnet.co.jp/api/search?zipcode=${raw}`
    );
    const data = await res.json();

    if (data.results && data.results.length > 0) {
      const r = data.results[0];
      const address = `${r.address1}${r.address2}${r.address3}`;
      addressInput.value = address;
      showToast(`住所を取得しました: ${address}`, "success");
    } else {
      showToast("該当する住所が見つかりませんでした", "error");
    }
  } catch {
    showToast("住所の取得に失敗しました。ネットワークを確認してください", "error");
  } finally {
    btn.disabled = false;
    btn.textContent = "住所検索";
  }
}

// Enterキーで住所検索
document.getElementById("postalInput").addEventListener("keydown", (e) => {
  if (e.key === "Enter") {
    e.preventDefault();
    lookupAddress();
  }
});

// ===== docx内のXMLテキストを置換する =====
async function replaceInDocx(templateUrl, name, address, postal) {
  const res = await fetch(templateUrl);
  const blob = await res.arrayBuffer();
  const zip = await JSZip.loadAsync(blob);

  const xmlFiles = [
    "word/document.xml",
    "word/header1.xml",
    "word/header2.xml",
    "word/header3.xml",
    "word/footer1.xml",
    "word/footer2.xml",
    "word/footer3.xml",
  ];

  for (const path of xmlFiles) {
    const file = zip.file(path);
    if (!file) continue;

    let xml = await file.async("string");

    // Word splits text across <w:r> runs. We need to handle replacement
    // on the concatenated text of <w:t> elements within a paragraph,
    // then reconstruct.
    xml = replaceInWordXml(xml, ORIGINAL_NAME, name);
    xml = replaceInWordXml(xml, ORIGINAL_ADDRESS, address);
    xml = replaceInWordXml(xml, ORIGINAL_POSTAL, postal);

    zip.file(path, xml);
  }

  return await zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}

// Word XMLでは1つのテキストが複数の<w:r>/<w:t>に分割されることがあるため、
// パラグラフ単位で結合してから置換し、最初のrunに全テキストを入れ残りを空にする
function replaceInWordXml(xml, oldText, newText) {
  if (!xml.includes("w:t")) return xml;

  // パラグラフ単位で処理
  const pRegex = /<w:p[\s>][\s\S]*?<\/w:p>/g;

  return xml.replace(pRegex, (pXml) => {
    // このパラグラフの全テキストを取得
    const tRegex = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g;
    const runs = [];
    let match;
    let fullText = "";

    while ((match = tRegex.exec(pXml)) !== null) {
      runs.push({ full: match[0], text: match[1], index: match.index });
      fullText += match[1];
    }

    if (!fullText.includes(oldText)) return pXml;

    // 置換を実行
    const newFullText = fullText.split(oldText).join(newText);

    if (runs.length === 0) return pXml;

    // 最初のrunに全テキストを入れ、残りは空にする
    let result = pXml;
    for (let i = runs.length - 1; i >= 0; i--) {
      const replacement = i === 0
        ? runs[i].full.replace(/>([^<]*)<\/w:t>/, ` xml:space="preserve">${newFullText}</w:t>`)
        : runs[i].full.replace(/>([^<]*)<\/w:t>/, ` xml:space="preserve"></w:t>`);
      result = result.substring(0, runs[i].index) + replacement + result.substring(runs[i].index + runs[i].full.length);
    }

    return result;
  });
}

// ===== xlsx内のテキストを置換する =====
async function replaceInXlsx(templateUrl, name, 番号) {
  const res = await fetch(templateUrl);
  const arrayBuf = await res.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuf);

  workbook.eachSheet((sheet) => {
    sheet.eachRow((row) => {
      row.eachCell((cell) => {
        if (typeof cell.value === "string" && cell.value.includes(ORIGINAL_NAME)) {
          cell.value = cell.value.replace(ORIGINAL_NAME, name);
        }
      });
    });
  });

  // 管理番号の置換（Sheet1のG1セル）
  if (番号) {
    const ws = workbook.getWorksheet(1);
    if (ws) {
      const cell = ws.getCell("G1");
      if (cell.value && typeof cell.value === "string") {
        cell.value = cell.value.replace(/\d+/, String(番号));
      }
    }
  }

  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

// ===== メインの書類生成処理 =====
document.getElementById("genForm").addEventListener("submit", async (e) => {
  e.preventDefault();

  const name = document.getElementById("nameInput").value.trim();
  const postal = document.getElementById("postalInput").value.trim() || ORIGINAL_POSTAL;
  const address = document.getElementById("addressInput").value.trim();
  const number = document.getElementById("numberInput").value.trim();

  if (!name || !address) {
    showToast("名前と住所は必須です", "error");
    return;
  }

  const btn = document.getElementById("submitBtn");
  const btnText = document.getElementById("btnText");
  const spinner = document.getElementById("spinner");

  btn.disabled = true;
  btnText.textContent = "生成中...";
  spinner.style.display = "block";

  try {
    // 3つのテンプレートを並列で処理
    const [contractBlob, importantBlob, coverBlob] = await Promise.all([
      replaceInDocx("templates/契約書テンプレート.docx", name, address, postal),
      replaceInDocx("templates/重要事項説明書テンプレート.docx", name, address, postal),
      replaceInXlsx("templates/表紙テンプレート.xlsx", name, number ? parseInt(number) : null),
    ]);

    // ZIPに詰める
    const zip = new JSZip();
    zip.file(`契約書_${name}.docx`, contractBlob);
    zip.file(`重要事項説明書_${name}.docx`, importantBlob);
    zip.file(`表紙_${name}.xlsx`, coverBlob);

    const zipBlob = await zip.generateAsync({ type: "blob" });
    saveAs(zipBlob, `契約書類_${name}.zip`);

    showToast("書類を生成しました！ダウンロードを確認してください", "success");
  } catch (err) {
    console.error(err);
    showToast(`エラーが発生しました: ${err.message}`, "error");
  } finally {
    btn.disabled = false;
    btnText.textContent = "書類を生成してダウンロード";
    spinner.style.display = "none";
  }
});

// ===== トースト通知 =====
function showToast(message, type) {
  const toast = document.getElementById("toast");
  toast.textContent = message;
  toast.className = "toast" + (type ? ` ${type}` : "");
  toast.style.display = "block";
  clearTimeout(toast._timer);
  toast._timer = setTimeout(() => {
    toast.style.display = "none";
  }, 4000);
}
