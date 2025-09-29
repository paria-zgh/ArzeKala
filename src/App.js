import React, { useState } from "react";
import * as XLSX from "xlsx";

export default function ExcelProcessor() {
  const [data, setData] = useState([]);

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const arrayBuffer = evt.target?.result;
      if (!arrayBuffer) return;

      const wb = XLSX.read(new Uint8Array(arrayBuffer), { type: "array" });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const jsonData = XLSX.utils.sheet_to_json(ws, { defval: "" });
      setData(jsonData);
    };
    reader.readAsArrayBuffer(file);
  };

  const processData = () => {
    if (data.length === 0) {
      alert("ابتدا یک فایل آپلود کنید.");
      return;
    }

    const sorted = [...data].sort((a, b) =>
      String(a["تالار"] || "").localeCompare(String(b["تالار"] || ""), "fa")
    );

    // --- تعریف بلوک‌ها ---
    const keywordsSub = [
      "مس", "فولاد", "ضایعات", "پالت چوبی", "بشکه خالی", "ورق",
      "پشم شیشه", "کنسانتره", "اقلام تجهیزات", "اقلام تکمیل خودرو","شمش",
      "سپری", "تختال", "بتنی"
    ];
    const petrochemKeywords = ["پلی", "الیاف استیپل اکریلیک"];

    const blocks = {
      keywordBlockSub: [],
      keywordBlockPetroleumFromSub: [],
      auctionBlockStoneOrCathode: [],
      auctionBlockVacuum: [],
      petrochemBlock: [],
      otherRows: []
    };

    sorted.forEach((row) => {
      const curTalar = row["تالار"] || "";
      const namaKala = row["نام کالا"] || "";

      if (curTalar === "تالار فرعی") {
        if (["نفتی", "نفت", "وکیوم", "قیر", "روغن"].some(kw => namaKala.includes(kw))) {
          blocks.keywordBlockPetroleumFromSub.push(row);
        } else if (keywordsSub.some(kw => namaKala.includes(kw))) {
          blocks.keywordBlockSub.push(row);
        } else if (petrochemKeywords.some(kw => namaKala.includes(kw))) {
          blocks.petrochemBlock.push(row);
        } else {
          blocks.otherRows.push(row);
        }
      } else if (curTalar === "تالار حراج باز") {
        if (namaKala.includes("سنگ") || namaKala.includes("مس کاتد")) {
          blocks.auctionBlockStoneOrCathode.push(row);
        } else if (namaKala.includes("وکیوم")) {
          blocks.auctionBlockVacuum.push(row);
        } else if (petrochemKeywords.some(kw => namaKala.includes(kw))) {
          blocks.petrochemBlock.push(row);
        } else {
          blocks.otherRows.push(row);
        }
      } else {
        blocks.otherRows.push(row);
      }
    });

    // --- شروع ترکیب نهایی ---
    const finalRows = [...blocks.otherRows];

    let industrialIndex = finalRows.map(r => r["تالار"] || "").lastIndexOf("تالار صنعتی");
    if (industrialIndex === -1) industrialIndex = finalRows.length - 1;
    if (blocks.keywordBlockSub.length > 0) {
      finalRows.splice(industrialIndex + 1, 0, ...blocks.keywordBlockSub);
      industrialIndex += blocks.keywordBlockSub.length;
    }
    if (blocks.auctionBlockStoneOrCathode.length > 0) {
      finalRows.splice(industrialIndex + 1, 0, ...blocks.auctionBlockStoneOrCathode);
      industrialIndex += blocks.auctionBlockStoneOrCathode.length;
    }

    let petroleumIndex = finalRows.map(r => r["تالار"] || "").lastIndexOf("تالار فرآورده های نفتی");
    if (petroleumIndex === -1) petroleumIndex = finalRows.length - 1;
    if (blocks.auctionBlockVacuum.length > 0) {
      finalRows.splice(petroleumIndex + 1, 0, ...blocks.auctionBlockVacuum);
      petroleumIndex += blocks.auctionBlockVacuum.length;
    }
    if (blocks.keywordBlockPetroleumFromSub.length > 0) {
      finalRows.splice(petroleumIndex + 1, 0, ...blocks.keywordBlockPetroleumFromSub);
      petroleumIndex += blocks.keywordBlockPetroleumFromSub.length;
    }

    let petroIndex = finalRows.map(r => r["تالار"] || "").lastIndexOf("تالار پتروشیمی");
    if (petroIndex === -1) petroIndex = finalRows.length - 1;
    if (blocks.petrochemBlock.length > 0) {
      finalRows.splice(petroIndex + 1, 0, ...blocks.petrochemBlock);
      petroIndex += blocks.petrochemBlock.length;
    }

    // --- اضافه کردن insert فقط قبل از اولین ردیف تالارهای مشخص ---
    const insertTalarNames = [
      "تالار صنعتی",
      "تالار فرآورده های نفتی",
      "تالار کالای صادراتی کيش"
    ];

    const processed = [];
    const inserted = new Set();

    finalRows.forEach(row => {
      const curTalar = row["تالار"] || "";
      if (insertTalarNames.includes(curTalar) && !inserted.has(curTalar)) {
        processed.push({}); // insert قبل از اولین ردیف تالار مشخص
        inserted.add(curTalar);
      }
      processed.push(row);
    });

    // --- محاسبه union headerها ---
    const headerSet = new Set();
    processed.forEach((r) => Object.keys(r).forEach((k) => headerSet.add(k)));
    const headersArr = Array.from(headerSet);

    // ترتیب ستون‌ها
    const producer = "تولید کننده کالا";
    const delivery = "محل تحویل";
    if (headersArr.includes(producer) && headersArr.includes(delivery)) {
      const withoutProducer = headersArr.filter((h) => h !== producer);
      const idxDelivery = withoutProducer.indexOf(delivery);
      if (idxDelivery !== -1) withoutProducer.splice(idxDelivery, 0, producer);
      headersArr.length = 0;
      withoutProducer.forEach((h) => headersArr.push(h));
    }

    // اضافه کردن مقدار پایه
    const priceIdx = headersArr.indexOf("قیمت");
    if (priceIdx !== -1) headersArr.splice(priceIdx + 1, 0, "مقدار پایه");

    // حذف ستون‌های غیرضروری
    const excludedColumns = ["حجم", "تعداد محموله"];
    const filteredHeaders = headersArr.filter(h => !excludedColumns.includes(h));

    // بازسازی ردیف‌ها
    const reordered = processed.map(row => {
      const nr = {};
      filteredHeaders.forEach(h => {
        if (h === "مقدار پایه") {
          const volume = parseFloat(row?.["حجم"]);
          nr[h] = !isNaN(volume) ? volume / 1000 : "";
        } else {
          nr[h] = row?.[h] ?? "";
        }
      });
      return nr;
    });

    // ساخت فایل Excel
    const ws = XLSX.utils.json_to_sheet(reordered, { header: filteredHeaders, skipHeader: false });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "نتیجه");
    wb.Workbook = wb.Workbook || { Views: [{}] };
    wb.Workbook.Views[0].RTL = true;
    ws["!rtl"] = true;

    // اسم فایل خروجی با تاریخ
    const dateStr = new Date().toLocaleDateString("fa-IR").replace(/\//g, "-");
    XLSX.writeFile(wb, `خروجی-${dateStr}.xlsx`);
  };

  return (
    <div className="container mt-4" dir="rtl">
      <h3>پردازش اکسل (RTL)</h3>

      <input
        type="file"
        accept=".xlsx,.xls"
        className="form-control mb-2"
        onChange={handleFileUpload}
      />

      <button className="btn btn-primary" onClick={processData} disabled={data.length === 0}>
        پردازش و دانلود خروجی
      </button>
    </div>
  );
}
