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
      "پشم شیشه", "کنسانتره", "اقلام تجهیزات", "اقلام تکمیل خودرو",
      "سپری", "تختال", "بتنی"
    ];

    const mainRows = [];
    const keywordBlockSub = [];
    const keywordBlockPetroleumFromSub = [];
    const auctionBlockStoneOrCathode = [];
    const auctionBlockVacuum = [];

    sorted.forEach((row) => {
      const curTalar = row["تالار"] || "";
      const namaKala = row["نام کالا"] || "";

      // بلوک تالار فرعی
      if (curTalar === "تالار فرعی") {
        if (["نفتی", "نفت", "وکیوم", "قیر", "روغن"].some(kw => namaKala.includes(kw))) {
          keywordBlockPetroleumFromSub.push(row);
        } else if (keywordsSub.some(kw => namaKala.includes(kw))) {
          keywordBlockSub.push(row);
        }
      }
      // تالار حراج باز که نام کالا شامل "سنگ" یا "مس کاتد" است
      else if (curTalar === "تالار حراج باز" &&
               (namaKala.includes("سنگ") || namaKala.includes("مس کاتد"))) {
        auctionBlockStoneOrCathode.push(row);
      }
      // تالار حراج باز که نام کالا شامل "وکیوم" است
      else if (curTalar === "تالار حراج باز" && namaKala.includes("وکیوم")) {
        auctionBlockVacuum.push(row);
      }
      // سایر ردیف‌ها
      else {
        mainRows.push(row);
      }
    });

    // --- پیدا کردن آخرین تالار صنعتی ---
    let industrialIndex = mainRows.map(r => r["تالار"] || "").lastIndexOf("تالار صنعتی");
    if (industrialIndex === -1) industrialIndex = mainRows.length - 1;

    // --- اضافه کردن بلوک تالار فرعی زیر صنعتی ---
    const finalRows = [...mainRows];
    if (keywordBlockSub.length > 0) {
      finalRows.splice(industrialIndex + 1, 0, ...keywordBlockSub);
      industrialIndex += keywordBlockSub.length;
    }

    // --- اضافه کردن بلوک ردیف‌های "سنگ یا مس کاتد" زیر صنعتی ---
    if (auctionBlockStoneOrCathode.length > 0) {
      finalRows.splice(industrialIndex + 1, 0, ...auctionBlockStoneOrCathode);
      industrialIndex += auctionBlockStoneOrCathode.length;
    }

    // --- اضافه کردن بلوک ردیف‌های "وکیوم" تالار حراج ---
    let petroleumIndex = finalRows.map(r => r["تالار"] || "").lastIndexOf("تالار فرآورده های نفتی");
    if (petroleumIndex === -1) petroleumIndex = finalRows.length - 1;

    if (auctionBlockVacuum.length > 0) {
      finalRows.splice(petroleumIndex + 1, 0, ...auctionBlockVacuum);
      petroleumIndex += auctionBlockVacuum.length;
    }

    // --- اضافه کردن بلوک ردیف‌های نفتی از تالار فرعی زیر تالار فرآورده های نفتی ---
    if (keywordBlockPetroleumFromSub.length > 0) {
      finalRows.splice(petroleumIndex + 1, 0, ...keywordBlockPetroleumFromSub);
      petroleumIndex += keywordBlockPetroleumFromSub.length;
    }

    // --- اضافه کردن insert بعد از هر تالار ---
    const processed = [];
    let currentTalar = null;
    finalRows.forEach(row => {
      const curTalar = row["تالار"] || "";
      if (curTalar !== currentTalar && currentTalar !== null) {
        processed.push({}); // insert بعد از اتمام هر تالار
      }
      processed.push(row);
      currentTalar = curTalar;
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
    ["حجم", "تعداد محموله"].forEach(col => {
      const idx = headersArr.indexOf(col);
      if (idx !== -1) headersArr.splice(idx, 1);
    });

    // بازسازی ردیف‌ها
    const reordered = processed.map(row => {
      const nr = {};
      headersArr.forEach(h => {
        if (h === "مقدار پایه") {
          const volume = row?.حجم ? parseFloat(row["حجم"]) : 0;
          nr[h] = volume ? volume / 1000 : "";
        } else {
          nr[h] = row?.[h] ?? "";
        }
      });
      return nr;
    });

    // ساخت فایل Excel
    const ws = XLSX.utils.json_to_sheet(reordered, { header: headersArr, skipHeader: false });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "نتیجه");
    wb.Workbook = wb.Workbook || { Views: [{}] };
    wb.Workbook.Views[0].RTL = true;
    ws["!rtl"] = true;

    XLSX.writeFile(wb, "خروجی.xlsx");
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

      <button className="btn btn-primary" onClick={processData}>
        پردازش و دانلود خروجی
      </button>
    </div>
  );
}
