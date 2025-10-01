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
      const jsonData = XLSX.utils.sheet_to_json(ws, { defval: "" }).map((row) => {
        const nr = {};
        Object.entries(row).forEach(([k, v]) => {
          nr[k.trim()] = v;
        });
        return nr;
      });
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

    // --- دسته‌بندی ردیف‌ها ---
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
          blocks.petrochemBlock.push(row);
        }
      }
      else if (curTalar === "تالار حراج باز") {
        if (namaKala.includes("سنگ") || namaKala.includes("مس کاتد")) {
          blocks.auctionBlockStoneOrCathode.push(row);
        } else if (namaKala.includes("وکیوم")) {
          blocks.auctionBlockVacuum.push(row);
        } else if (petrochemKeywords.some(kw => namaKala.includes(kw))) {
          blocks.petrochemBlock.push(row);
        } else {
          blocks.petrochemBlock.push(row);
        }
      }
      else {
        blocks.otherRows.push(row);
      }
    });

    // شروع ترکیب نهایی
    const finalRows = [...blocks.otherRows];

    // --- اضافه کردن بلوک تالار فرعی زیر صنعتی ---
    let industrialIndex = finalRows.map(r => r["تالار"] || "").lastIndexOf("تالار صنعتی");
    if (industrialIndex === -1) industrialIndex = finalRows.length - 1;
    if (blocks.keywordBlockSub.length > 0) {
      finalRows.splice(industrialIndex + 1, 0, ...blocks.keywordBlockSub);
      industrialIndex += blocks.keywordBlockSub.length;
    }

    // --- اضافه کردن بلوک سنگ/مس کاتد زیر صنعتی ---
    if (blocks.auctionBlockStoneOrCathode.length > 0) {
      finalRows.splice(industrialIndex + 1, 0, ...blocks.auctionBlockStoneOrCathode);
      industrialIndex += blocks.auctionBlockStoneOrCathode.length;
    }

    // --- اضافه کردن بلوک وکیوم زیر فرآورده های نفتی ---
    let petroleumIndex = finalRows.map(r => r["تالار"] || "").lastIndexOf("تالار فرآورده های نفتی");
    if (petroleumIndex === -1) petroleumIndex = finalRows.length - 1;
    if (blocks.auctionBlockVacuum.length > 0) {
      finalRows.splice(petroleumIndex + 1, 0, ...blocks.auctionBlockVacuum);
      petroleumIndex += blocks.auctionBlockVacuum.length;
    }

    // --- اضافه کردن نفتی‌های تالار فرعی زیر فرآورده های نفتی ---
    if (blocks.keywordBlockPetroleumFromSub.length > 0) {
      finalRows.splice(petroleumIndex + 1, 0, ...blocks.keywordBlockPetroleumFromSub);
      petroleumIndex += blocks.keywordBlockPetroleumFromSub.length;
    }

    // --- اضافه کردن پتروشیمی ---
    let petroIndex = finalRows.map(r => r["تالار"] || "").lastIndexOf("تالار پتروشیمی");
    if (petroIndex === -1) petroIndex = finalRows.length - 1;
    if (blocks.petrochemBlock.length > 0) {
      finalRows.splice(petroIndex + 1, 0, ...blocks.petrochemBlock);
      petroIndex += blocks.petrochemBlock.length;
    }

    // --- اضافه کردن insert فقط بالای اولین ردیف تالارهای خاص ---
    const insertTalarNames = [
      "تالار صنعتی",
      "تالار فرآورده های نفتی",
      "تالار حراج همزمان",
      "تالار کالای صادراتی کيش"
    ];
    let processed = [];
    const inserted = new Set();
    finalRows.forEach(row => {
      const curTalar = row["تالار"] || "";
      if (insertTalarNames.includes(curTalar) && !inserted.has(curTalar)) {
        processed.push({}); // insert marker
        inserted.add(curTalar);
      }
      processed.push(row);
    });

    // --- بررسی ردیف‌ها تا اولین insert برای کلمات خاص ---
    const specialKeywords = ["بوتادین استایرن","استایرن بوتادین","تیشو", "پلی", "الیاف استیپل اکریلیک"];
    const firstInsertIndex = processed.findIndex(row => Object.keys(row).length === 0);

    if (firstInsertIndex > 0) {
      const beforeInsert = processed.slice(0, firstInsertIndex);
      const afterInsert = processed.slice(firstInsertIndex);

      const specialRows = [];
      const normalRows = [];

      beforeInsert.forEach(row => {
        const namaKala = row["نام کالا"] || "";
        if (specialKeywords.some(kw => namaKala.includes(kw))) {
          specialRows.push(row);
        } else {
          normalRows.push(row);
        }
      });

      if (specialRows.length > 0) {
        processed = [...normalRows, {}, ...specialRows, ...afterInsert];
      }
    }

    // --- تقسیم به بلوک‌ها و مرتب‌سازی داخل هر بلوک بر اساس "نام کالا" ---
    const blocksSorted = [];
    let currentBlock = [];

    processed.forEach(row => {
      const isInsert = Object.keys(row).length === 0;
      if (isInsert) {
        if (currentBlock.length > 0) {
          currentBlock.sort((a, b) =>
            String(a["نام کالا"] || "").localeCompare(String(b["نام کالا"] || ""), "fa")
          );
          blocksSorted.push(...currentBlock);
          currentBlock = [];
        }
        blocksSorted.push(row); // insert خودش
      } else {
        currentBlock.push(row);
      }
    });

    if (currentBlock.length > 0) {
      currentBlock.sort((a, b) =>
        String(a["نام کالا"] || "").localeCompare(String(b["نام کالا"] || ""), "fa")
      );
      blocksSorted.push(...currentBlock);
    }

    // --- محاسبه union headerها ---
    const headerSet = new Set();
    blocksSorted.forEach((r) => Object.keys(r).forEach((k) => headerSet.add(k)));
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
    const reordered = blocksSorted.map(row => {
      const nr = {};
      headersArr.forEach(h => {
        if (h === "مقدار پایه") {
          const volume = parseFloat(row?.["حجم"]);
          nr[h] = !isNaN(volume) ? (volume / 1000).toFixed(2) : "";
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

    // اسم فایل خروجی با تاریخ
    const dateStr = new Date().toISOString().slice(0,10); // YYYY-MM-DD
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
