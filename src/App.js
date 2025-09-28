import React, { useState } from "react";
import * as XLSX from "xlsx";

export default function ExcelProcessor() {
  const [data, setData] = useState([]);

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const data = new Uint8Array(evt.target.result);
      const wb = XLSX.read(data, { type: "array" });
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

    // 1) مرتب‌سازی بر اساس "تالار"
    const sorted = [...data].sort((a, b) =>
      String(a["تالار"] || "").localeCompare(String(b["تالار"] || ""), "fa")
    );

    // 2) جابه‌جایی کالاهای مشخص به تالار "صنعتی" بعد از مرتب‌سازی
    const industrialKeywords = [
      "سنگ","آهن","مس","شمش فولاد","ضایعات","پالت چوبی","بشکه خالی",
      "ورق","پشم شیشه","کنسانتره","اقلام تجهیزات","اقلام تکمیل خودرو",
      "سپری","تختال بتنی","تراورس"
    ];

    const processed = sorted.map((row) => {
      const name = row["نام کالا"] ? String(row["نام کالا"]).trim() : "";
      const hasAuctionAndSub = row["تالار"] === "حراج باز" 
                               && row["تالار فرعی"] !== undefined 
                               && row["تالار فرعی"] !== null 
                               && String(row["تالار فرعی"]).trim() !== "";
      const isIndustrial = industrialKeywords.some(keyword => name.includes(keyword));

      if (hasAuctionAndSub && isIndustrial) {
        return { ...row, "تالار": "صنعتی" };
      }
      return row;
    });

    // 3) اضافه کردن ردیف خالی بین تغییر تالارها
    const finalRows = [];
    let prevTalar = null;
    processed.forEach((row) => {
      const cur = row["تالار"] || "";
      if (prevTalar !== null && cur !== prevTalar) {
        finalRows.push({});
      }
      finalRows.push(row);
      prevTalar = cur;
    });

    // 4) محاسبه unionِ headerها
    const headerSet = new Set();
    finalRows.forEach((r) => Object.keys(r).forEach((k) => headerSet.add(k)));
    const headersArr = Array.from(headerSet);

    // 5) ترتیب ستون‌ها (تولیدکننده قبل از محل تحویل)
    const producer = "تولید کننده کالا";
    const delivery = "محل تحویل";
    if (headersArr.includes(producer) && headersArr.includes(delivery)) {
      const withoutProducer = headersArr.filter((h) => h !== producer);
      const idxDelivery = withoutProducer.indexOf(delivery);
      if (idxDelivery !== -1) withoutProducer.splice(idxDelivery, 0, producer);
      headersArr.length = 0;
      withoutProducer.forEach((h) => headersArr.push(h));
    }

    // 6) اضافه کردن ستون "مقدار پایه" بعد از "قیمت"
    const priceIdx = headersArr.indexOf("قیمت");
    if (priceIdx !== -1) {
      headersArr.splice(priceIdx + 1, 0, "مقدار پایه");
    }

    // 7) حذف ستون‌های "حجم" و "تعداد محموله"
    ["حجم", "تعداد محموله"].forEach((col) => {
      const idx = headersArr.indexOf(col);
      if (idx !== -1) headersArr.splice(idx, 1);
    });

    // 8) بازسازی ردیف‌ها با پر کردن مقادیر خالی + محاسبه مقدار پایه
    const reordered = finalRows.map((row) => {
      const nr = {};
      headersArr.forEach((h) => {
        if (h === "مقدار پایه") {
          const volume = row && row["حجم"] ? parseFloat(row["حجم"]) : 0;
          nr[h] = volume ? volume / 1000 : "";
        } else {
          nr[h] = row && Object.prototype.hasOwnProperty.call(row, h) ? row[h] : "";
        }
      });
      return nr;
    });

    // 9) تبدیل به sheet و دانلود
    const ws = XLSX.utils.json_to_sheet(reordered, { header: headersArr, skipHeader: false });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "نتیجه");
    if (!wb.Workbook) wb.Workbook = {};
    if (!wb.Workbook.Views) wb.Workbook.Views = [{}];
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
