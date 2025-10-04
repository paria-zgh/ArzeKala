import React, { useState } from "react";
import ExcelJS from "exceljs";
import jalaali from "jalaali-js";

export default function ExcelProcessor() {
  const [data, setData] = useState([]);

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (evt) => {
      const arrayBuffer = evt.target?.result;
      if (!arrayBuffer) return;

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(arrayBuffer);
      const worksheet = workbook.worksheets[0];

      const jsonData = [];
      const columns = worksheet.getRow(1).values.slice(1).map((v) => String(v || "").trim());

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        const obj = {};
        columns.forEach((col, idx) => {
          const cellValue = row.getCell(idx + 1).value;
          obj[col] = cellValue ?? "";
        });
        jsonData.push(obj);
      });

      setData(jsonData);
    };
    reader.readAsArrayBuffer(file);
  };

  const applyAllBorders = (cell) => {
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };
  };

  const processData = async () => {
    if (data.length === 0) {
      alert("ابتدا یک فایل آپلود کنید.");
      return;
    }

    const sorted = [...data].sort((a, b) =>
      String(a["تالار"] || "").localeCompare(String(b["تالار"] || ""), "fa")
    );

    const keywordsSub = [
      "مس","فولاد","ضایعات","پالت چوبی","بشکه خالی","ورق",
      "پشم شیشه","کنسانتره","اقلام تجهیزات","اقلام تکمیل خودرو","شمش",
      "سپری","تختال","بتنی"
    ];
    const petrochemKeywords = ["پلی","الیاف استیپل اکریلیک"];

    const blocks = { keywordBlockSub: [], keywordBlockPetroleumFromSub: [], auctionBlockStoneOrCathode: [], auctionBlockVacuum: [], petrochemBlock: [], otherRows: [] };

    sorted.forEach((row) => {
      const curTalar = row["تالار"] || "";
      const namaKala = row["نام کالا"] || "";

      if (curTalar === "تالار فرعی") {
        if (["نفتی","نفت","وکیوم","قیر","روغن"].some(kw => namaKala.includes(kw))) {
          blocks.keywordBlockPetroleumFromSub.push(row);
        } else if (keywordsSub.some(kw => namaKala.includes(kw))) {
          blocks.keywordBlockSub.push(row);
        } else if (petrochemKeywords.some(kw => namaKala.includes(kw))) {
          blocks.petrochemBlock.push(row);
        } else blocks.petrochemBlock.push(row);
      } else if (curTalar === "تالار حراج باز") {
        if (namaKala.includes("سنگ") || namaKala.includes("مس کاتد")) {
          blocks.auctionBlockStoneOrCathode.push(row);
        } else if (namaKala.includes("وکیوم")) {
          blocks.auctionBlockVacuum.push(row);
        } else if (petrochemKeywords.some(kw => namaKala.includes(kw))) {
          blocks.petrochemBlock.push(row);
        } else blocks.petrochemBlock.push(row);
      } else blocks.otherRows.push(row);
    });

    const finalRows = [...blocks.otherRows];

    let industrialIndex = finalRows.map(r => r["تالار"]||"").lastIndexOf("تالار صنعتی");
    if (industrialIndex === -1) industrialIndex = finalRows.length - 1;
    if (blocks.keywordBlockSub.length>0){ finalRows.splice(industrialIndex+1,0,...blocks.keywordBlockSub); industrialIndex += blocks.keywordBlockSub.length; }
    if (blocks.auctionBlockStoneOrCathode.length>0){ finalRows.splice(industrialIndex+1,0,...blocks.auctionBlockStoneOrCathode); industrialIndex += blocks.auctionBlockStoneOrCathode.length; }

    let petroleumIndex = finalRows.map(r => r["تالار"]||"").lastIndexOf("تالار فرآورده های نفتی");
    if (petroleumIndex === -1) petroleumIndex = finalRows.length -1;
    if (blocks.auctionBlockVacuum.length>0){ finalRows.splice(petroleumIndex+1,0,...blocks.auctionBlockVacuum); petroleumIndex += blocks.auctionBlockVacuum.length; }
    if (blocks.keywordBlockPetroleumFromSub.length>0){ finalRows.splice(petroleumIndex+1,0,...blocks.keywordBlockPetroleumFromSub); petroleumIndex += blocks.keywordBlockPetroleumFromSub.length; }

    let petroIndex = finalRows.map(r => r["تالار"]||"").lastIndexOf("تالار پتروشیمی");
    if (petroIndex===-1) petroIndex=finalRows.length-1;
    if (blocks.petrochemBlock.length>0){ finalRows.splice(petroIndex+1,0,...blocks.petrochemBlock); petroIndex += blocks.petrochemBlock.length; }

    const insertTalarNames = ["تالار صنعتی","تالار فرآورده های نفتی","تالار حراج همزمان","تالار کالای صادراتی کيش"];
    let processed=[];
    const inserted = new Set();
    finalRows.forEach(row => { const curTalar=row["تالار"]||""; if(insertTalarNames.includes(curTalar)&&!inserted.has(curTalar)){ processed.push({}); inserted.add(curTalar); } processed.push(row); });

    const specialKeywords = ["بوتادین استایرن","استایرن بوتادین","تیشو","پلی","الیاف استیپل اکریلیک"];
    const firstInsertIndex = processed.findIndex(row=>Object.keys(row).length===0);
    if(firstInsertIndex>0){
      const beforeInsert = processed.slice(0,firstInsertIndex);
      const afterInsert = processed.slice(firstInsertIndex);
      const specialRows=[], normalRows=[];
      beforeInsert.forEach(row=>{
        const namaKala=row["نام کالا"]||"";
        if(specialKeywords.some(kw=>namaKala.includes(kw))) specialRows.push(row);
        else normalRows.push(row);
      });
      if(specialRows.length>0) processed=[...normalRows,{},...specialRows,...afterInsert];
    }

    const blocksSorted=[];
    let currentBlock=[];
    processed.forEach(row=>{
      const isInsert=Object.keys(row).length===0;
      if(isInsert){
        if(currentBlock.length>0){
          currentBlock.sort((a,b)=>String(a["نام کالا"]||"").localeCompare(String(b["نام کالا"]||""),"fa"));
          blocksSorted.push(...currentBlock);
          currentBlock=[];
        }
        blocksSorted.push(row);
      } else currentBlock.push(row);
    });
    if(currentBlock.length>0){ currentBlock.sort((a,b)=>String(a["نام کالا"]||"").localeCompare(String(b["نام کالا"]||""),"fa")); blocksSorted.push(...currentBlock); }

    // --- هدرها و rename ستون‌ها ---
    let headersArr = Array.from(new Set(blocksSorted.flatMap(r=>Object.keys(r)))).filter(h => h!=="تالار" && h!=="تاریخ عرضه");
    headersArr = headersArr.map(h=>{
      if(h==="قیمت") return "قیمت پایه"; // تغییر اصلی
      if(h==="کد") return "کد عرضه";
      if(h==="تسویه") return "نوع تسویه";
      if(h==="حداکثر افزایش حجم سفارش") return "حداکثر افزایش عرضه";
      return h;
    });

    const producer="تولید کننده کالا", delivery="محل تحویل";
    if(headersArr.includes(producer) && headersArr.includes(delivery)){
      const withoutProducer = headersArr.filter(h=>h!==producer);
      const idxDelivery = withoutProducer.indexOf(delivery);
      if(idxDelivery!==-1) withoutProducer.splice(idxDelivery,0,producer);
      headersArr=withoutProducer;
    }

    const priceIdx=headersArr.indexOf("قیمت پایه");
    if(priceIdx!==-1) headersArr.splice(priceIdx+1,0,"مقدار پایه");
    ["حجم","تعداد محموله"].forEach(col=>{ const idx=headersArr.indexOf(col); if(idx!==-1) headersArr.splice(idx,1); });

    const reordered = blocksSorted.map(row=>{
      const nr={};
      headersArr.forEach(h=>{
        const originalKey = h==="کد عرضه" ? "کد" :
                            h==="نوع تسویه" ? "تسویه" :
                            h==="حداکثر افزایش عرضه" ? "حداکثر افزایش حجم سفارش" :
                            h==="قیمت پایه" ? "قیمت" : h; // مقدار واقعی قیمت در ستون "قیمت پایه"
        if(h==="مقدار پایه"){ const volume=parseFloat(row?.["حجم"]); nr[h]=!isNaN(volume)?Math.round(volume/1000):""; } 
        else nr[h]=row?.[originalKey]??"";
      });
      return nr;
    });

    // --- ساخت Excel ---
    const workbook=new ExcelJS.Workbook();
    workbook.views=[{rightToLeft:true}];
    const worksheet = workbook.addWorksheet("نتیجه",{views:[{rightToLeft:true}]});

    // --- سه ردیف اول اصلی ---
    const titleText = "کارگزاری آینده نگر خوارزمی به مدیریت دکتر ذوقی 09123011311";
    const titleRow = worksheet.addRow([titleText]);
    titleRow.height=25;
    worksheet.mergeCells(1,1,1,headersArr.length);
    titleRow.eachCell(cell=>{
      cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFFFA500"}};
      cell.font={bold:true};
      cell.alignment={vertical:"middle",horizontal:"center"};
      applyAllBorders(cell);
    });

    const today=new Date(); let nextDay=new Date(today);
    let diffDays= today.getDay()===3?3:1;
    nextDay.setDate(today.getDate()+diffDays);
    const j=jalaali.toJalaali(nextDay);
    const jalaliDate=`${j.jy}/${String(j.jm).padStart(2,"0")}/${String(j.jd).padStart(2,"0")}`;
    const daysFa=["یکشنبه","دوشنبه","سه‌شنبه","چهارشنبه","پنجشنبه","جمعه","شنبه"];
    const dayFa = daysFa[nextDay.getDay()];
    const supplyText=`عرضه X روز ${dayFa} مورخ ${jalaliDate}`;
    const supplyRow = worksheet.addRow([supplyText]);
    supplyRow.height=20;
    worksheet.mergeCells(2,1,2,headersArr.length);
    supplyRow.eachCell(cell=>{
      cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFFFFF00"}};
      cell.font={bold:true};
      cell.alignment={vertical:"middle",horizontal:"center"};
      applyAllBorders(cell);
    });

    const headerRow = worksheet.addRow(headersArr);
    headerRow.height=20;
    headerRow.eachCell(cell=>{
      cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFD3D3D3"}};
      cell.font={bold:true};
      cell.alignment={vertical:"middle",horizontal:"center"};
      applyAllBorders(cell);
    });

    // --- داده‌ها و درج سه ردیف اول بالای هر insert ---
    reordered.forEach(row=>{
      const isInsert=Object.keys(row).length===0;
      if(isInsert){
        [titleRow, supplyRow, headerRow].forEach(r=>{
          const newRow = worksheet.addRow(r.values.slice(1));
          newRow.height=r.height;
          r.eachCell((cell, colNumber)=>{
            const newCell=newRow.getCell(colNumber);
            newCell.fill=cell.fill;
            newCell.font=cell.font;
            newCell.alignment=cell.alignment;
            applyAllBorders(newCell);
          });
        });

        const insertRow = worksheet.addRow([]);
        insertRow.eachCell(applyAllBorders);

      } else {
        const dataRow = worksheet.addRow(headersArr.map(h=>row[h]));
        dataRow.eachCell(applyAllBorders);
      }
    });

    const dateStr=new Date().toISOString().slice(0,10);
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer],{type:"application/octet-stream"});
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a");
    a.href=url; a.download=`خروجی-${dateStr}.xlsx`;
    document.body.appendChild(a); a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  return (
    <div className="container mt-4" dir="rtl">
      <h3>پردازش اکسل (RTL)</h3>
      <input type="file" accept=".xlsx,.xls" className="form-control mb-2" onChange={handleFileUpload}/>
      <button className="btn btn-primary" onClick={processData} disabled={data.length===0}>
        پردازش و دانلود خروجی
      </button>
    </div>
  );
}
