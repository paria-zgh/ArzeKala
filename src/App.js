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
      "پشم شیشه","کنسانتره","اقلام تجهیزات","اقلام تکمیلی خودرو",
      "شمش","سپری","تختال","بتنی","گرانول نقره","فلز"
    ];
    const petrochemKeywords = ["پلی","الیاف استیپل اکریلیک"];

    const blocks = {
      keywordBlockSub: [],
      keywordBlockPetroleumFromSub: [],
      auctionBlockStoneOrCathode: [],
      auctionBlockVacuum: [],
      petrochemBlock: [],
      otherRows: [],
    };

    sorted.forEach((row) => {
      const curTalar = row["تالار"] || "";
      const namaKala = row["نام کالا"] || "";

      if(namaKala.includes("پلیمریک")){
        blocks.petrochemBlock.push(row);
        return;
      }

      if(curTalar==="تالار فرعی"){
        if(["نفتی","نفت","وکیوم","قیر","روغن"].some(kw=>namaKala.includes(kw))){
          blocks.keywordBlockPetroleumFromSub.push(row);
        } else if(keywordsSub.some(kw=>namaKala.includes(kw))){
          blocks.keywordBlockSub.push(row);
        } else if(petrochemKeywords.some(kw=>namaKala.includes(kw))){
          blocks.petrochemBlock.push(row);
        } else blocks.petrochemBlock.push(row);
      } else if(curTalar==="تالار حراج باز"){
        if(namaKala.includes("سنگ") || namaKala.includes("مس کاتد") || namaKala.includes("فلز")){
          blocks.auctionBlockStoneOrCathode.push(row);
        } else if(namaKala.includes("وکیوم")){
          blocks.auctionBlockVacuum.push(row);
        } else if(petrochemKeywords.some(kw=>namaKala.includes(kw))){
          blocks.petrochemBlock.push(row);
        } else blocks.petrochemBlock.push(row);
      } else {
        blocks.otherRows.push(row);
      }
    });

    const finalRows = [...blocks.otherRows];
    let industrialIndex = finalRows.map(r => r["تالار"] || "").lastIndexOf("تالار صنعتی");
    if(industrialIndex===-1) industrialIndex=finalRows.length-1;
    if(blocks.keywordBlockSub.length>0){
      finalRows.splice(industrialIndex+1,0,...blocks.keywordBlockSub);
      industrialIndex+=blocks.keywordBlockSub.length;
    }
    if(blocks.auctionBlockStoneOrCathode.length>0){
      finalRows.splice(industrialIndex+1,0,...blocks.auctionBlockStoneOrCathode);
      industrialIndex+=blocks.auctionBlockStoneOrCathode.length;
    }

    let petroleumIndex = finalRows.map(r=>r["تالار"]||"").lastIndexOf("تالار فرآورده های نفتی");
    if(petroleumIndex===-1) petroleumIndex=finalRows.length-1;
    if(blocks.auctionBlockVacuum.length>0){
      finalRows.splice(petroleumIndex+1,0,...blocks.auctionBlockVacuum);
      petroleumIndex+=blocks.auctionBlockVacuum.length;
    }
    if(blocks.keywordBlockPetroleumFromSub.length>0){
      finalRows.splice(petroleumIndex+1,0,...blocks.keywordBlockPetroleumFromSub);
      petroleumIndex+=blocks.keywordBlockPetroleumFromSub.length;
    }

    let petroIndex = finalRows.map(r=>r["تالار"]||"").lastIndexOf("تالار پتروشیمی");
    if(petroIndex===-1) petroIndex=finalRows.length-1;
    if(blocks.petrochemBlock.length>0){
      finalRows.splice(petroIndex+1,0,...blocks.petrochemBlock);
      petroIndex+=blocks.petrochemBlock.length;
    }

    const insertTalarNames=["تالار صنعتی","تالار فرآورده های نفتی","تالار سیمان","تالار کالای صادراتی کيش"];
    let processed=[];
    const inserted=new Set();
    finalRows.forEach(row=>{
      const curTalar=row["تالار"]||"";
      if(insertTalarNames.includes(curTalar) && !inserted.has(curTalar)){
        processed.push({});
        inserted.add(curTalar);
      }
      processed.push(row);
    });

    const specialKeywords=["بوتادین استایرن","استایرن بوتادین","تیشو","پلی","الیاف استیپل اکریلیک"];
    const firstInsertIndex=processed.findIndex(row=>Object.keys(row).length===0);
    if(firstInsertIndex>0){
      const beforeInsert=processed.slice(0,firstInsertIndex);
      const afterInsert=processed.slice(firstInsertIndex);
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
          currentBlock.sort((a,b)=>{
            if((a["تالار"]||"")==="تالار سیمان"&&(b["تالار"]||"")==="تالار سیمان"){
              return String(a["تولید کننده کالا"]||"").localeCompare(String(b["تولید کننده کالا"]||""),"fa");
            }
            return String(a["نام کالا"]||"").localeCompare(String(b["نام کالا"]||""),"fa");
          });
          blocksSorted.push({__HEADER__:true});
          blocksSorted.push(...currentBlock);
          currentBlock=[];
        }
        blocksSorted.push(row);
      } else currentBlock.push(row);
    });
    if(currentBlock.length>0){
      currentBlock.sort((a,b)=>{
        if((a["تالار"]||"")==="تالار سیمان"&&(b["تالار"]||"")==="تالار سیمان"){
          return String(a["تولید کننده کالا"]||"").localeCompare(String(b["تولید کننده کالا"]||""),"fa");
        }
        return String(a["نام کالا"]||"").localeCompare(String(b["نام کالا"]||""),"fa");
      });
      blocksSorted.push({__HEADER__:true});
      blocksSorted.push(...currentBlock);
    }

    let headersArr=Array.from(new Set(blocksSorted.flatMap(r=>Object.keys(r)))).filter(
      h=>h!=="تاریخ عرضه" && h!=="__HEADER__" && h!=="تعداد محموله" && h!=="قیمت پایه"
    );
    headersArr=headersArr.map(h=>{
      if(h==="حجم") return "مقدار پایه";
      if(h==="قیمت") return "قیمت پایه";
      if(h==="کد") return "کد عرضه";
      if(h==="تسویه") return "نوع تسویه";
      if(h==="حداکثر افزایش حجم سفارش") return "حداکثر افزایش عرضه";
      return h;
    });

    const producer="تولید کننده کالا", delivery="محل تحویل";
    if(headersArr.includes(producer) && headersArr.includes(delivery)){
      const withoutProducer=headersArr.filter(h=>h!==producer);
      const idxDelivery=withoutProducer.indexOf(delivery);
      if(idxDelivery!==-1) withoutProducer.splice(idxDelivery,0,producer);
      headersArr=withoutProducer;
    }

    const reordered=blocksSorted.map(row=>{
      if(row.__HEADER__) return row;
      const nr={};
      headersArr.forEach(h=>{
        let originalKey=
          h==="مقدار پایه"?"حجم":
          h==="قیمت پایه"?"قیمت":
          h==="کد عرضه"?"کد":
          h==="نوع تسویه"?"تسویه":
          h==="حداکثر افزایش عرضه"?"حداکثر افزایش حجم سفارش":
          h;
        let value=row[originalKey]??"";
        if(h==="مقدار پایه"){
          const vol=Number(value);
          nr[h]=isNaN(vol)?"":vol/1000;
        } else if(h==="قیمت پایه"){
          const price=Number(value);
          nr[h]=isNaN(price)?"":price.toLocaleString("en-US");
        } else nr[h]=value;
      });
      return nr;
    });

    // --- بلوک‌های ویژه انتهایی ---
    const addSpecialBlock=(rows)=> {
      if(rows.length>0){
        reordered.push({__HEADER__:true});
        rows.forEach(r=>reordered.push({...r}));
      }
    };

    // بلوک‌های قبلی (پلی و سبک/سنگین و ...) بدون تغییر
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی") && r["نام کالا"]?.includes("بطری")));
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی") && r["نام کالا"]?.includes("نساجی") && !r["نام کالا"]?.toLowerCase().includes("off")));
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی") && r["نام کالا"]?.includes("نساجی") && r["نام کالا"]?.toLowerCase().includes("off")));
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی") && r["نام کالا"]?.includes("شیمیایی") && !r["نام کالا"]?.toLowerCase().includes("off")));
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی") && r["نام کالا"]?.includes("شیمیایی") && r["نام کالا"]?.toLowerCase().includes("off")));
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی") && r["نام کالا"]?.includes("وینیل") && r["نام کالا"]?.includes("کلراید")&& r["نام کالا"]?.includes("S")));
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی") && r["نام کالا"]?.includes("وینیل") && r["نام کالا"]?.includes("کلراید")&& r["نام کالا"]?.includes("E")));
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی") && r["نام کالا"]?.includes("استایرن"))) ;
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("منو اتیلن گلایکول") || r["نام کالا"]?.includes("دی اتیلن گلایکول")));
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("تری اتیلن گلایکول"))) ;
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("اسید ترفتالیک"))) ;
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی") && r["نام کالا"]?.includes("سبک") && r["نام کالا"]?.includes("تزریقی") && !r["نام کالا"]?.toLowerCase().includes("off")));
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی") && r["نام کالا"]?.includes("سبک") && r["نام کالا"]?.includes("خطی") && !r["نام کالا"]?.toLowerCase().includes("off")));
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی") && r["نام کالا"]?.includes("سبک") && r["نام کالا"]?.includes("فیلم") && !r["نام کالا"]?.toLowerCase().includes("off")));
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی") && r["نام کالا"]?.includes("سبک") && r["نام کالا"]?.toLowerCase().includes("off")));
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی") && r["نام کالا"]?.includes("سنگین") &&
      (r["نام کالا"]?.includes("اکستروژن") || r["نام کالا"]?.includes("لوله")) && !r["نام کالا"]?.toLowerCase().includes("off")));
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی") && r["نام کالا"]?.includes("سنگین") &&
      (r["نام کالا"]?.includes("PEWAX") || r["نام کالا"]?.includes("کلوخه")) && !r["نام کالا"]?.toLowerCase().includes("off")));
    ["بادی","تزریقی","فیلم","دورانی"].forEach(k=>{
      addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی") && r["نام کالا"]?.includes("سنگین") &&
        r["نام کالا"]?.includes(k) && !r["نام کالا"]?.toLowerCase().includes("off")));
    });
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی") && r["نام کالا"]?.includes("سنگین") &&
      r["نام کالا"]?.toLowerCase().includes("off")));

    // --- بلوک ویژه تولوئن / متیلن --- 
    const newSpecialRows = reordered.filter(r => 
      r["نام کالا"]?.includes("تولوئن دی ایزو سیانات") || 
      r["نام کالا"]?.includes("متیلن دی فنیل دی ایزوسیانات")||
      r["نام کالا"]?.includes("متیلن دی فنیل ایزوسیانات")

    );

    if(newSpecialRows.length > 0){
      reordered.push({__HEADER__:true});
      const excludeColumns = ["کد عرضه", "تولید کننده کالا", "محل تحویل", "حداقل خرید"];
      const specialHeaders = headersArr.filter(h => !excludeColumns.includes(h));

      newSpecialRows.forEach(r=>{
        const rowData = specialHeaders.map(h => r[h] ?? "");
        const newRow = {};
        specialHeaders.forEach((h,i)=>{ newRow[h] = rowData[i]; });
        reordered.push(newRow);
      });
    }

    // --- تولید اکسل ---
    const workbook=new ExcelJS.Workbook();
    workbook.views=[{rightToLeft:true}];
    const worksheet=workbook.addWorksheet("نتیجه",{views:[{rightToLeft:true}]});

    const titleText="کارگزاری آینده نگر خوارزمی به مدیریت دکتر ذوقی 09123011311";
    const today=new Date(); let nextDay=new Date(today);
    let diffDays=today.getDay()===3?3:1;
    nextDay.setDate(today.getDate()+diffDays);
    const j=jalaali.toJalaali(nextDay);
    const jalaliDate=`${j.jy}/${String(j.jm).padStart(2,"0")}/${String(j.jd).padStart(2,"0")}`;
    const daysFa=["یکشنبه","دوشنبه","سه‌شنبه","چهارشنبه","پنجشنبه","جمعه","شنبه"];
    const dayFa=daysFa[nextDay.getDay()];
    const supplyText=`عرضه X روز ${dayFa} مورخ ${jalaliDate}`;

    const styleRow=(row,bgColor,bold=true)=>{
      row.eachCell(cell=>{
        if(bgColor) cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:bgColor}};
        cell.font={name:"B Nazanin",bold};
        cell.alignment={vertical:"middle",horizontal:"center"};
        applyAllBorders(cell);
      });
    };

    const addTitleRows=(ws)=>{
      const titleRow=ws.addRow([titleText]);
      titleRow.height=25;
      ws.mergeCells(titleRow.number,1,titleRow.number,headersArr.length);
      styleRow(titleRow,"FFFFA500");

      const supplyRow=ws.addRow([supplyText]);
      supplyRow.height=20;
      ws.mergeCells(supplyRow.number,1,supplyRow.number,headersArr.length);
      styleRow(supplyRow,"FFFFFF00");

      const headerRow=ws.addRow(headersArr);
      headerRow.height=20;
      styleRow(headerRow,"FFD3D3D3");
    };

    addTitleRows(worksheet);

    let rowCounter=0;
    reordered.forEach(row=>{
      if(row.__HEADER__) addTitleRows(worksheet);
      else if(Object.keys(row).length===0){
        const insertRow=worksheet.addRow([]);
        styleRow(insertRow,null,false);
      } else {
        const dataRow=worksheet.addRow(headersArr.map(h=>row[h]));
        styleRow(dataRow,null,false);
        if(rowCounter%2===0){
          dataRow.eachCell(cell=>{
            cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFF5DEB3"}};
          });
        }
        rowCounter++;
      }
    });

    const dateStr=new Date().toISOString().slice(0,10);
    const buffer=await workbook.xlsx.writeBuffer();
    const blob=new Blob([buffer],{type:"application/octet-stream"});
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a");
    a.href=url;
    a.download=`خروجی-${dateStr}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
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
      <button
        className="btn btn-primary"
        onClick={processData}
        disabled={data.length===0}
      >
        پردازش و دانلود خروجی
      </button>
    </div>
  );
}
