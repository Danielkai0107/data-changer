import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

import './App.scss';

function App() {
  const [jsonData, setJsonData] = useState('');
  const [excelData, setExcelData] = useState(null);

  const flattenJson = (data) => {
  return data.map(item => {
    const variants = item.variants;
    delete item.variants;
    variants.forEach((variant, index) => {
      Object.keys(variant).forEach(key => {
        item[`variant_${index}_${key}`] = variant[key];
      });
    });
    return item;
  });
};

const unflattenJson = (data) => {
  return data.map(item => {
    const keys = Object.keys(item);
    const variants = [];
    keys.forEach(key => {
      if (key.startsWith('variant_')) {
        const [ , index, variantKey] = key.split('_');
        if (!variants[index]) variants[index] = {};
        variants[index][variantKey] = item[key];
        delete item[key];
      }
    });
    item.variants = variants;
    return item;
  });
};

  const handleJsonChange = (event) => {
    setJsonData(event.target.value);
  };

  const handleJsonToExcel = () => {try {
      const jsonObject = JSON.parse(jsonData);
      const flattenedJsonObject = flattenJson(jsonObject);
      const ws = XLSX.utils.json_to_sheet(flattenedJsonObject);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
      const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });

      const buf = new ArrayBuffer(wbout.length);
      const view = new Uint8Array(buf);
      
      for (let i=0; i<wbout.length; i++) view[i] = wbout.charCodeAt(i) & 0xFF;

      saveAs(new Blob([buf], {type: 'application/octet-stream'}), 'data.xlsx');
    } catch(e) {
      console.error("Error", e);
      alert("轉換過程中出現錯誤。請檢查你的JSON數據。");
    }
  };

  const handleExcelChange = (event) => {
    const file = event.target.files[0];
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      setExcelData(workbook);
    };
    reader.readAsArrayBuffer(file);
  };

  const handleExcelToJson = () => {
    try {
      const wsname = excelData.SheetNames[0];
      const ws = excelData.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws);
      const unflattenedData = unflattenJson(data);
      setJsonData(JSON.stringify(unflattenedData, null, 2));
    } catch(e) {
      console.error("Error", e);
      alert("轉換過程中出現錯誤。請檢查你的Excel文件。");
    }
  };

  const handleCopyJson = async () => {
  try {
    await navigator.clipboard.writeText(jsonData);
    alert("JSON數據已複製到剪貼簿！");
  } catch (err) {
    console.error('Failed to copy text: ', err);
  }
};


  return (
    <div className="app">
      <div className="json-input">
        <textarea value={jsonData} onChange={handleJsonChange} />
        <button onClick={handleJsonToExcel}>轉換成 Excel</button>
        <button onClick={handleCopyJson}>複製 JSON</button>
      </div>
      <div className="excel-input">
        <input type="file" onChange={handleExcelChange} />
        <button onClick={handleExcelToJson}>轉換成 JSON</button>
      </div>
    </div>
  );
}

export default App;
