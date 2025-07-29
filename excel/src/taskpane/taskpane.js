/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// Office.onReady((info) => {
//   if (info.host === Office.HostType.Excel) {
//     document.getElementById("sideload-msg").style.display = "none";
//     document.getElementById("app-body").style.display = "flex";
//     document.getElementById("run").onclick = run;
//   }
// });

// export async function run() {
//   try {
//     await Excel.run(async (context) => {
//       const range = context.workbook.getSelectedRange();

//       // Read the range address.
//       range.load("address");

//       // Update the fill color.
//       range.format.fill.color = "yellow";

//       await context.sync();
//       console.log(`The range address was ${range.address}.`);
//     });
//   } catch (error) {
//     console.error(error);
//   }
// }

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("btnLoad").onclick = loadHeaders;
    document.getElementById("btnTranspose").onclick = transposeData;
  }
});

async function loadHeaders() {
  await Excel.run(async context => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const headerRow = parseInt(document.getElementById("headerRow").value) || 1;
    const range = sheet.getRange(`1:1`);
    range.load("values");
    await context.sync();

    const headerValues = range.values[headerRow - 1];
    const list = document.getElementById("columnList");
    list.innerHTML = "";

    headerValues.forEach((val, idx) => {
      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.id = `col${idx}`;
      checkbox.value = idx;
      list.appendChild(checkbox);

      const label = document.createElement("label");
      label.htmlFor = `col${idx}`;
      label.innerText = ` ${val || "(空)"} `;
      list.appendChild(label);
      list.appendChild(document.createElement("br"));
    });
  });
}

async function transposeData() {
  await Excel.run(async context => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const headerRow = parseInt(document.getElementById("headerRow").value) || 1;
    const range = sheet.getUsedRange();
    range.load("values, columnCount, rowCount");
    await context.sync();

    const values = range.values;
    const header = values[headerRow - 1];
    const data = values.slice(headerRow);

    const selectedCols = [];
    header.forEach((val, idx) => {
      const checkbox = document.getElementById(`col${idx}`);
      if (checkbox && checkbox.checked) selectedCols.push(idx);
    });

    const fixedCols = header.map((_, i) => i).filter(i => !selectedCols.includes(i));

    const outputColNames = document.getElementById("outputCols").value.split(",").map(x => x.trim());
    const blockSize = outputColNames.length;

    if (selectedCols.length % blockSize !== 0) {
      alert("所选列数量必须能整除输出列数");
      return;
    }

    const transposedRows = [];

    data.forEach(row => {
      const fixed = fixedCols.map(i => row[i]);
      const blockCount = selectedCols.length / blockSize;
      for (let b = 0; b < blockCount; b++) {
        const line = [...fixed];
        line.push(b + 1);
        for (let i = 0; i < blockSize; i++) {
          line.push(row[selectedCols[b * blockSize + i]]);
        }
        transposedRows.push(line);
      }
    });

    // 创建新工作表并写入数据
    const newSheet = context.workbook.worksheets.add("转置结果");
    const headerRowOut = fixedCols.map(i => header[i]).concat([document.getElementById("indexCol").value], outputColNames);
    newSheet.getRangeByIndexes(0, 0, transposedRows.length + 1, headerRowOut.length).values = [headerRowOut, ...transposedRows];
    newSheet.activate();

    await context.sync();
    alert("转置完成！");
  });
}

