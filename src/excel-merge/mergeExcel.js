const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 定义要提取的列名
const columnsToExtract = ['手机', '电话', '年级', '提交时间'];

// 定义文件夹路径
const folderPath = '金数据-提取和合并';

// 读取文件夹中的所有文件
fs.readdir(folderPath, (err, files) => {
  if (err) {
    console.error('读取文件夹时出错:', err);
    return;
  }

  // 过滤出所有的 Excel 文件
  const excelFiles = files.filter(file => ['.xlsx', '.xls'].includes(path.extname(file)));

  // 存储所有数据的数组
  const allData = [];

  // 遍历每个 Excel 文件
  excelFiles.forEach(file => {
    const filePath = path.join(folderPath, file);
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // 获取表头
    const headers = XLSX.utils.sheet_to_json(worksheet, { header: 1 })[0];

    // 找到要提取的列的索引
    const columnIndices = columnsToExtract.map(column => {
      let ind = -1
      headers.find(h => {
        if (h.includes(column)) {
          ind = headers.indexOf(h)
        }
      })
      return ind
    });

    // 读取数据
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 }).slice(1);

    // 提取特定列的数据
    const extractedData = data.map(row => {
      const newRow = columnIndices.map(index => row[index]);
      // 提取文件名中的特定部分
      let fileNamePart = file.split('_');
      if (fileNamePart.length > 1) {
        if (fileNamePart[fileNamePart.length - 1][0] === '2') {
          fileNamePart = fileNamePart.slice(0, fileNamePart.length - 1);
        }
      } else {
        fileNamePart = [ fileNamePart[0].split('.')[0] ];
      }
      newRow.push(fileNamePart.join('_').split('.')[0]);
      return newRow;
    });

    // 将提取的数据添加到所有数据数组中
    allData.push(...extractedData);
  });

  // 创建新的工作簿和工作表
  const newWorkbook = XLSX.utils.book_new();
  const newWorksheet = XLSX.utils.aoa_to_sheet([[...columnsToExtract, '文件名']].concat(allData));

  // 将工作表添加到工作簿中
  XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, '合并数据');

  // 保存合并后的文件
  XLSX.writeFile(newWorkbook, '合并后的文件.xlsx');

  console.log('合并完成，结果已保存到 合并后的文件.xlsx');
});
