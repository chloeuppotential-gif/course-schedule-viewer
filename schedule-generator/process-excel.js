const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

console.log('开始执行【最终版】Excel处理脚本...');

try {
    // 明确指定输入和输出文件的路径
    const excelInputPath = '../schedule.xlsx';
    const jsonOutputPath = './data.json';

    console.log(`正在尝试读取文件: ${excelInputPath}`);
    const workbook = xlsx.readFile(excelInputPath);

    const sheetName = workbook.SheetNames[0];
    if (!sheetName) {
        throw new Error('Excel文件中没有任何工作表(Sheet)。');
    }
    console.log(`成功读取工作表: ${sheetName}`);

    const worksheet = workbook.Sheets[sheetName];
    const jsonData = xlsx.utils.sheet_to_json(worksheet);

    console.log(`成功将Excel转换为JSON，共 ${jsonData.length} 行数据。`);

    fs.writeFileSync(jsonOutputPath, JSON.stringify(jsonData, null, 2));
    console.log(`成功将数据写入到文件: ${jsonOutputPath}`);

} catch (error) {
    console.error('脚本执行过程中发生严重错误:', error.message);
    process.exit(1); // 以失败状态退出
}
