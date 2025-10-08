// .github/workflows/process-excel.js

// 引入必要的库
const xlsx = require('xlsx');
const fs = require('fs');

console.log("开始执行Excel处理脚本...");

try {
    // 定义Excel文件路径
    const excelFilePath = 'schedule.xlsx';

    // 检查文件是否存在
    if (!fs.existsSync(excelFilePath)) {
        throw new Error(`错误：在仓库根目录找不到 ${excelFilePath} 文件。`);
    }

    // 读取Excel文件
    const workbook = xlsx.readFile('../schedule.xlsx');
    console.log("成功读取 schedule.xlsx 文件。");

    // 获取第一个工作表的名称
    const sheetName = workbook.SheetNames[0];
    if (!sheetName) {
        throw new Error("错误：Excel文件中没有任何工作表(Sheet)。");
    }
    console.log(`正在处理工作表: ${sheetName}`);

    // 获取工作表对象
    const worksheet = workbook.Sheets[sheetName];

    // 将工作表内容转换为JSON对象数组
    // 'sheet_to_json' 会自动将第一行作为对象的键（key）
    const jsonData = xlsx.utils.sheet_to_json(worksheet);
    console.log(`成功从Excel转换了 ${jsonData.length} 行数据。`);

    // 将JSON数据写入 data.json 文件
    // 使用 null, 2 参数格式化JSON，使其更易读（可选）
    fs.writeFileSync('data.json', JSON.stringify(jsonData, null, 2));
    console.log("成功生成 data.json 文件！");
    
    // 脚本成功结束
    process.exit(0);

} catch (error) {
    // 如果过程中出现任何错误，打印错误信息并以失败状态退出
    console.error("脚本执行失败:", error.message);
    process.exit(1);
}
