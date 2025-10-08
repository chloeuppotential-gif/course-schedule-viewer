const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path'); // 添加缺失的path模块导入

// 定义文件路径
const excelFilePath = path.join(__dirname, '..', 'schedule.xlsx');
const outputFilePath = path.join(__dirname, 'data.json');

try {
  console.log("开始处理Excel文件...");
  console.log("Excel文件路径:", excelFilePath);
  
  // 检查Excel文件是否存在
  if (!fs.existsSync(excelFilePath)) {
    console.error(`错误：找不到Excel文件: ${excelFilePath}`);
    process.exit(1);
  }

  // 读取Excel文件
  const workbook = xlsx.readFile(excelFilePath, { cellDates: true });
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  
  // 转换为JSON
  const jsonData = xlsx.utils.sheet_to_json(worksheet);
  console.log(`从Excel中读取了${jsonData.length}行数据`);
  
  // 数据处理
  const processedTasks = [];
  let currentTeacher = '', currentTopic = '', currentCourseStart = null, currentCourseEnd = null;
  
  const formatDate = (date) => {
    if (!date) return null;
    
    try {
      const d = new Date(date);
      if (isNaN(d.getTime())) return null;
      
      const year = d.getFullYear();
      const month = String(d.getMonth() + 1).padStart(2, '0');
      const day = String(d.getDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    } catch (e) {
      console.error("日期转换错误:", e);
      return null;
    }
  };
  
  jsonData.forEach((row, index) => {
    // 跳过空行
    if (Object.values(row).every(val => val === null || val === undefined || String(val).trim() === '')) {
      return;
    }
    
    // 处理任课教师
    if (row['任课教师'] && String(row['任课教师']).trim() !== '') {
      currentTeacher = String(row['任课教师']).trim();
    }
    
    // 处理教学主题
    if (row['教学主题'] && String(row['教学主题']).trim() !== '') {
      currentTopic = String(row['教学主题']).trim();
    }
    
    // 处理课程起止日期
    const rowCourseStart = formatDate(row['开课日期']);
    const rowCourseEnd = formatDate(row['结课日期']);
    if (rowCourseStart && rowCourseEnd) {
      currentCourseStart = rowCourseStart;
      currentCourseEnd = rowCourseEnd;
    }
    
    // 处理课时起止日期
    const sessionStart = formatDate(row['起始日期']);
    const sessionEnd = formatDate(row['结束日期']);
    
    // 创建任务对象
    if (currentTeacher && currentTopic && sessionStart && sessionEnd) {
      const task = {
        teacher: currentTeacher,
        topic: currentTopic,
        session: row['课时'] ? String(row['课时']).trim() : 'N/A',
        courseStart: currentCourseStart,
        courseEnd: currentCourseEnd,
        sessionStart: sessionStart,
        sessionEnd: sessionEnd
      };
      
      processedTasks.push(task);
    }
  });
  
  // 写入JSON文件
  fs.writeFileSync(outputFilePath, JSON.stringify(processedTasks, null, 2));
  console.log(`成功生成data.json，包含${processedTasks.length}条记录`);

} catch (error) {
  console.error('处理Excel文件时出错:', error);
  
  // 创建一个空的数据文件
  fs.writeFileSync(outputFilePath, JSON.stringify([], null, 2));
  process.exit(1);
}
