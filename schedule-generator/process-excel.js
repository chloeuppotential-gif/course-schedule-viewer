const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// 定义输入和输出文件的路径
// GitHub Actions运行时，__dirname是当前脚本所在目录 (schedule-generator)
// 我们需要用 '..' 返回上一级，才能找到 schedule.xlsx 和 public 目录
const excelFilePath = path.join(__dirname, '..', 'schedule.xlsx');
const outputJsonPath = path.join(__dirname, '..', 'public', 'data.json');

console.log(`正在读取Excel文件: ${excelFilePath}`);

// --- 核心数据处理逻辑 ---
function processExcelData(data) {
    const generatedTasks = [];
    let sessionCounter = {};

    data.forEach(rule => {
        if (!rule['教学主题'] || !rule['课程开始日期']) return;

        const topic = String(rule['教学主题']).trim();
        const teacher = String(rule['任课教师']).trim();
        const repetitionType = String(rule['重复类型']).trim();
        const repetitionDetail = rule['重复细节'] ? String(rule['重复细节']) : '';
        const totalSessions = parseInt(rule['总课时'], 10);
        const sessionDuration = parseInt(rule['单次课时长(天)'], 10) || 1;
        
        // Excel中的日期可能会被解析为数字，需要特殊处理确保是Date对象
        const courseStartDate = new Date(rule['课程开始日期']);
        const courseEndDate = rule['课程结束日期'] ? new Date(rule['课程结束日期']) : new Date('2099-12-31');

        sessionCounter[topic] = 0;
        let currentDate = new Date(courseStartDate);

        // 循环直到课程结束日期或达到总课时数
        while (currentDate <= courseEndDate && (!totalSessions || sessionCounter[topic] < totalSessions)) {
            let isValidDay = false;
            switch (repetitionType) {
                case '每周': {
                    const weekDays = repetitionDetail.split(',').map(d => parseInt(d.trim(), 10));
                    // getDay() 返回 0(周日)-6(周六)，我们调整为 1(周一)-7(周日)
                    const currentDayOfWeek = currentDate.getDay() === 0 ? 7 : currentDate.getDay();
                    if (weekDays.includes(currentDayOfWeek)) isValidDay = true;
                    break;
                }
                case '每两周': {
                    const weekDays = repetitionDetail.split(',').map(d => parseInt(d.trim(), 10));
                    const currentDayOfWeek = currentDate.getDay() === 0 ? 7 : currentDate.getDay();
                    
                    // 计算从开始日期到现在过了多少周
                    const timeDiff = currentDate.getTime() - courseStartDate.getTime();
                    const weekDiff = Math.floor(timeDiff / (1000 * 60 * 60 * 24 * 7));

                    // 如果周数差是偶数（0, 2, 4...），并且星期也对，则有效
                    if (weekDiff % 2 === 0 && weekDays.includes(currentDayOfWeek)) {
                        isValidDay = true;
                    }
                    break;
                }
                case '每月': {
                    const monthDays = repetitionDetail.split(',').map(d => parseInt(d.trim(), 10));
                    if (monthDays.includes(currentDate.getDate())) isValidDay = true;
                    break;
                }
                case '单次': {
                    // 对于单次，只在课程开始那天生成
                    if (currentDate.getTime() === courseStartDate.getTime()) isValidDay = true;
                    break;
                }
            }

            if (isValidDay) {
                sessionCounter[topic]++;
                const sessionStart = new Date(currentDate);
                const sessionEnd = new Date(sessionStart);
                sessionEnd.setDate(sessionEnd.getDate() + sessionDuration - 1);
                generatedTasks.push({
                    teacher: teacher, topic: topic, session: `课时 ${sessionCounter[topic]}`,
                    courseStart: formatDate(courseStartDate), courseEnd: formatDate(courseEndDate),
                    sessionStart: formatDate(sessionStart), sessionEnd: formatDate(sessionEnd)
                });
            }
            // 对于'单次'类型，处理完一次就跳出循环
            if (repetitionType === '单次' && isValidDay) {
                break;
            }
            
            currentDate.setDate(currentDate.getDate() + 1);
        }
    });
    return generatedTasks;
}

function formatDate(date) {
    if (!date || isNaN(date.getTime())) return null;
    // 确保我们处理的是UTC日期，避免时区问题
    const d = new Date(date);
    const year = d.getUTCFullYear();
    const month = String(d.getUTCMonth() + 1).padStart(2, '0');
    const day = String(d.getUTCDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}
// --- 主执行逻辑 ---
try {
    // 读取Excel文件
    const workbook = xlsx.readFile(excelFilePath);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    // 使用 { cellDates: true } 确保日期列被正确解析为Date对象
    const jsonData = xlsx.utils.sheet_to_json(worksheet, { cellDates: true });

    console.log("Excel数据读取成功，开始处理...");
    const tasks = processExcelData(jsonData);

    // 将处理好的数据写入到 public/data.json
    fs.writeFileSync(outputJsonPath, JSON.stringify(tasks, null, 2));
    console.log(`成功生成 data.json 文件，路径: ${outputJsonPath}。包含 ${tasks.length} 条课程记录。`);

} catch (error) {
    console.error("生成数据时出错:", error);
    process.exit(1); // 退出并报告错误，这样GitHub Actions会知道任务失败了
}
