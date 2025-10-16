// 文件名: process-excel.js
const xlsx = require('xlsx');
const fs = require('fs');

// 读取Excel文件
function processExcel() {
    console.log('开始处理Excel文件...');
    const workbook = xlsx.readFile('schedule.xlsx');
    
    // 获取Sheet名称
    const sheetNames = workbook.SheetNames;
    console.log(`Excel包含 ${sheetNames.length} 个工作表:`, sheetNames);
    
    // 处理教师课程数据 (Sheet1)
    const teacherSheet = workbook.Sheets[sheetNames[0]];
    const teacherData = xlsx.utils.sheet_to_json(teacherSheet);
    console.log(`从 ${sheetNames[0]} 读取了 ${teacherData.length} 条记录`);
    
    // 处理学生学习数据 (Sheet2), 如果存在
    let studentData = [];
    if (sheetNames.length > 1) {
        const studentSheet = workbook.Sheets[sheetNames[1]];
        studentData = xlsx.utils.sheet_to_json(studentSheet);
        console.log(`从 ${sheetNames[1]} 读取了 ${studentData.length} 条记录`);
    }
    
    // 转换教师数据
    const courses = teacherData.filter(row => row['任课教师'] && row['教学主题']).map(row => ({
        teacher: row['任课教师'],
        topic: row['教学主题'],
        session: row['课时'],
        courseStart: formatDate(row['开课日期']),
        sessionStart: formatDate(row['起始日期']),
        sessionEnd: formatDate(row['结束日期']),
        courseEnd: formatDate(row['结课日期'])
    }));
    
    // 处理学生数据
    const students = [];
    if (studentData.length > 0) {
        // 创建一个学生记录
        const student = {
            id: "student_同学",
            name: "同学",
            studySessions: []
        };
        
        // 添加学习记录
        studentData.forEach((row, index) => {
            if (!row['受课同学'] || !row['学习课程']) return;
            
            const studySession = {
                id: `study_${index}`,
                topic: row['学习课程'],
                session: row['学习课时'],
                startTime: formatDate(row['开始时间']),
                endTime: formatDate(row['结束时间']),
                // 添加默认值，您可以根据需要修改
                duration: 0,
                completed: false,
                notes: ''
            };
            
            // 计算时长（如果有日期）
            if (studySession.startTime && studySession.endTime) {
                const startDate = new Date(studySession.startTime);
                const endDate = new Date(studySession.endTime);
                // 计算天数差
                const diffTime = Math.abs(endDate - startDate);
                const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
                studySession.duration = diffDays * 60; // 假设每天学习60分钟
            }
            
            student.studySessions.push(studySession);
        });
        
        students.push(student);
    }
    
    // 写入JSON文件
    fs.writeFileSync('data.json', JSON.stringify(courses, null, 2));
    console.log(`已生成 data.json 文件，包含 ${courses.length} 个教师课程记录`);
    
    fs.writeFileSync('students.json', JSON.stringify(students, null, 2));
    console.log(`已生成 students.json 文件，包含 ${students.length} 个学生记录`);
    
    console.log('处理完成!');
}

// 格式化日期
function formatDate(dateValue) {
    if (!dateValue) return '';
    
    let date;
    if (typeof dateValue === 'string') {
        // 移除时间部分如果存在
        const dateStr = dateValue.split(' ')[0];
        date = new Date(dateStr);
    } else {
        date = new Date(dateValue);
    }
    
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    
    return `${year}-${month}-${day}`;
}

// 执行处理
processExcel();
