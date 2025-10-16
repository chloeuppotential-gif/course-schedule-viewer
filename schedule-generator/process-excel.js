const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

function processCourseData(rawData) {
    try {
        console.log('开始处理课程数据...');
        console.log('原始数据示例（前2行）:', rawData.slice(0, 2)); // 打印前两行数据用于调试

        // ✅ 精确映射你的Excel表头
        rawData = rawData.map(row => {
            const mappedRow = {
                // 教师字段映射
                teacher: row['任课教师'] || row['教师'] || row['老师'] || row['Teacher'] || row.teacher || '',
                
                // 主题字段映射
                topic: row['教学主题'] || row['主题'] || row['课程主题'] || row['课程名称'] || row['Topic'] || row.topic || '',
                
                // 课时字段映射
                session: row['课时'] || row['节次'] || row['课节'] || row['Session'] || row.session || '',
                
                // 日期字段映射 - 根据你的Excel结构
                courseStart: row['开课日期'] || row['课程开始'] || row['courseStart'] || row.courseStart || '',
                courseEnd: row['结课日期'] || row['课程结束'] || row['courseEnd'] || row.courseEnd || '',
                sessionStart: row['起始日期'] || row['课时开始'] || row['sessionStart'] || row.sessionStart || '',
                sessionEnd: row['结束日期'] || row['课时结束'] || row['sessionEnd'] || row.sessionEnd || '',
                
                // 学生字段映射（如果有的话）
                student: row['学生'] || row['姓名'] || row['学员'] || row['Student'] || row.student || ''
            };
            
            // 数据清理：移除时间部分，只保留日期
            if (mappedRow.courseStart && mappedRow.courseStart.includes(' ')) {
                mappedRow.courseStart = mappedRow.courseStart.split(' ')[0];
            }
            if (mappedRow.courseEnd && mappedRow.courseEnd.includes(' ')) {
                mappedRow.courseEnd = mappedRow.courseEnd.split(' ')[0];
            }
            if (mappedRow.sessionStart && mappedRow.sessionStart.includes(' ')) {
                mappedRow.sessionStart = mappedRow.sessionStart.split(' ')[0];
            }
            if (mappedRow.sessionEnd && mappedRow.sessionEnd.includes(' ')) {
                mappedRow.sessionEnd = mappedRow.sessionEnd.split(' ')[0];
            }
            
            console.log('映射后的行数据:', mappedRow); // 调试输出每一行映射结果
            return mappedRow;
        });

        // 过滤掉空行（所有主要字段都为空的行）
        const filteredData = rawData.filter(row => {
            const hasData = row.teacher || row.topic || row.session || row.courseStart || row.courseEnd;
            if (!hasData) {
                console.log('过滤掉空行:', row);
            }
            return hasData;
        });

        const processedData = filteredData.map((row, index) => ({
            id: index + 1,
            ...row,
            processed: true
        }));

        console.log(`成功处理 ${processedData.length} 条课程数据`);
        console.log('处理后的数据示例（前3条）:', processedData.slice(0, 3));
        return processedData;
    } catch (error) {
        console.error('处理课程数据时出错:', error);
        return rawData;
    }
}

// 处理学生数据（从Sheet2提取）
function extractStudentData(workbook) {
    try {
        console.log('开始提取学生数据...');
        
        const sheetNames = workbook.SheetNames;
        console.log('所有工作表:', sheetNames);
        
        // 查找学生相关的工作表
        let studentSheet = null;
        if (sheetNames.includes('Sheet2')) {
            studentSheet = workbook.Sheets['Sheet2'];
            console.log('找到Sheet2，尝试提取学生数据');
        }
        
        if (!studentSheet) {
            console.log('未找到学生数据工作表，返回空数组');
            return [];
        }
        
        const studentRawData = XLSX.utils.sheet_to_json(studentSheet);
        console.log('Sheet2原始数据:', studentRawData);
        
        if (studentRawData.length === 0) {
            console.log('Sheet2没有数据');
            return [];
        }
        
        // 映射学生数据字段
        const studySessions = studentRawData.map((row, index) => ({
            id: `study_${index}`,
            studentName: row['受课同学'] || row['学生'] || row['姓名'] || '同学',
            topic: row['学习课程'] || row['课程'] || row['主题'] || '',
            session: row['学习课时'] || row['课时'] || '',
            startTime: (row['开始时间'] || '').split(' ')[0], // 只取日期部分
            endTime: (row['结束时间'] || '').split(' ')[0],   // 只取日期部分
            duration: 60, // 默认60分钟
            completed: false,
            notes: ''
        }));
        
        // 按学生分组
        const studentGroups = {};
        studySessions.forEach(session => {
            const studentName = session.studentName;
            if (!studentGroups[studentName]) {
                studentGroups[studentName] = {
                    id: `student_${studentName}`,
                    name: studentName,
                    studySessions: []
                };
            }
            studentGroups[studentName].studySessions.push(session);
        });
        
        const students = Object.values(studentGroups);
        console.log(`成功提取 ${students.length} 名学生的数据`);
        return students;
        
    } catch (error) {
        console.error('提取学生数据时出错:', error);
        return [];
    }
}

function processExcel() {
    try {
        console.log('开始处理Excel文件...');
        console.log('当前工作目录:', process.cwd());
        console.log('目录内容:', fs.readdirSync(process.cwd()).join(', '));
        
        // 查找Excel文件
        let workbook;
        let excelPath = '';
        
        const possiblePaths = [
            'schedule.xlsx',                           
            path.join(__dirname, 'schedule.xlsx'),     
            '../schedule.xlsx',                         
            path.resolve(process.cwd(), 'schedule.xlsx') 
        ];
        
        let fileFound = false;
        for (const p of possiblePaths) {
            console.log(`尝试路径: ${p}`);
            if (fs.existsSync(p)) {
                excelPath = p;
                console.log(`✅ 找到Excel文件: ${excelPath}`);
                fileFound = true;
                break;
            }
        }
        
        if (!fileFound) {
            throw new Error('❌ 无法找到schedule.xlsx文件，所有尝试路径均失败');
        }
        
        // 读取Excel文件
        workbook = XLSX.readFile(excelPath);
        
        // 获取所有工作表
        const sheetNames = workbook.SheetNames;
        console.log(`📋 工作表列表: ${sheetNames.join(', ')}`);
        
        // 处理第一个工作表（教师课程数据）
        const firstSheet = workbook.Sheets[sheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet);
        console.log(`📊 Sheet1原始数据数量: ${jsonData.length} 条记录`);
        
        if (jsonData.length === 0) {
            console.warn('⚠️ Sheet1没有数据');
        } else {
            console.log('📝 Sheet1第一行数据:', jsonData[0]);
        }
        
        // 处理课程数据
        const processedData = processCourseData(jsonData);
        
        // 提取学生数据
        const students = extractStudentData(workbook);
        
        // 写入文件
        console.log('💾 开始写入JSON文件...');
        
        fs.writeFileSync('data.json', JSON.stringify(processedData, null, 2));
        console.log('✅ 成功生成 data.json');
        
        fs.writeFileSync('students.json', JSON.stringify(students, null, 2));
        console.log('✅ 成功生成 students.json');
        
        // 输出最终统计
        console.log('\n📈 处理完成统计:');
        console.log(`- 教师课程记录: ${processedData.length} 条`);
        console.log(`- 学生数据: ${students.length} 人`);
        console.log(`- 学习记录总数: ${students.reduce((sum, s) => sum + (s.studySessions ? s.studySessions.length : 0), 0)} 条`);
        
    } catch (error) {
        console.error('❌ 处理Excel时发生错误:', error);
        process.exit(1);
    }
}

// 执行主函数
processExcel();
