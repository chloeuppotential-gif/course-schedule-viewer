const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

function processCourseData(rawData) {
    try {
        console.log('开始处理课程数据...');
        console.log('原始数据示例（前2行）:', rawData.slice(0, 2));

        // 精确映射Excel表头到需要的字段
        const processedData = rawData.map((row, index) => {
            const mappedRow = {
                id: index + 1,
                
                // 教师字段映射
                teacher: row['任课教师'] || row['教师'] || row['老师'] || '',
                
                // 主题字段映射
                topic: row['教学主题'] || row['主题'] || row['课程主题'] || '',
                
                // 课时字段映射
                session: row['课时'] || row['节次'] || '',
                
                // ✅ 关键修复：日期字段映射
                courseStart: cleanDate(row['开课日期'] || row['课程开始'] || ''),
                courseEnd: cleanDate(row['结课日期'] || row['课程结束'] || ''),
                sessionStart: cleanDate(row['起始日期'] || row['课时开始'] || ''),
                sessionEnd: cleanDate(row['结束日期'] || row['课时结束'] || ''),
                
                processed: true
            };
            
            console.log(`第${index + 1}行映射结果:`, mappedRow);
            return mappedRow;
        });

        // 过滤掉空行
        const filteredData = processedData.filter(row => {
            const hasData = row.teacher || row.topic || row.session || row.courseStart;
            return hasData;
        });

        console.log(`✅ 成功处理 ${filteredData.length} 条课程数据`);
        return filteredData;
    } catch (error) {
        console.error('❌ 处理课程数据时出错:', error);
        return [];
    }
}

// 清理日期格式的辅助函数
function cleanDate(dateStr) {
    if (!dateStr) return '';
    
    // 移除时间部分，只保留日期 (例: "2025-10-15 00:00:00" -> "2025-10-15")
    if (typeof dateStr === 'string' && dateStr.includes(' ')) {
        return dateStr.split(' ')[0];
    }
    
    // 如果是Excel日期数字，转换为字符串
    if (typeof dateStr === 'number') {
        const date = XLSX.SSF.parse_date_code(dateStr);
        return `${date.y}-${String(date.m).padStart(2, '0')}-${String(date.d).padStart(2, '0')}`;
    }
    
    return String(dateStr);
}

// 处理学生数据
function extractStudentData(workbook) {
    try {
        console.log('开始提取学生数据...');
        
        const sheetNames = workbook.SheetNames;
        console.log('所有工作表:', sheetNames);
        
        if (!sheetNames.includes('Sheet2')) {
            console.log('❌ 未找到Sheet2，返回空数组');
            return [];
        }
        
        const studentSheet = workbook.Sheets['Sheet2'];
        const studentRawData = XLSX.utils.sheet_to_json(studentSheet);
        console.log('Sheet2原始数据:', studentRawData);
        
        if (studentRawData.length === 0) {
            console.log('⚠️ Sheet2没有数据');
            return [];
        }
        
        // 映射学生数据字段
        const studySessions = studentRawData.map((row, index) => ({
            id: `study_${index}`,
            studentName: row['受课同学'] || row['学生'] || '同学',
            topic: row['学习课程'] || row['课程'] || '',
            session: row['学习课时'] || row['课时'] || '',
            startTime: cleanDate(row['开始时间'] || ''),
            endTime: cleanDate(row['结束时间'] || ''),
            duration: 60,
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
        console.log(`✅ 成功提取 ${students.length} 名学生的数据`);
        return students;
        
    } catch (error) {
        console.error('❌ 提取学生数据时出错:', error);
        return [];
    }
}

function processExcel() {
    try {
        console.log('🚀 开始处理Excel文件...');
        console.log('当前工作目录:', process.cwd());
        
        // 查找Excel文件
        const possiblePaths = [
            'schedule.xlsx',                           
            path.join(__dirname, 'schedule.xlsx'),     
            '../schedule.xlsx',                         
            path.resolve(process.cwd(), 'schedule.xlsx') 
        ];
        
        let excelPath = '';
        let fileFound = false;
        
        for (const p of possiblePaths) {
            console.log(`🔍 尝试路径: ${p}`);
            if (fs.existsSync(p)) {
                excelPath = p;
                console.log(`✅ 找到Excel文件: ${excelPath}`);
                fileFound = true;
                break;
            }
        }
        
        if (!fileFound) {
            throw new Error('❌ 无法找到schedule.xlsx文件');
        }
        
        // 读取Excel文件
        const workbook = XLSX.readFile(excelPath);
        const sheetNames = workbook.SheetNames;
        console.log(`📋 工作表列表: ${sheetNames.join(', ')}`);
        
        // 处理第一个工作表
        const firstSheet = workbook.Sheets[sheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet);
        console.log(`📊 Sheet1数据量: ${jsonData.length} 条记录`);
        
        // 处理数据
        const processedData = processCourseData(jsonData);
        const students = extractStudentData(workbook);
        
        // 写入文件
        console.log('💾 写入JSON文件...');
        
        fs.writeFileSync('data.json', JSON.stringify(processedData, null, 2));
        console.log('✅ 成功生成 data.json');
        
        fs.writeFileSync('students.json', JSON.stringify(students, null, 2));
        console.log('✅ 成功生成 students.json');
        
        // 最终统计
        console.log('\n📈 处理完成统计:');
        console.log(`- 教师课程记录: ${processedData.length} 条`);
        console.log(`- 学生数据: ${students.length} 人`);
        console.log(`- 学习记录总数: ${students.reduce((sum, s) => sum + (s.studySessions?.length || 0), 0)} 条`);
        
    } catch (error) {
        console.error('❌ 处理Excel时发生错误:', error);
        process.exit(1);
    }
}

// 执行主函数
processExcel();
