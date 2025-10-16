const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 清理日期格式的辅助函数
function cleanDate(dateStr) {
    if (!dateStr) return '';
    
    // 移除时间部分，只保留日期
    if (typeof dateStr === 'string' && dateStr.includes(' ')) {
        return dateStr.split(' ')[0];
    }
    
    // 如果是Excel日期数字，转换为字符串
    if (typeof dateStr === 'number') {
        try {
            const date = XLSX.SSF.parse_date_code(dateStr);
            return `${date.y}-${String(date.m).padStart(2, '0')}-${String(date.d).padStart(2, '0')}`;
        } catch (e) {
            return '';
        }
    }
    
    return String(dateStr);
}

// 处理课程数据 - 强制映射到英文字段名
function processCourseData(rawData) {
    try {
        console.log('🔄 开始处理课程数据...');
        console.log('📊 原始数据总数:', rawData.length);
        console.log('📋 原始数据示例（前2行）:', JSON.stringify(rawData.slice(0, 2), null, 2));

        const processedData = rawData.map((row, index) => {
            // 🔧 强制映射：中文字段 → 英文字段
            const mappedRow = {
                id: index + 1,
                
                // 教师信息映射
                teacher: row['任课教师'] || row['教师'] || row['老师'] || '',
                
                // 主题信息映射
                topic: row['教学主题'] || row['主题'] || row['课程主题'] || '',
                
                // 课时信息映射
                session: row['课时'] || row['节次'] || '',
                
                // 🎯 关键：日期字段强制映射为英文
                courseStart: cleanDate(row['开课日期'] || ''),
                courseEnd: cleanDate(row['结课日期'] || ''),
                sessionStart: cleanDate(row['起始日期'] || ''),
                sessionEnd: cleanDate(row['结束日期'] || ''),
                
                processed: true
            };
            
            console.log(`✅ 第${index + 1}行映射完成:`, {
                teacher: mappedRow.teacher,
                topic: mappedRow.topic,
                courseStart: mappedRow.courseStart,
                courseEnd: mappedRow.courseEnd
            });
            
            return mappedRow;
        });

        // 过滤掉完全空的行
        const filteredData = processedData.filter(row => {
            const hasValidData = row.teacher || row.topic || row.session || row.courseStart;
            return hasValidData;
        });

        console.log(`🎉 课程数据处理完成: ${filteredData.length} 条有效记录`);
        console.log('📝 处理后的数据示例:', JSON.stringify(filteredData[0], null, 2));
        
        return filteredData;
        
    } catch (error) {
        console.error('❌ 处理课程数据时出错:', error);
        return [];
    }
}

// 处理学生数据
function extractStudentData(workbook) {
    try {
        console.log('👥 开始提取学生数据...');
        
        const sheetNames = workbook.SheetNames;
        console.log('📄 所有工作表:', sheetNames);
        
        if (!sheetNames.includes('Sheet2')) {
            console.log('⚠️ 未找到Sheet2，跳过学生数据');
            return [];
        }
        
        const studentSheet = workbook.Sheets['Sheet2'];
        const studentRawData = XLSX.utils.sheet_to_json(studentSheet);
        console.log('🔍 Sheet2原始数据量:', studentRawData.length);
        
        if (studentRawData.length === 0) {
            console.log('📭 Sheet2为空');
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
        console.log(`👨‍🎓 成功提取 ${students.length} 名学生的数据`);
        return students;
        
    } catch (error) {
        console.error('❌ 提取学生数据时出错:', error);
        return [];
    }
}

// 主处理函数
function processExcel() {
    try {
        console.log('🚀 开始处理Excel文件...');
        console.log('📂 当前工作目录:', process.cwd());
        
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
            console.log(`🔍 检查路径: ${p}`);
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
        console.log('📖 读取Excel文件...');
        const workbook = XLSX.readFile(excelPath);
        const sheetNames = workbook.SheetNames;
        console.log(`📋 发现工作表: ${sheetNames.join(', ')}`);
        
        // 处理第一个工作表（教师课程数据）
        const firstSheet = workbook.Sheets[sheetNames[0]];
        const rawJsonData = XLSX.utils.sheet_to_json(firstSheet);
        console.log(`📊 Sheet1原始数据量: ${rawJsonData.length} 条`);
        
        if (rawJsonData.length === 0) {
            throw new Error('❌ Excel文件中没有数据');
        }
        
        // 🎯 关键步骤：处理并映射数据
        console.log('🔄 开始数据映射转换...');
        const processedCourseData = processCourseData(rawJsonData);
        
        // 处理学生数据
        const studentsData = extractStudentData(workbook);
        
        // 写入文件
        console.log('💾 写入JSON文件...');
        
        // 确保写入的是处理后的数据（英文字段名）
        fs.writeFileSync('data.json', JSON.stringify(processedCourseData, null, 2));
        console.log('✅ data.json 写入完成');
        
        fs.writeFileSync('students.json', JSON.stringify(studentsData, null, 2));
        console.log('✅ students.json 写入完成');
        
        // 验证写入的文件
        const writtenData = JSON.parse(fs.readFileSync('data.json', 'utf8'));
        console.log('🔍 验证写入的数据字段名:', Object.keys(writtenData[0] || {}));
        
        // 最终统计
        console.log('\n📈 === 处理完成统计 ===');
        console.log(`📚 教师课程记录: ${processedCourseData.length} 条`);
        console.log(`👥 学生数据: ${studentsData.length} 人`);
        console.log(`📝 学习记录总数: ${studentsData.reduce((sum, s) => sum + (s.studySessions?.length || 0), 0)} 条`);
        console.log('🎯 数据字段名已统一为英文格式');
        console.log('✅ 处理完成！');
        
    } catch (error) {
        console.error('💥 处理Excel时发生严重错误:', error);
        console.error('错误堆栈:', error.stack);
        process.exit(1);
    }
}

// 执行主函数
console.log('🎬 启动Excel处理程序...');
processExcel();
