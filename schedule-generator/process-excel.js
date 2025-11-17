const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// -------------------------------------------------------------
// 🧩 辅助函数：清理日期格式
// -------------------------------------------------------------
function cleanDate(dateStr) {
    if (!dateStr) return '';

    if (typeof dateStr === 'string' && dateStr.includes(' ')) {
        return dateStr.split(' ')[0];
    }

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

// -------------------------------------------------------------
// 🧩 课程数据处理函数（中文字段 → 英文字段）
// -------------------------------------------------------------
function processCourseData(rawData) {
    try {
        console.log('🔄 开始处理课程数据...');
        console.log('📊 原始数据总数:', rawData.length);
        console.log('📋 原始数据示例（前2行）:', JSON.stringify(rawData.slice(0, 2), null, 2));

        const processedData = rawData.map((row, index) => {
            const mappedRow = {
                id: index + 1,
                teacher: row['任课教师'] || row['教师'] || row['老师'] || '',
                topic: row['教学主题'] || row['主题'] || row['课程主题'] || '',
                session: row['课时'] || row['节次'] || '',
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

        // ✅ 修改点：不再提前丢掉空行，让我们确保所有数据都读进来
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

// -------------------------------------------------------------
// 🧩 学生数据提取函数
// -------------------------------------------------------------
function extractStudentData(workbook) {
    try {
        console.log('👥 开始提取学生数据...');
        const sheetNames = workbook.SheetNames;
        console.log('📄 所有工作表:', sheetNames);

        if (!sheetNames.includes('Sheet2') && !sheetNames.includes('sheet2')) {
            console.log('⚠️ 未找到Sheet2，跳过学生数据');
            return [];
        }

        const studentSheet = workbook.Sheets['Sheet2'] || workbook.Sheets['sheet2'];
        const studentRawData = XLSX.utils.sheet_to_json(studentSheet, {
            defval: '',
            blankrows: true,
            range: 0
        });
        console.log('🔍 Sheet2原始数据量:', studentRawData.length);

        if (studentRawData.length === 0) {
            console.log('📭 Sheet2为空');
            return [];
        }

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

// -------------------------------------------------------------
// 🧩 主处理逻辑
// -------------------------------------------------------------
function processExcel() {
    try {
        console.log('🚀 开始处理Excel文件...');
        console.log('📂 当前工作目录:', process.cwd());

        // 可能的路径
        const possiblePaths = [
            'schedule.xlsx',
            path.join(__dirname, 'schedule.xlsx'),
            '../schedule.xlsx',
            path.resolve(process.cwd(), 'schedule.xlsx')
        ];

        let excelPath = '';
        for (const p of possiblePaths) {
            console.log(`🔍 检查路径: ${p}`);
            if (fs.existsSync(p)) {
                excelPath = p;
                console.log(`✅ 找到Excel文件: ${excelPath}`);
                break;
            }
        }

        if (!excelPath) throw new Error('❌ 无法找到 schedule.xlsx 文件');

        // 读取Excel
        console.log('📖 读取Excel文件...');
        const workbook = XLSX.readFile(excelPath);
        const sheetNames = workbook.SheetNames;
        console.log(`📋 发现工作表: ${sheetNames.join(', ')}`);

        // ✅ 修改点1：强制匹配 sheet1 或 Sheet1
        const firstSheet =
            workbook.Sheets['sheet1'] ||
            workbook.Sheets['Sheet1'] ||
            workbook.Sheets[sheetNames[0]];

        // ✅ 修改点2：防止截断数据的参数配置
        const rawJsonData = XLSX.utils.sheet_to_json(firstSheet, {
            defval: '',
            blankrows: true,
            range: 0
        });

        console.log(`📊 Sheet1原始数据量: ${rawJsonData.length} 条`);
        console.log('📊 原始数据前5行预览:', JSON.stringify(rawJsonData.slice(0, 5), null, 2));

        if (rawJsonData.length === 0) {
            throw new Error('❌ Excel文件中没有数据');
        }

        // 处理课程数据
        const processedCourseData = processCourseData(rawJsonData);
        // 处理学生数据
        const studentsData = extractStudentData(workbook);

        // 写出文件
        console.log('💾 写入JSON文件...');
        fs.writeFileSync('data.json', JSON.stringify(processedCourseData, null, 2));
        fs.writeFileSync('students.json', JSON.stringify(studentsData, null, 2));
        console.log('✅ data.json & students.json 写入完成');

        // 验证写入
        const writtenData = JSON.parse(fs.readFileSync('data.json', 'utf8'));
        console.log('🔍 验证写入的数据字段名:', Object.keys(writtenData[0] || {}));

        // 统计信息
        console.log('\n📈 === 处理完成统计 ===');
        console.log(`📚 教师课程记录: ${processedCourseData.length} 条`);
        console.log(`👥 学生数据: ${studentsData.length} 人`);
        console.log(
            `📝 学习记录总数: ${studentsData.reduce(
                (sum, s) => sum + (s.studySessions?.length || 0),
                0
            )} 条`
        );
        console.log('🎯 数据字段名已统一为英文格式');
        console.log('✅ 处理完成！');
    } catch (error) {
        console.error('💥 处理Excel时发生严重错误:', error);
        console.error('错误堆栈:', error.stack);
        process.exit(1);
    }
}

// 🚀 启动程序
console.log('🎬 启动Excel处理程序...');
processExcel();
