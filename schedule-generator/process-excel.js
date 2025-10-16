const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

// 添加主函数
function main() {
    console.log('开始处理Excel文件...');
    console.log('当前工作目录:', process.cwd());
    console.log('目录内容:', fs.readdirSync(process.cwd()).join(', '));
    
    try {
        // 读取Excel文件
        let workbook;
        let excelPath = '';
        
        // 按照优先级尝试不同路径
        const possiblePaths = [
            'schedule.xlsx',                           // 当前目录
            path.join(__dirname, 'schedule.xlsx'),     // 使用__dirname
            '../schedule.xlsx',                         // 上级目录
            path.resolve(process.cwd(), 'schedule.xlsx') // 绝对路径
        ];
        
        // 尝试找到文件
        let fileFound = false;
        for (const p of possiblePaths) {
            console.log(`尝试路径: ${p}`);
            if (fs.existsSync(p)) {
                excelPath = p;
                console.log(`找到Excel文件: ${excelPath}`);
                fileFound = true;
                break;
            }
        }
        
        if (!fileFound) {
            throw new Error('无法找到schedule.xlsx文件，所有尝试路径均失败');
        }
        
        // 读取Excel文件
        workbook = xlsx.readFile(excelPath);
        
        // 获取所有工作表
        const sheetNames = workbook.SheetNames;
        console.log(`工作表列表: ${sheetNames.join(', ')}`);
        
        // 获取第一个工作表数据
        const firstSheet = workbook.Sheets[sheetNames[0]];
        const jsonData = xlsx.utils.sheet_to_json(firstSheet);
        console.log(`成功读取第一个工作表数据，共 ${jsonData.length} 条记录`);
        
        // 处理数据：整理成所需的格式
        const processedData = processCourseData(jsonData);
        
        // 提取学生数据
        const students = extractStudentData(jsonData);
        
        // 写入data.json和students.json
        fs.writeFileSync('data.json', JSON.stringify(processedData, null, 2));
        console.log('成功生成 data.json');
        
        fs.writeFileSync('students.json', JSON.stringify(students, null, 2));
        console.log('成功生成 students.json');
        
        console.log('所有处理完成');
        
    } catch (error) {
        console.error('处理Excel时发生错误:', error);
        process.exit(1); // 非零退出码表示错误
    }
}

// 处理课程数据
function processCourseData(rawData) {
    // 将原始数据转换为所需格式
    try {
        console.log('开始处理课程数据...');
        
        // 根据您的具体需求处理数据
        // 这里假设您需要保留原始结构但可能需要添加一些计算字段
        const processedData = rawData.map((row, index) => {
            return {
                id: index + 1,  // 添加ID
                ...row,         // 保留所有原始字段
                processed: true // 标记为已处理
            };
        });
        
        console.log(`成功处理 ${processedData.length} 条课程数据`);
        return processedData;
    } catch (error) {
        console.error('处理课程数据时出错:', error);
        return rawData; // 出错时返回原始数据
    }
}

// 提取学生数据
function extractStudentData(rawData) {
    try {
        console.log('开始提取学生数据...');
        
        // 假设Excel中有一个"学生"列
        // 提取所有学生，并去重
        const allStudents = new Set();
        
        rawData.forEach(row => {
            // 根据您的数据结构，找到表示学生的字段
            // 例如，如果有"学生"或"姓名"列：
            if (row.Student) allStudents.add(row.Student);
            if (row.学生) allStudents.add(row.学生);
            if (row.Name) allStudents.add(row.Name);
            if (row.姓名) allStudents.add(row.姓名);
        });
        
        // 转换为数组并添加ID
        const students = Array.from(allStudents)
            .filter(student => student) // 过滤空值
            .map((student, index) => {
                return {
                    id: index + 1,
                    name: student
                };
            });
        
        console.log(`成功提取 ${students.length} 名学生数据`);
        return students;
    } catch (error) {
        console.error('提取学生数据时出错:', error);
        return []; // 出错时返回空数组
    }
}

// 执行主函数
main();
