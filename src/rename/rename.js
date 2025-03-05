const fs = require('fs');
const path = require('path');

// 源文件夹和目标文件夹
const sourceDir = './A';
const targetDir = './B';

// 确保目标文件夹存在
if (!fs.existsSync(targetDir)) {
    fs.mkdirSync(targetDir);
}

// 读取源文件夹中的所有文件
fs.readdir(sourceDir, (err, files) => {
    if (err) {
        console.error('读取文件夹出错:', err);
        return;
    }

    files.forEach(file => {
        // 获取文件的完整路径
        const sourcePath = path.join(sourceDir, file);

        // 分离文件名和扩展名
        const ext = path.extname(file);
        const nameWithoutExt = path.basename(file, ext);

        // 移除从"-备战"开始到末尾的所有文字
        const newNameWithoutExt = nameWithoutExt.split('-备战')[0];

        // 组合新文件名（加回扩展名）
        const newFileName = newNameWithoutExt + ext;

        // 设置目标路径
        const targetPath = path.join(targetDir, newFileName);

        // 复制文件到新位置
        fs.copyFile(sourcePath, targetPath, (err) => {
            if (err) {
                console.error(`复制文件 ${file} 时出错:`, err);
                return;
            }
            console.log(`成功处理文件: ${file} -> ${newFileName}`);
        });
    });
});