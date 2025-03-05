const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

function processAllWordFiles() {
    const inputRoot = './docs';
    const outputRoot = './output';

    // 递归遍历目录
    function walk(dir) {
        const files = fs.readdirSync(dir);
        files.forEach(file => {
            const filePath = path.join(dir, file);
            if (fs.statSync(filePath).isDirectory()) {
                walk(filePath);
            } else if (path.extname(file) === '.docx') {
                // 新增临时文件过滤逻辑
                if (file.startsWith('.~') || file.startsWith('~$')) { // 跳过临时文件
                    console.log(`跳过临时文件: ${filePath}`);
                    return;
                }

                const relPath = path.relative(inputRoot, dir);
                const outDir = path.join(outputRoot, relPath);
                if (!fs.existsSync(outDir)) {
                    fs.mkdirSync(outDir, { recursive: true });
                }

                const outputPath = path.join(outDir, file);
                try {
                    // 增加文件存在性检查
                    if (!fs.existsSync(filePath)) {
                        console.error(`文件不存在: ${filePath}`);
                        return;
                    }
                    execSync(`python doc_handler.py "${filePath}" "${outputPath}"`);
                    console.log(`Processed: ${filePath}`);
                } catch (error) {
                    console.error(`处理失败: ${filePath}`, error.message);
                }
            }
        });
    }

    walk(inputRoot);
}

processAllWordFiles();
