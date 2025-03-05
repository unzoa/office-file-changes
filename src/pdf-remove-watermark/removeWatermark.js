const fs = require('fs');
const path = require('path');
const { PDFDocument, rgb } = require('pdf-lib');

async function removeWatermark(pdfPath) {
    // 读取PDF文件
    const data = await fs.promises.readFile(pdfPath);
    const pdfDoc = await PDFDocument.load(data);

    // 遍历所有页面
    const pages = pdfDoc.getPages();
    pages.forEach(page => {
        // 获取页面尺寸
        const { width, height } = page.getSize();

        // 在固定位置绘制白色矩形覆盖水印（根据实际位置调整坐标）
        page.drawRectangle({
            x: 0,  // 右侧位置
            y: height - 40,          // 底部位置
            width: width,
            height: 30,
            color: rgb(1, 1, 1), // 白色
            borderWidth: 0,
        });
    });

    // 保存修改后的PDF（添加压缩选项）
    const pdfBytes = await pdfDoc.save({
        useObjectStreams: true,  // 启用对象流
        compress: true           // 启用压缩
    });

    const outputPath = path.join('processed', path.basename(pdfPath));
    await fs.promises.writeFile(outputPath, pdfBytes);
}

// 处理files目录下所有PDF文件
async function processAllPDFs() {
    const dir = './files';
    const files = fs.readdirSync(dir);

    // 创建输出目录
    if (!fs.existsSync('./processed')) {
        fs.mkdirSync('./processed');
    }

    for (const file of files) {
        if (path.extname(file) === '.pdf') {
            await removeWatermark(path.join(dir, file));
            console.log(`已处理文件: ${file}`);
        }
    }
}

processAllPDFs().catch(console.error);
