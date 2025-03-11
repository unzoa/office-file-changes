# 下载文件处理

## PDF遮盖否一部分

  > 源文件夹：files，目标文件夹：processed

  ```bash
  node removeWatermark.js
  ```

## 去掉word最后一页

  > 源文件夹：docs，目标文件夹：output

  ```bash
  node process_word.js
  ```

## 去掉页眉

  > 源文件夹：docs，目标文件夹：docs

  ```bash
  python src/remove_header.py
  ```

## Word转PDF

  > 源文件夹：docs，目标文件夹：pdfs

  ### macos

  > 使用了命令行版本 libreoffice，**很慢**

  ```bash
  python src/word-2-pdf/word2pdf.py
  ```

## 合并金数据excel

  > 源文件夹：金数据-提取和合并，目标文件夹：/
  ```bash
  node src/merge_excel.py
  ```