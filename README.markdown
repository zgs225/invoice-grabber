README
===

本项目是一个 python 脚本。用于从邮件服务器中获取发票，再通过百度的 OCR
服务识别发票信息后，根据发票信息重命名发票文件，并生成报销所需要的 csv 文件。

## 依赖

* python 3.x
* Firefox
* 百度智能云文字识别账号

## 使用方式

* 安装需要的依赖 `pip install -r requirements.txt`
* 复制 `config.example.json` 为 `config.json`, 并修改必要的参数
* `python main.py`

该脚本执行4个步骤的操作：

* 下载：从邮箱中下载发票到 `config.download_dir` 中
* OCR: 将 `config.download_dir` 中的发票通过百度 OCR 识别，并将结果写入 `config.output.recognized_json` 中
* 重命名：根据发票信息重命名成报销需要的名称到 `config.output.dir` 目录
* 生成报销记录：生成报销记录到 `config.output.result_csv` 中

在运行脚本时，可以通过添加参数跳过上述的任意阶段：

* `--no-email`: 跳过下载
* `--no-ocr`: 跳过 OCR
* `--no-rename`: 跳过重命名
* `--no-excel`: 跳过生成报销记录
