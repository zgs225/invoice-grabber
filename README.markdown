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

## 配置

配置文件 `config.json` 中包含了以下信息：

`server`：IMAP4 邮件服务器地址；
`username`：登录邮箱的用户名；
`password`：登录邮箱的密码；
`proxy`：代理服务器地址,支持 socks4, socks5 和 http;
`download_dir`：电子发票保存的目录；
`since`：搜索电子发票的起始日期；
`util`：搜索电子发票的截止日期；
`baidu_ocr.app_id`：百度 OCR 应用 ID；
`baidu_ocr.api_key`：百度 OCR API Key；
`baidu_ocr.secret_key`：百度 OCR Secret Key；
`output.recognized_json`：保存 OCR 结果的文件，格式为 `json`;
`output.name`：你的名字，用于发票的重命名和报销记录的生成;
`output.dir`：保存 OCR 结果、重命名发票文件和报销记录的目录;
`output.result_csv`：保存生成的报销记录的文件，格式为 `csv`;
