# Excel Power Query 查询提取器

一个简单易用的工具，用于从Excel文件中提取Power Query查询，支持 .xlsx, .xlsm, .xlsb 格式。

## 功能特性

- ✅ 支持多种Excel格式：.xlsx, .xlsm, .xlsb
- ✅ 提取Excel文件中的所有Power Query查询
- ✅ 以美观的格式显示查询名称和M代码
- ✅ 支持下载为Markdown格式文件，便于后续编辑和分享
- ✅ 友好的用户界面，清晰的操作流程
- ✅ 详细的错误提示和进度反馈

## 技术栈

- **后端**：Python Flask
- **前端**：HTML, CSS, JavaScript
- **依赖库**：
  - Python: zipfile, base64, lxml, flask
  - JavaScript: JSZip

## 安装和运行

### 1. 克隆或下载项目

将项目文件下载到本地目录。

### 2. 安装依赖

打开命令行终端，进入项目目录，运行以下命令安装所需依赖：

```bash
pip install flask lxml
```

### 3. 启动Flask应用

在项目目录下运行以下命令：

```bash
python app.py
```

应用将在 http://127.0.0.1:5000 启动。

### 4. 访问应用

打开浏览器，访问 http://127.0.0.1:5000 即可使用。

## 使用方法

### 1. 选择Excel文件

点击"选择 Excel 文件"按钮，浏览并选择包含Power Query查询的Excel文件。

### 2. 提取查询

点击"提取 Power Query 查询"按钮，等待处理完成。

### 3. 查看结果

页面将显示提取到的所有Power Query查询，包括查询名称和完整的M代码。

### 4. 下载Markdown文件

如果成功提取到查询，页面会显示"下载为 Markdown 文件"按钮。点击该按钮，浏览器会自动下载生成的Markdown文件，文件名为 `原文件名_PowerQuery_Queries.md`。

## 项目结构

```
PowerQuery/
├── app.py                 # Flask后端应用
├── templates/
│   └── index.html         # 前端HTML页面
├── power_query_extractor.py  # 调试和测试脚本
├── README.md              # 项目说明文档
└── 工作簿1.xlsx           # 示例Excel文件（可选）
```

## 工作原理

1. **文件上传**：前端将Excel文件上传到Flask后端
2. **文件解析**：后端使用Python打开Excel文件，查找 `customXml` 目录下的XML文件
3. **提取DataMashup**：解析XML文件，提取 `DataMashup` 元素的Base64内容
4. **解码和提取**：解码Base64内容，提取嵌入式ZIP文件，从中读取 `Formulas/Section1.m` 文件
5. **解析M代码**：解析Section1.m文件中的Power Query M代码，提取所有查询
6. **返回结果**：将查询结果以JSON格式返回给前端
7. **显示和下载**：前端显示查询结果，并支持下载为Markdown文件

## Markdown文件格式

生成的Markdown文件格式示例：

```markdown
# Excel Power Query 查询

从文件: 工作簿1.xlsx 提取

共找到 2 个查询

## 查询 1: 查询1

```powerquery
let
    源 = #shared,
    转换为表 = Record.ToTable(源)
in
    转换为表
```

## 查询 2: 查询2

```powerquery
let
    源 = 3215,
    转换为表 = #table(1, {{源}})
in
    转换为表
```
```

## 注意事项

1. 仅支持包含Power Query查询的Excel文件
2. .xlsb格式文件的支持可能有限
3. 大文件可能需要较长时间处理
4. 请确保Excel文件没有密码保护

## 许可证

本项目采用 Apache 2.0 许可证。

## 贡献

欢迎提交Issue和Pull Request，帮助改进这个工具。

## 联系方式

如有问题或建议，请通过GitHub Issues反馈。
