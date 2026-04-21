# paper-format
格式化毕业论文

本项目用于第一步或第二步格式化毕业论文文档，应该是使用python。

1.人为设置好一级标题，二级标题，正文
2.选择一级标题字体字号，选择二级标题的字体字号，选择正文的字体字号
3.运行程序修改文件字体字号

基本类似于这种较为固定和简单的格式处理。


字号设置（通用标准）：正文小四字号；封面题目二号黑体；一级标题小三号黑体；二级标题四号黑体；三级标题小四宋体加粗；摘要、关键词小四宋体；参考文献小四宋体。页脚与页码
页脚与页码
忽略封面
图表编号
参考文件
致谢格式

## MVP 使用方式（Python 3.13 / 3.14）

1. 安装依赖：

```bash
python -m pip install -r requirements.txt
```

2. 配置 `config.json`（默认文档目录就是当前项目目录）：

```json
{
  "directory": ".",
  "styles": {
    "body": { "east_asia_font": "宋体", "size_pt": 12, "bold": false, "first_line_indent_chars": 2 },
    "h1": { "east_asia_font": "黑体", "size_pt": 15, "bold": false },
    "h2": { "east_asia_font": "黑体", "size_pt": 14, "bold": false },
    "h3": { "east_asia_font": "宋体", "size_pt": 12, "bold": true }
  },
  "page": {
    "top_margin_pt": 72,
    "bottom_margin_pt": 72,
    "left_margin_pt": 90,
    "right_margin_pt": 90
  }
}
```

3. 运行脚本（目录参数可省略，默认读取 `config.json` 中的 `directory`）：

```bash
python format_docs.py "D:\your\docx\directory"
# 或
python format_docs.py
```

也可以指定配置文件：

```bash
python format_docs.py --config "D:\path\to\config.json"
```

4. 输出结果：

- 脚本会递归处理该目录下所有 `.docx` 文件（跳过临时文件 `~$` 开头）。
- 每个原文件会在同目录生成一个副本，文件名后缀为 `_formatted`。

当前 MVP 规则：

- 正文（`Normal`）：小四（12pt），宋体，首行缩进 2 字符（按字号换算）
- 正文行间距：1.5 倍
- 一级标题（`Heading 1`）：小三（15pt），黑体
- 二级标题（`Heading 2`）：四号（14pt），黑体
- 三级标题（`Heading 3`）：小四（12pt），宋体加粗
- 英文和数字字体：`Times New Roman`（中文保持各样式东亚字体）
- 页边距（cm）：上 2.54，下 2.54，左 3.17，右 3.17
- 页眉：`某某大学学士论文`，距顶端 1.5cm，宋体五号，居中
- 页脚：距底端 1.75cm，页码五号（`Times New Roman`），居中
- 首页页眉页脚留空；页码显示从第二页开始（起始值按配置可调）
- 若文档中存在标题为“目录”的段落，会自动插入 TOC 域，并在 Word 打开时触发域更新