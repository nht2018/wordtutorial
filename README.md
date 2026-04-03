
# 脚本使用说明

为了写大百科全书中的条目，我们准备了本目录中的脚本和规范文件用于将 Markdown 条目生成成 Word 文档。主要脚本和规范文件包括：

- `create_entry.py`
- `build_all_docx.sh`
- `markdown_rule.md`
- `prompt_tutorial.md` 


## 依赖

运行 `create_entry.py` 前，需要安装：

```bash
pip install python-docx math2docx
```

## 文件说明

### `create_entry.py`

作用：

- 读取一个 Markdown 条目文件
- 按固定模板生成 `.docx`
- 支持标题、摘要、目录、正文、参考文献
- 支持行内公式和行间公式
- 公式优先写成 Word 原生数学公式

示例：

```bash
python3 create_entry.py -f contents/随机梯度下降.md -o 随机梯度下降.docx
```

效果：

- 读取 `contents/随机梯度下降.md`
- 在当前工作目录生成 `随机梯度下降.docx`

参数说明：

- `-f`：必填，指定输入 Markdown 文件
- `-o`：可选，指定输出 Word 文件路径；如果不传，默认输出到当前目录，文件名与 Markdown 同名

### `build_all_docx.sh`

作用：

- 批量遍历 `contents/` 下所有 `.md`
- 逐个调用 `create_entry.py`
- 将生成的 `.docx` 输出到项目根目录

示例：

```bash
bash build_all_docx.sh
```

效果：

- 扫描项目根目录下的 `contents/*.md`
- 逐个生成对应的 `.docx`
- 输出到项目根目录

### `markdown_rule.md`

作用：

- 规定 Markdown 输入格式
- 说明 front matter、章节标题、公式、参考文献的写法
- 这个文件可以喂给 AI，让它根据你的要求生成对应的 Markdown 文件。

## 推荐使用方式

1. 先按 `markdown_rule.md` 编写或修改 `contents/*.md`
2. 先用 `create_entry.py` 单独生成一个条目，检查格式是否正确
3. 如果要批量生成项目根目录 `contents/` 下的全部条目，再运行 `build_all_docx.sh`


### `prompt_tutorial.md`

以后还有类似的写word的任务怎么办？可以参考 `prompt_tutorial.md` 中的教程生成脚本，将排版和内容分离，形成可复用的工作流。

