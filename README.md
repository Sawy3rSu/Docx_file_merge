### README.md

```markdown
# 合并多个 .docx 文件脚本

## 功能描述

本脚本用于合并同一目录下的所有 `.docx` 文件，并在每个文件内容前加上文件名。每个文件内容之间通过分页符进行区分。

## 环境准备

1. **Python环境**：
   - Python 3.x（推荐3.12或更高版本）

2. **安装依赖库**：
   - `python-docx`：用于读取和写入 `.docx` 文件。

   安装方法：
   ```bash
   pip install python-docx
   ```

## 使用方法

1. **编辑脚本**：
   - 修改脚本中的输入和输出目录路径。
   - 脚本默认读取当前目录下的 `.docx` 文件，并将合并后的文件保存在同一目录下。

2. **运行脚本**：
   - 将脚本保存为 `merge_docx.py`。
   - 在命令行或终端中，切换到脚本所在的目录。
   - 运行脚本：
     ```bash
     python merge_docx.py
     ```

## 脚本结构

### 文件结构

```plaintext
- merge_docx.py
- README.md
```

### 脚本代码

```python
from docx import Document
import os

def merge_docx_files(output_filename, directory):
    """
    合并同一目录下的所有.docx文件到一个文档中，并在每个文件内容前加上文件名。
    
    :param output_filename: 输出文件名（含路径）
    :param directory: 包含.docx文件的目录
    """
    # 创建一个新的Document对象作为输出文档
    merged_document = Document()

    # 遍历指定目录下的所有.docx文件
    for filename in os.listdir(directory):
        if filename.endswith('.docx'):
            file_path = os.path.join(directory, filename)
            print(f"正在处理文件: {file_path}")
            
            # 读取当前文件
            current_document = Document(file_path)
            
            # 在当前文件内容前加上文件名，并设置为标题格式
            heading_paragraph = merged_document.add_paragraph()
            heading_run = heading_paragraph.add_run(f'{filename}\n\n')
            heading_paragraph.style = 'Heading 1'  # 设置为标题样式
            
            # 将当前文件的内容逐段追加到输出文档
            for paragraph in current_document.paragraphs:
                new_paragraph = merged_document.add_paragraph(paragraph.text)
                # 复制段落格式
                new_paragraph.style = current_document.styles[paragraph.style.name]
                
                # 复制段落内的Run对象
                for run in paragraph.runs:
                    new_run = new_paragraph.add_run(run.text or '')
                    new_run.bold = run.bold
                    new_run.italic = run.italic
                    new_run.underline = run.underline
                    new_run.font.size = run.font.size
                    new_run.font.name = run.font.name
                    
            # 添加一个分页符以区分不同文件
            merged_document.add_page_break()

    # 保存合并后的文档
    output_path = os.path.join(directory, output_filename)
    print(f"尝试保存到: {output_path}")
    try:
        merged_document.save(output_path)
        print(f"合并完成，已保存为 {output_filename}")
    except Exception as e:
        print(f"保存文件时发生错误: {e}")

# 定义输入和输出目录及文件名
input_directory = r'C:\Users\Yi\Desktop\scraper'
output_filename = 'merged.docx'

# 调用函数合并文件
merge_docx_files(output_filename, input_directory)
```

## 注意事项

1. **路径格式**：
   - 在Windows系统中，路径中的反斜杠`\`需要转义为`\\`，或者使用原始字符串`r'C:\path\to\your\directory'`。

2. **权限**：
   - 确保你有足够的权限在指定目录下读取文件和写入文件。

3. **文件名冲突**：
   - 如果输出目录下已有同名文件，将会被覆盖。

## 示例

假设你的输入文件夹为 `C:\Users\Yi\Desktop\scraper`，并且你想将合并后的文件保存在同一目录下，文件名为 `merged.docx`。

1. **编辑脚本**：
   - 修改 `input_directory` 和 `output_filename` 变量。

2. **运行脚本**：
   - 运行脚本：
     ```bash
     python merge_docx.py
     ```

## 联系方式

如有任何问题或建议，请在项目仓库中提交Issue。

---
