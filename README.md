# Excel 数据 → Word 模板填充工具

将 Excel 表格数据批量填充到 Word 模板中，生成个性化文档。
效果展示
<img width="1005" height="995" alt="image" src="https://github.com/user-attachments/assets/01884738-77bc-4b8f-a563-517f73528a21" />
参考数据展示
<img width="1016" height="201" alt="image" src="https://github.com/user-attachments/assets/20dfb2b9-12d0-467c-a8f6-b346bd1bdc47" />
<img width="990" height="960" alt="image" src="https://github.com/user-attachments/assets/6761e5d3-ee28-4284-b61a-5f6d508b5de9" />

## 功能特性

- 📊 **批量生成** - 读取 Excel 一行数据，生成一个 Word 文档
- 🏷️ **占位符替换** - 自动替换模板中的 `{{字段名}}` 占位符
- 📁 **智能命名** - 自动使用姓名/标题等字段生成文件名
- 🖱️ **交互模式** - 无需记命令，图形化选择文件
- 📋 **字段匹配** - 智能对比 Excel 表头与模板占位符

## 使用方法

### 方式一：命令行参数

```bash
python excel-to-word-template.py <excel文件> <word模板> <输出目录>
```

**示例：**
```bash
python excel-to-word-template.py data.xlsx template.docx output/
```

### 方式二：交互模式

直接运行，程序会引导你选择文件：

```bash
python excel-to-word-template.py
```

## 模板写法

在 Word 文档中使用 `{{字段名}}` 作为占位符：

```
您好，{{姓名}}！

您的订单 {{订单号}} 已确认，
将于 {{日期}} 送达 {{地址}}。
```

Excel 表头对应字段名：

| 姓名 | 订单号 | 日期 | 地址 |
|------|--------|------|------|
| 张三 | A001 | 2026-04-01 | 北京市朝阳区 |

## 依赖安装

```bash
pip install openpyxl python-docx
```

或直接安装所有依赖：

```bash
pip install -r requirements.txt
```

## 项目结构

```
excel-to-word-generator/
├── excel-to-word-template.py   # 主程序
├── requirements.txt            # Python 依赖
├── samples/                    # 示例文件
│   ├── sample_data.xlsx       # 示例 Excel 数据
│   └── sample_template.docx    # 示例 Word 模板
└── output/                    # 生成文件输出目录
```

## 运行示例

```
==================================================
📄 Excel 数据 → Word 模板填充工具
==================================================

📂 选择 Excel 数据文件
   (直接按回车使用文件对话框，或手动输入路径)
   输入路径: data.xlsx
   ✅ 已选择: data.xlsx

📂 选择 Word 模板文件
   输入路径: template.docx
   ✅ 已选择: template.docx

📁 选择输出文件夹
   输入路径: output
   ✅ 已选择: output

🔍 正在分析 Word 模板...
   ✅ 找到 4 个占位符: ['姓名', '订单号', '日期', '地址']

确认继续？(y/n): y

🚀 开始生成 Word 文件...
   ✅ [1/100] 张三.docx
   ✅ [2/100] 李四.docx
   ...

🎉 完成！成功生成 100 个 Word 文件
📁 输出目录: output
==================================================
```

## 注意事项

1. **字段匹配** - Excel 表头需与模板 `{{字段名}}` 完全一致（区分大小写）
2. **日期处理** - Excel 日期会自动转换为 `YYYY-MM-DD` 格式
3. **文件名冲突** - 重复文件名会自动添加序号（如 `张三_1.docx`）
4. **批量确认** - 处理超过 100 行时会提示确认

## 技术栈

- Python 3.7+
- [openpyxl](https://openpyxl.readthedocs.io/) - Excel 读写
- [python-docx](https://python-docx.readthedocs.io/) - Word 操作
- tkinter（可选）- 图形化文件选择
