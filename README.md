# Excel 中英文翻译工具

一个简单的 Excel 文件中英文翻译工具，支持批量翻译 Excel 文件中的中文内容为英文。

## 功能特点

- 支持多个工作表的批量翻译
- 保留原始列，在旁边新增对应的英文列
- 自动保存为新文件，不会修改原文件
- 现代化的用户界面
- 实时显示翻译进度
- 支持取消操作

## 使用方法

1. 下载并运行 `Excel翻译工具.exe`
2. 点击"选择文件"按钮，选择要翻译的 Excel 文件
3. 点击"开始翻译"按钮（或按 Ctrl+Enter）开始翻译
4. 等待翻译完成，翻译后的文件会自动保存在原文件所在目录
   - 新文件名为：原文件名_translated.xlsx

## 开发环境

- Python 3.12
- 依赖库：
  - deep_translator
  - pandas
  - openpyxl
  - tkinter

## 安装依赖
```bash
pip install -r requirements.txt
```

## 从源码运行
```bash
python excel_translator.py
```

## 注意事项

- 使用前请确保电脑已连接网络
- 翻译过程中请勿关闭程序
- 建议先用小文件测试效果

## 许可证

MIT License
