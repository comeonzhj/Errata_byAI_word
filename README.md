# AI勘误器 Word版 (AI Errata Tool for Word)

[English](README_EN.md) | 简体中文

一个基于AI的Word文档勘误工具，可以自动检查和标注文档中的错别字，并提供修改建议。

## 功能特性

- 🔍 自动检查Word文档中的错别字
- 📝 智能提供修改建议
- 🎯 高亮显示需要修改的地方
- 📊 实时显示处理进度
- 📋 支持导出勘误日志
- ⚙️ 灵活的API配置

## 环境要求

- Python 3.6+
- macOS系统（目前仅支持macOS）
- 以下Python包：
  - PyQt6 >= 6.4.0
  - python-docx >= 0.8.11
  - requests >= 2.28.0
  - json5 >= 0.9.5
  - pyinstaller >= 5.13.0（仅打包需要）

## 安装步骤

1. 克隆仓库：
```bash
git clone https://github.com/yourusername/Errata_byAI_word.git
cd Errata_byAI_word
```

2. 安装依赖：
```bash
pip3 install -r requirements.txt
```

## 使用方法

### 运行源码

```bash
python3 main.py
```

### 打包应用

```bash
pyinstaller main.spec
```
打包后的应用将在 `dist` 目录中生成。

### 配置说明

1. 点击「模型配置」按钮
2. 填写以下信息：
   - API URL：您的API接口地址
   - API Key：您的API密钥
   - 模型名称：使用的模型名称
   - System Prompt：系统提示词（已提供默认配置）

### 使用流程

1. 启动应用
2. 完成API配置
3. 点击「打开文档」选择Word文件
4. 点击「开始勘误」
5. 等待处理完成
6. 查看生成的带有修改建议的新文档
7. 可选：导出勘误日志

## 注意事项

- API密钥等敏感信息存储在用户目录下的 `.errata_word/config.json` 中
- 支持的文档格式为 .docx
- 建议在处理大文档时确保网络稳定

## 开源协议

本项目采用 MIT 协议开源，详见 [LICENSE](LICENSE) 文件。

## 贡献指南

欢迎提交Issue和Pull Request！

1. Fork 本仓库
2. 创建您的特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交您的改动 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 打开一个 Pull Request

## 联系方式

如有问题或建议，欢迎提交 Issue。