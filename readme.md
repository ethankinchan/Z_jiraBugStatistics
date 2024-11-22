# JIRA Bug Statistics Generator

## 项目简介
这是一个用于生成 JIRA bug 统计报告的自动化工具。它可以从 JIRA 中获取 bug 数据，并生成统计表格和可视化图表，帮助团队更好地了解 bug 的分布情况。

## 功能特点
- 自动连接 JIRA 并获取 bug 数据
- 按照优先级（Urgency）对 bug 进行分类统计
- 生成 bug 状态分布的统计表格（Excel格式）
- 生成 bug 状态分布的饼图可视化
- 支持数据缓存，提高查询效率
- 自动创建报告目录，使用时间戳避免文件覆盖

## 安装要求
- Python 3.7+
- 依赖包：
  ```bash
  pip install jira pandas matplotlib seaborn openpyxl
  ```

## 配置说明
1. 在项目根目录创建 `config.json` 文件： 
```
{
    "jira_server": "https://jira.zebra.com",
    "username": "JC5665",
    "password": "your_password_here"
}