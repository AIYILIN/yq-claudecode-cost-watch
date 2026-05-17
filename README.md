# API 花费统计仪表盘

DeepSeek API 用量和花费的可视化仪表盘，支持本地 CSV 数据管理和在线自动更新。

## 功能

- **数据可视化**：每日花费趋势、模型花费占比、用户花费排名、Token 用量分布
- **在线更新**：一键从 DeepSeek 平台自动拉取最新用量数据
- **本地编辑**：支持增删改数据行、用户别名管理、用户黑名单
- **主题切换**：深色/浅色模式，自动保存偏好

## 安装

```bash
# 1. 安装 Python 依赖
pip install flask playwright openpyxl

# 2. 安装 Playwright Chromium 浏览器 (可选，如果有 Chrome 则不需要)
playwright install chromium
```

## 启动

```bash
python app.py
```

浏览器打开 `http://localhost:5000`

## 在线更新

点击仪表盘右上角的 **"在线更新"** 按钮：

1. Chrome 浏览器会自动弹出，打开 [DeepSeek 用量页面](https://platform.deepseek.com/usage)
2. **首次使用**需要在浏览器中登录 DeepSeek 平台账号（后续自动记住登录状态）
3. 登录后脚本会自动点击"导出"按钮，下载数据压缩包
4. 数据会自动解压、转换、合并到本地 CSV 文件中
5. 仪表盘自动刷新显示最新数据

> 如果提示"未找到导出按钮"，请查看 `.fetch-diag.txt` 文件中的诊断信息，
> 将其中的按钮文本反馈给开发者以更新匹配规则。

## 文件说明

| 文件 | 说明 |
|------|------|
| `app.py` | Flask 后端 |
| `templates/index.html` | 前端仪表盘 |
| `fetch_online.py` | 在线数据爬取脚本 |
| `amount-YYYY-M.csv` | Token 用量明细数据 |
| `cost-YYYY-M.csv` | 每日花费汇总数据 |
| `user-aliases.json` | 用户名 → 备注名映射 |
| `deleted-users.json` | 已删除（隐藏）用户列表 |

## 数据格式

**用量文件** (`amount-YYYY-M.csv`)：

| 列名 | 说明 |
|------|------|
| user_id | 用户 ID |
| utc_date | UTC 日期 |
| model | 模型名称 |
| api_key_name | API Key 名称（用户） |
| api_key | API Key |
| type | Token 类型（output_tokens / input_cache_hit_tokens / input_cache_miss_tokens / request_count） |
| price | 单价（CNY） |
| amount | 数量 |

**花费文件** (`cost-YYYY-M.csv`)：

| 列名 | 说明 |
|------|------|
| user_id | 用户 ID |
| utc_date | UTC 日期 |
| model | 模型名称 |
| wallet_type | 钱包类型 |
| cost | 花费（CNY） |
| currency | 货币单位 |

## 技术栈

- **后端**：Python 3 + Flask
- **前端**：原生 HTML/CSS/JS + Chart.js
- **自动化**：Playwright（浏览器控制）
