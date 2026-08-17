# 南翔销售数据看板

自动从腾讯文档在线表格同步销售数据，生成可视化看板。

## 文件结构

- `sync.py` — 同步脚本，读取腾讯文档数据生成 data.js + orders_data.js
- `tdoc_reader.py` — 腾讯文档读取模块（支持本地 CLI 和 HTTP API 两种模式）
- `index.html` — 销售数据可视化看板
- `data.js` / `orders_data.js` — 自动生成的数据文件
- `.github/workflows/sync.yml` — GitHub Actions 定时同步

## GitHub Pages

看板部署在 GitHub Pages，访问地址：`https://<username>.github.io/<repo-name>/`

## 自动同步

GitHub Actions 每 2 小时自动运行 sync.py，读取腾讯文档最新数据并提交到仓库。

## 需要配置的 Secrets

在仓库 Settings → Secrets and variables → Actions 中添加：

- `TDOC_ACCESS_TOKEN` — 腾讯文档 OAuth access_token
- `TDOC_REFRESH_TOKEN` — 腾讯文档 OAuth refresh_token
- `TDOC_CLIENT_ID` — OAuth 客户端 ID
