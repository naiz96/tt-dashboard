# TT投放看板 - 云端部署

## 环境变量（必须在云平台设置）

| 变量名 | 说明 | 示例 |
|--------|------|------|
| `FEISHU_APP_ID` | 飞书应用 ID | `cli_xxxx` |
| `FEISHU_APP_SECRET` | 飞书应用密钥 | `xxxx` |
| `FEISHU_SPREADSHEET` | 飞书电子表格链接 | `https://xxx.feishu.cn/wiki/xxx?sheet=xxx` |

## 本地运行

```bash
python app.py
# 访问 http://localhost:8080
```

## 技术说明

- 纯 Python 标准库，无外部依赖
- 每次打开网页自动从飞书拉取最新数据（30秒缓存）
- dashboard.html 内嵌了离线数据作为兜底
- 支持 xlsx 文件上传（前端直接处理）
