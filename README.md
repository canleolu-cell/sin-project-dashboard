# CNGR 新加坡转口贸易监控看板

这个项目用 Excel 作为数据源，自动生成 `data/lots.json`，再由 `index.html` 读取 JSON 展示 GitHub Pages 看板。

## 日常更新

1. 更新 Excel 源文件。
2. 用最新文件覆盖 `data/lot-details.xlsx`。
3. 提交并推送到 GitHub。
4. GitHub Actions 会自动生成 `data/lots.json` 并发布 GitHub Pages。

## 本地生成数据

```powershell
python scripts/build-data.py --input data/lot-details.xlsx --output data/lots.json
```

如果要在本机预览页面，请在项目目录启动一个静态服务器，然后打开浏览器访问本地地址。
