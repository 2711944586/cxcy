# SkyGuard Web Demo 部署说明

## 本地运行
```powershell
cd 04_web_demo
npm install
npm run dev
```

## 打包
```powershell
npm run build
```

## 路由巡检
```powershell
npm run audit:routes
```

巡检报告输出：

```text
05_deliverables/SkyGuard_WebDemo_路由巡检报告.json
```

## 生成截图
```powershell
npm run dev
node scripts/screenshot.mjs
```

## 说明
Demo使用React + TypeScript + Vite，本地JSON作为演示数据。运行数据为演示样本，不代表真实飞行记录或真实客户部署。

## 当前验收

- `npm run build` 通过。
- `npm run audit:routes` 通过，22 个路由均无控制台错误、页面错误或资源加载失败。
- 截图包覆盖 21 张页面截图，包括首页、综合态势、指挥屏和移动端。
- Vite 已通过 manualChunks 拆分图表与图标依赖，当前构建无单 chunk 大于 500 kB 警告。
