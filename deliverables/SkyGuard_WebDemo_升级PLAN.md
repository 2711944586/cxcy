# SkyGuard WebDemo 升级 PLAN

本文件与 `skyguard-demo/WebDemo_升级PLAN.md` 保持同一口径，用于交付目录审阅。

## 定位

WebDemo 是低空运行监管工作台，用于证明 SkyGuard 能完成“态势、计划、围栏、工单、移动处置、报表”的连续业务闭环。

## 当前实现

| 模块 | 状态 |
|---|---|
| 首页 | 值班域、规则包、KPI、态势地图、事件详情、值班节奏、待处理队列 |
| 地图 | HUD、图层状态、风险图例、实时事件队列 |
| 演示中心 | 五类脚本和连续跳转路径 |
| 围栏规则 | 可触发模拟告警 |
| 事件工单 | 表格列表和事件详情 |
| 运行报表 | 多图表与导出/打印反馈 |
| 移动端 | 390px 截图可访问导航和事件列表 |
| 指挥屏 | 深色大屏、指标、地图、事件侧栏、状态条 |

## 验收命令

```powershell
cd skyguard-demo
npm run build
npm run dev -- --host 127.0.0.1
node scripts/screenshot.mjs
```

## 当前验收结果

- `npm run build` 通过。
- `npm run audit:routes` 通过，22 个路由均无控制台错误、页面错误和资源加载失败。
- 截图包包含 21 张 PNG。
- Vite 已通过 `manualChunks` 拆分图表与图标依赖，当前构建无单 chunk 大于 500 kB 警告。
