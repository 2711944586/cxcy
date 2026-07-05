# SkyGuard 低空智眼

SkyGuard 是面向重点区域低空运行监管的创业计划书与 WebDemo 项目，覆盖行业机会、痛点验证、产品服务、技术架构、数据体系、市场进入、商业模式、风险控制、财务预测、融资方案、附录材料和可运行前端演示。

## 在线演示

- GitHub Pages: https://2711944586.github.io/cxcy/
- 访问二维码: `deliverables/SkyGuard_GitHubPages_QR.png`
- 前端目录: `skyguard-demo/`
- 部署方式: GitHub Actions 自动构建 Vite 静态站并发布到 GitHub Pages

## 项目结构

```text
.
├── .github/workflows/deploy-pages.yml   # GitHub Pages 静态部署流程
├── deliverables/                        # Word、PDF、图表包、截图包等最终交付物
├── skyguard-demo/                       # React + TypeScript + Vite WebDemo
├── skyguard-plan/                       # 商业计划书、图表与交付物生成脚本
├── 数据以及图/                           # 原始数据、图片素材与整理清单
├── 参考/                                # 参考计划书与通用标准规范
└── README.md
```

`数据以及图/` 和 `参考/` 已随仓库保留，用于完整追溯原始数据、图片素材、参考计划书和课程要求。静态站部署仍以 `skyguard-demo/` 为入口，最终可交付版本以 `deliverables/` 中的 Word、PDF、截图包和图表包为准。

## 核心交付物

- `deliverables/SkyGuard_商业计划书.docx`: 最新完整版商业计划书
- `deliverables/SkyGuard_商业计划书.pdf`: 渲染版 PDF
- `deliverables/SkyGuard_WebDemo_截图包/`: WebDemo 页面截图
- `deliverables/SkyGuard_商业计划书_图表包.zip`: 图表素材包
- `deliverables/SkyGuard_财务测算表.xlsx`: 财务测算表
- `deliverables/SkyGuard_证据资料索引表.xlsx`: 来源与证据索引

## 本地运行 WebDemo

```bash
cd skyguard-demo
npm ci
npm run dev
```

本地构建：

```bash
cd skyguard-demo
npm run build
```

路由巡检：

```bash
cd skyguard-demo
npm run audit:routes
```

## 重新生成计划书交付物

```bash
python skyguard-plan/scripts/build_deliverables.py
```

生成脚本会刷新 `deliverables/` 下的 Word、PDF、图表、截图包索引、财务表和附录材料。重新生成前请关闭正在打开的 Word/WPS 文档窗口，避免文件被锁定。

## GitHub Pages 部署

推送到 `main` 分支后，`.github/workflows/deploy-pages.yml` 会自动执行：

1. 安装 `skyguard-demo` 依赖
2. 使用 `/cxcy/` 作为 Vite 静态资源前缀构建
3. 复制 `index.html` 为 `404.html`，支持直接访问内部路由
4. 发布到 GitHub Pages

也可以在 GitHub 仓库的 Actions 页面手动运行 `Deploy SkyGuard WebDemo`。
