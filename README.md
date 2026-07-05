# SkyGuard 低空智眼

SkyGuard 是面向重点区域低空运行监管的创业计划书与 WebDemo 项目，覆盖行业机会、痛点验证、产品服务、技术架构、数据体系、市场进入、商业模式、风险控制、财务预测、融资方案、附录材料和可运行前端演示。

## 在线演示

- GitHub Pages: https://2711944586.github.io/cxcy/
- 访问二维码: `05_deliverables/SkyGuard_GitHubPages_QR.png`
- 前端目录: `04_web_demo/`
- 部署方式: GitHub Actions 自动构建 Vite 静态站并发布到 GitHub Pages

## 项目结构

```text
.
├── .github/workflows/deploy-pages.yml   # GitHub Pages 静态部署流程
├── 01_references/                       # 课程要求、参考计划书、通用规范
├── 02_source_assets/                    # 原始数据、图片素材与整理清单
├── 03_plan_generation/                  # 商业计划书图表、表格与生成脚本
├── 04_web_demo/                         # React + TypeScript + Vite WebDemo
├── 05_deliverables/                     # Word、PDF、截图、表格等最终交付物
├── 99_archive/                          # 历史计划稿与旧版根目录资料
└── README.md
```

根目录命名统一采用 `NN_lower_snake_case`；中文标题保留在具体文档和素材文件名中。`02_source_assets/` 和 `01_references/` 用于追溯原始数据、图片素材、参考计划书和课程要求。静态站部署以 `04_web_demo/` 为入口，最终可交付版本以 `05_deliverables/` 中的 Word、PDF、截图和表格为准。

## 核心交付物

- `05_deliverables/SkyGuard_商业计划书.docx`: 最新完整版商业计划书
- `05_deliverables/SkyGuard_商业计划书.pdf`: 渲染版 PDF
- `05_deliverables/SkyGuard_商业计划书_封面图.png`: 封面视觉图
- `05_deliverables/SkyGuard_WebDemo_截图包/`: WebDemo 页面截图
- `05_deliverables/SkyGuard_财务测算表.xlsx`: 财务测算表
- `05_deliverables/SkyGuard_证据资料索引表.xlsx`: 来源与证据索引

## 本地运行 WebDemo

```bash
cd 04_web_demo
npm ci
npm run dev
```

本地构建 GitHub Pages 同款静态产物：

```bash
cd 04_web_demo
npm run build:pages
```

`build:pages` 会在 Vite 构建后把 `02_source_assets/_整理/数据`、`02_source_assets/_整理/图` 和 `01_references/` 复制到 `dist/source-data/`，线上数据中心页面会直接读取这些静态数据、图片和参考材料。

路由巡检：

```bash
cd 04_web_demo
npm run audit:routes
```

## 重新生成计划书交付物

```bash
python 03_plan_generation/scripts/build_deliverables.py
```

生成脚本会刷新 `05_deliverables/` 下的 Word、PDF、图表目录、截图索引、财务表和附录材料。重新生成前请关闭正在打开的 Word/WPS 文档窗口，避免文件被锁定。

## GitHub Pages 部署

推送到 `main` 分支后，`.github/workflows/deploy-pages.yml` 会自动执行：

1. 安装 `04_web_demo` 依赖
2. 使用 `/cxcy/` 作为 Vite 静态资源前缀构建
3. 复制全量原始数据、图片素材和参考材料到 `source-data/`
4. 复制 `index.html` 为 `404.html`，支持直接访问内部路由
5. 发布到 GitHub Pages

也可以在 GitHub 仓库的 Actions 页面手动运行 `Deploy SkyGuard WebDemo`。
