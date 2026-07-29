# 更新记录

## v2.3.1 - 2026-07-29

- 增加独立 Windows 安装程序。
- 安装程序自动扫描注册表和常见安装目录中的 CorelDRAW。
- 支持勾选多个 CDR 版本后批量安装。
- 支持勾选多个 CDR 版本后一键卸载。
- 安装时只部署插件运行文件，不覆盖已有 `config.json`、`svg-history.tsv` 和 `output`。
- 发布包不包含开发机上的 API Key、模型档案、生成历史和生成结果。
- 移除 SuperSVG 方案及其模型、虚拟环境和依赖。
- 保留 CorelDRAW PowerTRACE 与本地 VTracer 两种描摹方案。
- VTracer 保留三个模式：默认彩色 Logo / 扁平插画、黑白 / 线稿、照片 / 多色图 / 油画素材。
