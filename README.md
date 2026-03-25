# Research Paper Portal

个人使用的最小论文 DOCX 格式化与 GitHub 发布工具：FastAPI + Jinja2 + Bootstrap 5 + SQLite，数据与上传均限制在项目目录内。

## 功能

- Session 登录（默认用户见下方初始化脚本）
- 上传原始 DOCX → 按 `data/reference_samples/` 中的参考文档推导样式并格式化 → 下载
- 上传审阅后的最终 DOCX → 推送到 GitHub 仓库 `published/` 目录

## 目录结构

见项目内 `app/`、`data/`、`scripts/`；参考样式位于 `data/reference_samples/`。

## 部署（与其它站点隔离）

1. 将整个目录复制到独立路径，例如 `/opt/research_paper_portal/`（不要覆盖其它项目）。
2. 创建虚拟环境并安装依赖：

```bash
cd /opt/research_paper_portal
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

3. 复制环境变量：`cp .env.example .env`，编辑 `SECRET_KEY`、`PORT`、`GITHUB_TOKEN` 等。
4. 部署前检查端口是否空闲：`python scripts/check_port.py 8765`
5. 初始化数据库与用户：

```bash
export PYTHONPATH=$(pwd)
python scripts/seed_admin.py
```

6. 将参考样式 DOCX 放入 `data/reference_samples/`（文件名与 `.env` 中 `REFERENCE_STYLE_DOCX` 一致）。
7. 启动：

```bash
chmod +x run.sh
./run.sh
```

浏览器访问：`http://<服务器IP>:<PORT>`（默认端口见 `.env`，示例为 8765）。

**注意**：本仓库不自动修改 nginx/apache；若需域名与 HTTPS，请自行在反向代理中新增配置，勿影响现有站点。

## 默认账号

- 用户名：`admin`
- 密码：`admin123`

生产环境请尽快修改密码（需后续提供「改密」功能或手动更新 SQLite）。

## 参考样式更新方式

替换 `data/reference_samples/` 下对应 DOCX，或修改 `.env` 中 `REFERENCE_STYLE_DOCX` 指向新文件名，重启应用。详见该目录内 `README.txt`。

## 格式化说明与限制

- 从参考文档提取首段字体、字号、页边距等；无法完整克隆 Word 样式表时，使用学术默认（Times New Roman、层级标题、两端对齐等）。
- 按标题行启发式识别 Abstract、Introduction、References 等章节；表格通过 OOXML 深拷贝尽量保留。
- 图片、脚注、复杂域代码可能丢失或需人工复核。

## GitHub 发布

需 Personal Access Token（`repo` 权限）写入环境变量 `GITHUB_TOKEN`。服务端使用，勿写入前端或日志。

## systemd

参考 `research_paper_portal.service.example`，复制为独立 unit 文件后再 `enable`，勿覆盖系统已有服务。
