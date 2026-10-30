# 月度结算数据汇总

[在线使用月度结算助手](https://lyming77-monthly-electricity-settlement-streamlit-app-nurdrc.streamlit.app/)

上传现有的年度分析表和当月结算单，核对变更后生成更新后的 Excel。下一次使用时，上传上一次下载保存的分析表。

本目录可作为独立 GitHub 仓库部署到 Streamlit Community Cloud。当前线上运行环境使用 Python 3.14，入口为 `streamlit_app.py`。依赖固定为 Streamlit 1.66.1；这是截至 2026-10-29 核验的[官方最新稳定发行版](https://github.com/streamlit/streamlit/releases/tag/1.66.1)。

## 每月使用

1. 上传当前正在使用的年度分析表 `.xlsx`。
2. 上传新月份的结算单；可多选。按页面提示补充结算单不包含的参考价等数值。
3. 点击预览，核对识别的月份、场站、缺失文件和变更明细。
4. 生成并下载更新后的分析表和导入报告，保存在自己的电脑上。下载完整 ZIP 时同时包含更新前原表备份。

应用生成新文件，不会直接改写电脑上的原文件。关闭或刷新会话前下载结果；下次上传最新结果继续更新。单个上传文件上限 32 MB。网页部署后，上传文件会传输到该部署服务器处理。

## 部署到 Streamlit Community Cloud

1. 在自己的 GitHub 账号下新建一个独立仓库，例如 `monthly-settlement`。只上传**本目录内的程序文件**，包括隐藏目录 `.streamlit`；不要上传上级目录，也不要上传任何结算单、目标分析表或个人配置文件。
2. 打开 [Streamlit Community Cloud](https://share.streamlit.io/)，登录并关联 GitHub。部署私有仓库时，按页面要求授权访问该仓库。
3. 在对应 GitHub 用户的工作区点击 **Create app**，选择已有代码，填写仓库、分支以及入口 `streamlit_app.py`。
4. 在 **Advanced settings** 中选择 Python **3.14**，保存后部署。本程序不需要配置 API 密钥或 Secrets。部署成功后平台会提供一个 `streamlit.app` 地址。[官方部署步骤](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy)
5. 在应用设置的 **Sharing** 中选择 **Only specific people can view this app**，再添加需要访问的邮箱。私有仓库部署的应用默认私有，公共仓库部署的应用默认公开；应用是否公开可在此单独调整。Community Cloud 当前每次允许一个私有应用。[官方访问权限说明](https://docs.streamlit.io/deploy/streamlit-community-cloud/share-your-app)

仓库中的 `.gitignore` 已排除 Excel、导入结果和密钥文件。网页上传 GitHub 文件时仍需自行确认选择的是代码文件。本程序无需把业务数据写入仓库。

## 在本地运行

在 Python 3.12 环境中进入本目录。macOS/Linux 可执行：

```bash
python3.12 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m streamlit run streamlit_app.py
```

Windows PowerShell 可执行：

```powershell
py -3.12 -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe -m streamlit run streamlit_app.py
```

打开终端显示的本地地址，通常为 `http://localhost:8501`。项目配置位于 `.streamlit/config.toml`，因此应从本目录启动。

## Docker 自托管

已提供 Python 3.12 镜像配置，适合部署在自己的服务器或公司内网。进入本目录执行：

```bash
docker build -t monthly-settlement .
docker run --rm -p 127.0.0.1:8501:8501 monthly-settlement
```

在服务器本机访问 `http://localhost:8501`。对外提供服务时，由服务器的 HTTPS 反向代理转发至该端口，并按需要配置访问认证。镜像以普通用户运行，健康检查使用 Streamlit 的健康端点。无需额外安装系统级 Excel 软件。[Streamlit 官方 Docker 部署说明](https://docs.streamlit.io/deploy/tutorials/docker)

云端服务不会替代电脑上的归档目录。请保留原始结算单和每月下载的分析结果。

## 验证

安装依赖后运行：

```bash
python -m unittest discover -s tests -v
```

测试包含文件校验、输入变化后旧预览失效、会话隔离和网页流程。真实工作簿回归测试仅在本机从 `SETTLEMENT_FIXTURES` 指定目录读取；没有这些私人文件时会明确跳过，业务文件不需要上传到 GitHub。
