# GEC Market Analysis

本项目用于通过 Python + Playwright 获取广州电力交易中心绿电绿证交易平台中的绿证市场数据，并整理为 Excel 输出。

## 当前进展

- 已完成首页公开数据获取流程。
- 已支持获取指定年份全国绿证流出量前 10 省份数据。
- 已支持获取指定年份全国绿证流入量前 10 省份数据。
- 已将流出、流入结果分别写入 Excel 的不同 sheet：
  - `流出前10`
  - `流入前10`
- Excel 输出字段已整理为：
  - `排名`
  - `省份`
  - `年度累计量`
  - `01月数量`
  - `01月平均价格`
  - 后续月份同理
- Excel 中数量和价格列已设置为千位分隔，并保留两位小数。
- 已加入人工登录流程预留：
  - 有头 Chrome 打开页面；
  - 脚本自动填写账号；
  - 操作者在浏览器中手动输入密码、图形验证码、短信验证码并点击登录；
  - 脚本检测登录完成后继续后续数据获取。

## 使用方式

建议先创建并启用虚拟环境：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

安装 Python 依赖：

```powershell
pip install -r requirements.txt
```

首次使用 Playwright 时，需要安装浏览器运行环境：

```powershell
python -m playwright install chromium
```

运行脚本：

```powershell
python GEC_output_data.py
```

默认输出文件：

```text
GEC_2026_绿证流入流出前10.xlsx
```

如果需要访问登录后数据，可在脚本配置区将：

```python
LOGIN_REQUIRED = False
```

改为：

```python
LOGIN_REQUIRED = True
```

如需自动填充密码，可临时设置环境变量：

```powershell
$env:GEC_PASSWORD="你的密码"
python GEC_output_data.py
```

不建议将密码、cookie、登录态文件提交到 GitHub。

## 后续方向

- 根据登录后页面的数据位置，继续扩展数据抓取范围。
- 通过 Playwright 监听登录后页面请求，定位真实数据来源和查询参数。
- 将登录后数据按业务主题拆分为多个 Excel sheet。
- 为页面操作增加更完整的超时、等待、重试和日志机制。
- 将年份、Top N、输出文件名等参数改为命令行参数或配置文件。
- 增加数据校验逻辑，确保接口返回为空、字段缺失、月份缺失时有明确提示。

## 注意事项

- 网站存在验证码和短信验证，验证码相关步骤应由操作者在打开的 Chrome 页面中人工完成。
- 不建议高频请求或并发抓取，避免触发网站风控。
- `gec_storage_state.json`、日志、截图、Excel 输出文件属于本地运行产物，已通过 `.gitignore` 排除。
