# GEC Market Analysis

本项目用于通过 Python + Playwright 获取广州电力交易中心绿电绿证交易平台中的绿证市场数据，并整理为 Excel 报表。

## 当前功能

- 获取首页公开数据：
  - 2026 年全国绿证流出量前 10 省份；
  - 2026 年全国绿证流入量前 10 省份；
  - 各省各月数量和平均价格。
- 获取登录后的挂牌绿证产品数据：
  - 项目业主；
  - 电源类型；
  - 电力生产年月；
  - 补贴类型；
  - 消纳方式；
  - 可交易数量；
  - 已出售数量；
  - 单价。
- 支持三种运行模式：
  - `1`：仅获取流入/流出前十统计，不需要登录；
  - `2`：获取当前全部挂牌绿证统计，登录后全量遍历并重建 cache；
  - `3`：轻量化获取，登录后使用 cache 锚点提前停止翻页。
- 统一导出一份 Excel：
  - `流出前10`
  - `流入前10`
  - `挂牌产品`
- Excel 文件默认保存到当前电脑用户桌面，文件名带时间戳：

```text
GEC绿证市场情况_YYYYMMDD_HHMMSS.xlsx
```

## 安装依赖

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

## 运行方式

交互式运行：

```powershell
python -u GEC_output_data.py
```

脚本会提示选择运行模式：

```text
1. 仅获取流入/流出前十统计
2. 获取当前全部挂牌绿证统计
3. 轻量化获取
```

也可以直接指定模式：

```powershell
python -u GEC_output_data.py 1
python -u GEC_output_data.py 2
python -u GEC_output_data.py 3
```

## 登录说明

模式 `2` 和 `3` 需要登录。脚本会打开有头 Chrome，并自动填写账号和密码；图形验证码、短信验证码和点击登录仍由操作者在浏览器中完成。

账号和密码读取顺序：

```text
环境变量 GEC_USERNAME / GEC_PASSWORD
local_config.py
默认账号 / 手动输入
```

可在本机创建 `local_config.py`：

```python
GEC_USERNAME = "你的账号"
GEC_PASSWORD = "你的密码"
```

`local_config.py` 已加入 `.gitignore`，不会提交到 GitHub。

## Cache 机制

挂牌产品数据会写入本地 cache：

```text
cache/listed_products_cache.json
```

轻量化模式会使用 cache 中的历史产品作为锚点：

- cache 保存上次结果中前 100 个产品 key；
- 本次运行命中其中任意 20 个后停止继续翻页；
- 新出现的产品在 Excel 中标记为 `本次新增`；
- 旧产品标记为 `上次抓取结果`。

产品 key 使用：

```text
项目业主
电源类型
电力生产年月
补贴类型
消纳方式
单价
```

## 风控与稳定性

- 页面导航 timeout 已设置为 5 分钟，适应较慢网络。
- 挂牌产品页会优先切换到 `48条/页`，减少翻页次数。
- 翻页后会等待页面内容或页码实际变化，避免页面卡住时误读旧数据。
- 翻页间隔加入随机等待，降低访问频率。
- 检测到“访问过于频繁”提示时，会暂停后再继续。

## 注意事项

- 抓取挂牌产品时不要最小化 Chrome，最小化可能导致页面渲染或翻页变慢。
- 不建议高频运行或并发抓取，避免触发网站风控。
- 运行产物和敏感文件不会提交到 GitHub，包括：
  - Excel 输出；
  - `cache/`；
  - `logs/`；
  - `screenshots/`；
  - `local_config.py`；
  - `gec_storage_state.json`。
