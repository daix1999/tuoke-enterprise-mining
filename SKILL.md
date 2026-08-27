---
name: tuoke-enterprise-mining
slug: tuoke-enterprise-mining
displayName: 拓客企业挖掘
description: 通用 To B 拓客工具——通过关键词(经营范围)搜索去重，输出企业名单及联系方式；也支持按公司名直接搜索（文字单家/批量输入或喂表格名单），查完整电话邮箱。按用户身份(卖什么)+目标客户类型(卖给谁)匹配行业关键词，随机组合批量挖掘渠道商/集成商，自动补全电话邮箱。支持用户自主设定条件 + AI 自适应丰富需求，预留业务穿透等拓展能力。免费、无需 API Key。
version: 1.1.6
license: MIT
---

# 拓客企业挖掘

通用 To B 拓客工具：通过关键词（经营范围）搜索 → 去重 → 输出企业名单及联系方式。

数据源：风鸟（riskbird.com）高级搜索 API。搜索/查公司**免费不耗积分**，仅社保参保人数等详情耗积分。

## 核心能力

1. **身份 × 客户类型 → 关键词匹配**：先问身份(卖什么，如联想商用电脑/服务器/安防监控/软件/办公设备)，再问目标客户类型(渠道商/经销商 或 集成商/项目商)，自动匹配经营范围关键词
2. **随机关键词组合**：核心词必带 + 随机叠加细分词，每次跑一轮组合不同、结果不重复
3. **台账去重**：SQLite 三表(combos/runs/companies)，按企业名全局去重，串码识别组合
4. **自动补电话**：搜索名单后逐家查公司，免费提取完整电话邮箱（命中率约 98%）
5. **用户自主设定条件**：支持直接指定地区/注册资本/成立年限/联系方式/有无中标专利等
6. **AI 自适应丰富**：用户给模糊需求（"找卖电脑的小公司"），AI 自动补全中小型/存续/IT行业等条件
7. **限流保护**：检测风鸟 flagLimit 信号，额度用尽自动提示退出，不空跑
8. **按名搜索指定公司**：直接文字输入公司名（单家或批量），或喂 Excel/CSV/JSON/TXT 名单，逐家查完整电话邮箱/法人/地址/注册资本——无需先跑关键词挖掘

## 快速开始（一键跑完整流程）

```bash
python scripts/run_full.py \
  --db <数据目录>/ledger.db \
  --outdir <数据目录>/output \
  --identity "联想商用电脑" \
  --customer-type "渠道商/经销商" \
  --regcap "100￥500" \
  --count 50
```

一条命令串起「搜索 → 补电话 → 写回 → 分表导出」全程，输出精简摘要（省 token）。

不指定 identity/customer-type/regcap 时，自动随机组合。

## 5 分钟上手（三条入口怎么选）

| 你的场景 | 用哪个 | 一条命令 |
|---|---|---|
| **想挖一批新客户名单**（还没名单） | `run_full.py` 一键挖名单 | `python scripts/run_full.py --db 台账.db --outdir output --identity "联想商用电脑" --customer-type "渠道商/经销商" --count 50` |
| **手上已有名单，只缺电话**（表格/TXT/JSON） | `riskbird_batch.py` 喂文件 | `python scripts/riskbird_batch.py 名单.xlsx` |
| **就想查一两家公司**（比如客户发来一家） | `riskbird_batch.py` 输名字 | `python scripts/riskbird_batch.py --name "北京某科技有限公司"` |

- 三条路都**免费**（风鸟搜索/查公司不耗积分），输出 Excel 表格
- 不确定选哪条？先试第 3 条（输个公司名），体验最快
- 每条路的详细参数见下文对应章节

## 搜索指定公司信息（按名查 / 批量查）

两种用法：①**直接文字输入公司名**（单家或批量，无需准备文件）；②**喂表格/文本名单**（Excel/CSV/JSON/TXT）批量查。两种都逐家在风鸟查公司页，提取完整电话、邮箱、法人、地址、注册资本等。

### 用法一：直接文字输入公司名

```bash
# 单家搜索
python scripts/riskbird_batch.py --name "北京博维伟业有限公司"

# 批量搜索（可重复 --name，或等价用 --names）
python scripts/riskbird_batch.py --name "A公司" --name "B公司" --name "C公司"
python scripts/riskbird_batch.py --names "A公司" --names "B公司"

# 指定输出文件
python scripts/riskbird_batch.py --name "A公司" --out 结果.xlsx

# 输出格式：excel(默认) / json / json5 / jsonc / both
python scripts/riskbird_batch.py --name "A公司" --name "B公司" --format json
python scripts/riskbird_batch.py 名单.xlsx --format both
python scripts/riskbird_batch.py 名单.xlsx --format json5   # JSON5（可注释，需 pip install json5）
python scripts/riskbird_batch.py 名单.xlsx --format jsonc   # JSON with Comments（可注释，需 pip install json5）

# 自定义命名（占位符自动替换）
python scripts/riskbird_batch.py 名单.xlsx --name-format "渠道线索_[YYYY-MM-DD]"
python scripts/riskbird_batch.py 名单.xlsx --out "结果_[YYMMDD].xlsx"
```

### 用法二：喂名单文件（表格/文本批量查）

```bash
# Excel / CSV / JSON（自动识别表头“企业名/公司名/名称”列）
python scripts/riskbird_batch.py 名单.xlsx
python scripts/riskbird_batch.py 名单.csv
python scripts/riskbird_batch.py 名单.json
python scripts/riskbird_batch.py 名单.json5   # JSON5 名单（支持注释/单引号/尾逗号）
python scripts/riskbird_batch.py 名单.jsonc   # JSONC 名单（支持注释）

# TXT：每行一个公司名
python scripts/riskbird_batch.py 名单.txt

# 自定义输出 + 仅抽取校验（不查风鸟）
python scripts/riskbird_batch.py 名单.xlsx --out 结果.xlsx
python scripts/riskbird_batch.py 名单.xlsx --dry
```

### 说明

- **输入**：`.xlsx` / `.csv`（UTF-8 或 GBK 自动探测）/ `.json`（字符串数组）/ `.json5` / `.jsonc`（支持注释）/ `.txt`（每行一个公司名）。Excel/CSV 会自动识别“企业名/公司名/名称/company”等表头，取对应列；多工作表跨表按公司名去重；无表头时取第一列。
- **输出格式（`--format`）**：`excel`（默认，仅表格）/ `json`（仅结构化 JSON，列表含全部字段，便于二次处理、入库）/ `json5`（输出 `.json5`，JSON5 超集可注释，人可手工编辑，读取兼容）/ `jsonc`（输出 `.jsonc`，JSON with Comments，仅比标准 JSON 多注释，VS Code 生态原生兼容）/ `both`（Excel+JSON 都出）。无论选哪种，都会额外生成 `<基名>_state.json` 作为断点续跑内部状态文件。`json5`/`jsonc` 格式需要 `pip install json5`（其余格式不需要）；读取一律兼容三种格式，手改过带注释的文件也能续跑。
- **默认命名**：文件输入 → `名单_联系方式.xlsx`(+`_state.json`)；文字直输 → `搜索结果_联系方式.xlsx`(+`_state.json`)。Excel 输出列：序号|企业名|注册地点|法人|电话|邮箱|规模|注册资本|**公司状态**|来源表|经营范围（公司状态=在营/正常/注销/吊销，打电话前先筛）。
- **自定义命名**：`--name-format "<模板>"` 或 `--out "<路径>"` 均可，支持占位符自动替换：
  `[YYYY-MM-DD]` `[YYYYMMDD]` `[YYMMDD]` `[YYYY]` `[MM]` `[DD]` `[HHMMSS]` `[TS]` `[NAME]` `[COUNT]`
  （`[NAME]`=来源文件名/“搜索结果”，`[COUNT]`=企业数量，其余为日期时间片段）。示例：`--name-format "渠道线索_[YYYY-MM-DD]_[COUNT]"` → `渠道线索_2026-08-26_50.xlsx`。
- **断点续跑**：同名 `_state.json` 已存在的记录会跳过（含 NOT_FOUND 确定结果），只查未完成的；失败/中断后可重跑不重复消耗额度；错误记录按公司名替换不重复膨胀。
- 每 10 家打印一次进度；命中额度上限/需登录自动停止并提示。

## 输出样例（跑完是什么样）

**① 单家查询**（`--name` 只输一个名字）——直接在对话流打印，不生成文件：

```
📋 查询结果：北京正群欣世信息技术有限公司
────────────────────────────────────────────────
  电话：15810730150
  邮箱：limei@zqxsinfo.com
  法人：高巍
  注册资本：3000.0万元人民币
  成立日期：1999-07-26
  经营状态：在营
  统一信用代码：911101087177139649
  地址：北京市海淀区中关村东路66号世纪科贸大厦B座2706
```

**② 批量查询**（喂文件或多个名字）——生成 Excel 表格（同时自动生成 `_state.json` 用于断点续跑）：

| 序号 | 企业名 | 注册地点 | 法人 | 电话 | 邮箱 | 规模 | 注册资本 | **公司状态** | 来源表 | 经营范围 |
|---|---|---|---|---|---|---|---|---|---|---|
| 1 | 北京明通盈动科技有限公司 | 北京市海淀区苏州街33号1116室 | 邓金华 | 13911381126 | mahongyan1972@sina.com | - | 50.0万元人民币 | **注销** | 名单.xlsx | - |
| 2 | 北京东泰通信息技术有限公司 | 北京市通州区西集镇企业发展服务中心7625号 | 曹春娟 | 13366317117 | 787231647@qq.com | - | 200.0万元人民币 | **在营** | 名单.xlsx | - |

> **公司状态列怎么用**：`在营`/`正常` = 可以打电话；`注销`/`吊销` = 公司已停业，跳过；`未知` = 页面未展示，可人工确认。批量名单里常有 20%~30% 已注销/吊销，**先筛状态再打电话**能省大量无效外呼。

**③ 批量进度条**（每 10 家打印一次）：

```
✅ 已完成: 187 | ⏳ 待查询: 5
  进度 10/194 | 成功 9 | 失败 1
  ⛔ 额度/登录异常（NEED_LOGIN），停止后续查询   ← 出现即停，修复后重跑续跑
完成！成功: 189 | 失败: 5 | 总计: 194
```

## 脚本清单

| 脚本 | 作用 |
|---|---|
| `run_full.py` | **一键入口**：搜→补电话→写回→分表 |
| `riskbird_random_mining.py` | 核心引擎：combo(随机组合)/mine(搜索入库)/export(分表)/stats(统计) |
| `riskbird_batch.py` | 按名搜索：给定公司名(文件/文字直输)或名单，逐家查完整电话邮箱/法人/地址/注册资本 |
| `industry_identities.json` | 身份×客户类型→关键词配置（可自行增删行业） |

## 关键参数（mine / run_full 通用）

| 参数 | 说明 | 示例 |
|---|---|---|
| `--identity` | 用户身份(卖什么) | 联想商用电脑、服务器/数据中心、安防监控、软件/SaaS、办公设备/通信 |
| `--customer-type` | 目标客户类型 | 渠道商/经销商、集成商/项目商 |
| `--scope` | 经营范围(覆盖自动) | 计算机销售+系统集成（加号=同时包含） |
| `--region` | 地区 | 北京、上海、深圳… |
| `--regcap` | 注册资本区间(万元) | 100￥500（全角￥分隔） |
| `--esdate` | 成立年限 | 5￥3（3-5年） |
| `--contact` | 联系方式 | has_lianxi_phone(有手机) |
| `--has-bid-win` | 只搜有中标记录的 | 布尔开关 |
| `--has-patent` | 只搜有专利的 | 布尔开关 |
| `--has-soft-copyright` | 只搜有软著的 | 布尔开关 |

## 前置条件

1. **风鸟登录态（一次性）**：脚本通过 agent-browser 操作已登录的风鸟页面取数据，需要先登录一次风鸟，登录态会持久化，之后不用重复登录：
   ```bash
   agent-browser --profile <固定目录> open https://www.riskbird.com/
   ```
   在弹出的浏览器窗口登录你的风鸟账号（搜索/查公司免费不耗积分）。
   > 为什么需要 agent-browser？风鸟的搜索接口依赖浏览器登录态 cookie，脚本用 agent-browser 打开已登录会话来查询，这样免费且稳定。**没有 agent-browser 也能用**：手动用 Chrome 打开风鸟登录，只要登录态有效，原理一样。
2. Python 环境：`requests` 库（脚本会自动从 agent-browser 会话拿登录态 cookie）；`json5` 库仅在用 `--format json5/jsonc` 时需要。

## 关键机制说明

- **经营范围多词**：加号 `+` / 空格 = 同时包含(AND)；逗号 `,` = 任一包含(OR)
- **企业状态**：默认 `status=1`（在营/存续），过滤注销吊销停业
- **注册资本默认中小型**：上限 1000 万，过滤巨头和空壳
- **限流信号**：搜索返回 `flagLimit=="1"` 表示当日额度用尽，脚本自动检测退出
- **输出格式**：Excel，字段=序号|企业名|注册地点|法人|电话|邮箱|规模|注册资本|公司状态|来源表|经营范围（公司状态：在营/正常/注销/吊销，打电话前先筛）；每次跑一轮一份独立分表（文件名=关键词_日期_串码.xlsx），不合并总表

## 常见问题 FAQ

**Q1. 提示「风鸟未登录或会话已过期（NEED_LOGIN）」怎么办？**
登录态不是永久的，过一阵子会失效。重新登录一次即可（见上文"前置条件"），登录态持久化后重跑同一条命令，已查完的公司会跳过，只补剩下的。

**Q2. 提示「未匹配到该企业（NOT_FOUND）」是公司不存在吗？**
不一定是。①先核对公司名是否与**工商登记全称**一致（漏字/多字/简称都会查不到）；②该企业可能已更名或查无登记；③脚本已对网络波动自动重试 2 次，若仍无结果可稍后再跑（风鸟搜索页是异步加载，偶发抓不到）。

**Q3. 批量跑一半断了/报错了怎么办？**
直接重跑同一条命令即可。脚本会自动读 `_state.json`，已完成的公司跳过，只查剩下的——**不重复消耗额度**。中途遇到 NEED_LOGIN/额度上限会停止，修好后重跑续跑。

**Q4. 为什么有的公司没电话/邮箱？**
风鸟页面上没展示该企业电话（企业未公开），或该企业已注销/吊销（这类直接跳过即可）。这是数据源本身的情况，不是查询失败。

**Q5. 名单里有「XX电脑超市」「XX销售部」这类名字查不到？**
脚本按"像公司名"的正则识别名单列，词表未覆盖"超市/销售部"等词时会过滤掉。**解法：把名单存成 TXT（每行一个公司名）喂入**，TXT 输入不做过滤，全部查询。

**Q6. 提示「daemon already running --profile ignored」？**
这是 agent-browser 已有一个旧 daemon 在跑且没带 profile。先 `agent-browser close` 把旧 daemon 杀掉，再带 `--profile` 重开；若提示连接被拒，是杀进程后锁未释放，再开一次即可。

**Q7. 登录二维码/弹窗打不开？**
风鸟登录弹窗只在**页面首次加载**时自动弹出，顶栏"登录/注册"链接是装饰性的点了没反应。用 `agent-browser open https://www.riskbird.com/` **重新加载页面**即可重新触发弹窗；二维码约 1-2 分钟过期，过期就重载再截。

**Q8. 提示「json5 库未安装」？**
`--format json5/jsonc` 需要 `pip install json5`；只用默认 `excel` 或 `json` 格式不需要。

**Q9. 输出文件在哪？字段什么意思？**
默认在**当前命令行所在目录**（或 `--out` 指定路径），命名规则见"说明"。Excel 11 列：序号|企业名|注册地点|法人|电话|邮箱|规模|注册资本|公司状态|来源表|经营范围；同时生成 `_state.json` 是断点续跑的内部状态文件（别删，删了重跑会重新全部查询）。

**Q10. 结果表里「公司状态=注销/吊销」的要不要打？**
不要。注销/吊销=公司已停业，电话多半打不通或已转他人。**先按公司状态列筛出「在营/正常」，再开始外呼**——批量名单里通常有 20%~30% 已注销吊销。

## 参考文档

- `references/风鸟高级搜索能力清单.md`：风鸟高级搜索 27 个筛选维度的字段名、取值、语义（实测逆向）
- `references/拓客条件自适应丰富规则.md`：AI 根据用户模糊需求自动补全搜索条件的规则

## 拓展路线

当前聚焦「关键词搜索 → 去重 → 企业名单及联系方式」，架构上预留以下能力扩展位：

1. **企业业务穿透**：对名单内企业做实际主营业务调查（中标/招投标数据反推实际业务，判断是否为联想/华为等品牌渠道）
2. **意向评分**：结合注册资本、中标记录、参保人数、合作动态等维度给企业打分，优先跟进高意向客户
3. **多数据源扩展**：当前数据源为风鸟，可扩展爱企查/企查查等（保持统一的 `build_cond` 条件模型）
4. **客户跟进管理**：名单输出后联动 CRM/台账，记录触达状态、跟进记录

## 致谢

基于 `riskbird-cominfo-batch`（陆凌燕律师，MIT License）改造扩展，原补电话能力保留，新增身份匹配、随机组合、台账去重、自主设条件、AI 辅助等拓客能力。
