# OTA CMS 自动化（Playwright）

用 Playwright 驱动真实浏览器，自动完成两类工作：

1. **CMS OTA 升级规则工作流**（`cms.intsmarthub.com`）：导入 Excel → 创建规则 → 回收 Rule ID → 逐条上线。
2. **DASS 安卓升级状态填报**（FineReport 报表站点）：把更新后的 Excel 填报/更新到 DASS，并在升级前做**升级成功率闸门**校验。

> 本项目与 `ota_v4`（生成 Excel 的 Web 工具）**完全独立**。衔接点只有一个：ota_v4 生成的
> `.xlsx`，作为本项目的输入。

## 为什么是 Playwright（UI 自动化），不是 API

目标系统**没有可用的后端 API**，导入/创建规则/上线/填报全是网页 GUI 操作。所以用 Playwright
模拟真实用户点击、填表、上传文件。DASS 是 FineReport 报表，单元格是非标准控件、查询后会话锚会变，
所有选择器与机制均来自只读探测（见 `.agents/tasks/` 下的 findings 文档）。

## 安装与配置（首次）

```bash
python3 -m venv .venv
./.venv/bin/python -m pip install -r requirements.txt
./.venv/bin/python -m playwright install chromium
# 从示例复制并填入真实凭据
cp config/.env.example config/.env
# 在 config/.env 中把 DASS_USERNAME / DASS_PASSWORD（及 CMS_*）从占位符 __set_me__ 改为真实值
```

- 凭据**绝不写死在代码里**，只从 `config/.env` 读取；该文件已在 `.gitignore` 中，不会被提交。
- `CMS_HEADED`（默认 `true`）控制是否开可见浏览器窗口。填报/上线的人工确认环节需要有头模式。

---

## 一、DASS 填报与成功率闸门：`./run_dass.sh`（主命令）

把更新后的 Excel 填进 DASS「安卓升级状态填报」。参数 `[OPTION] [Excel路径]`，两者都可选、顺序不限。

```bash
# init 模式（默认）：读 Excel → 插入 N 行 → 逐格填入
./run_dass.sh                          # 取 ota_v4 输出目录里最新的 .xlsx
./run_dass.sh "/路径/文件.xlsx"          # 指定 Excel

# update 模式：匹配已存在行并更新『当前阶段』『当前阶段时间』
./run_dass.sh ota50   "/路径/文件.xlsx"
./run_dass.sh ota200  "/路径/文件.xlsx"
./run_dass.sh ota1k   "/路径/文件.xlsx"
./run_dass.sh ota5k   "/路径/文件.xlsx"
./run_dass.sh ota20k  "/路径/文件.xlsx"
./run_dass.sh ota-all "/路径/文件.xlsx"
```

**OPTION 取值**：`init`（默认）/ `ota50` / `ota200` / `ota1k` / `ota5k` / `ota20k` / `ota-all`

- `init`：插入 N 行并逐格填入。
- `otaXX / ota-all`：不插入行，用「粗筛批量匹配 + 精准回退」两阶段定位已存在行并更新。

### 执行流程

1. **升级成功率闸门**（见下文「二」）——仅 `ota200/ota1k/ota5k/ota20k/ota-all` 触发；`ota50/init` 跳过。
   未通过则直接中止，不填报、不提交。
2. 通过闸门后 → 按模式填入/更新 → 截图并**暂停**（截图存 `artifacts/`）。
   - 填入/更新过程中，**每一条记录都会逐条打印写入明细**（行号、Rule ID、方案号、来源阶段、写入的
     当前阶段/当前阶段时间等），便于事后查对。
3. 终端提示时输入 `approve`（或 `提交` / `yes`）才执行「验证 → 提交」；输入其它内容则取消，不提交。
   「上线/提交」是高影响操作，脚本**绝不无人值守直接发布**。

> `ota50` 是首档、无上一阶段，所以不触发成功率闸门；`ota200` 及之后每一档，升级前都会先过闸门。

---

## 二、升级成功率闸门

从 `ota200` 档位起，在更新 DASS 之前自动校验本批每个 Rule ID 的**升级成功率**，不达标就中止整个事务。

### 规则

- **触发档位**：`ota200 / ota1k / ota5k / ota20k / ota-all`。`ota50` 与 `init` **不校验**。
- **阈值**：本批**所有** Rule ID 的升级成功率都 **≥ 50%** 才放行升级。
- **中止（abort）**：只要有任一 Rule ID 成功率 < 50%，**或**两阶段查询都查不到该 Rule ID，
  就中止——不执行 update、不进入人工提交环节，进程返回码为 `2`。

### 成功率从哪来

数据来自 DASS 的 **「安卓升级展示报表」** 报表页。该页有现成的百分比列 **「升级成功率」**（如 `49.09%`），
Rule ID 列为 **「升级策略ruleid」**。脚本直接解析该百分比，不自行用成功/失败设备数计算。

### 查询方式（两阶段）

- **Stage1（三字段查询）**：按 **方案号 + 目标版本 + 当前阶段** 查询，再在结果中按 Rule ID 核对成功率。
  其中「当前阶段」= 记录升级前在 DASS 的**实际阶段 = 上一阶段**：

  | 升级档位 | 查询用的当前阶段（上一阶段） |
  |---|---|
  | `ota200`  | `OTA50` |
  | `ota1k`   | `OTA200` |
  | `ota5k`   | `OTA1000` |
  | `ota20k`  | `OTA5000` |
  | `ota-all` | `OTA20000` |

- **Stage2（精准回退）**：对 Stage1 未命中的 Rule ID，再用 **方案号 + 目标版本 + 设备扩展信息 + 当前阶段**
  精准查询后按 Rule ID 定位。

### 日志

闸门校验时，**每个 Rule ID 都会逐条打印**成功率、是否 ≥ 50%（通过/不通过）、来源阶段（Stage1/Stage2），
末尾给出通过/不通过/未找到的汇总与最终放行/中止结论。

### 页面机制要点（来自只读探测）

- 报表页**没有「显示行数 LIMIT」框**；结果是**分页**的，脚本逐页遍历（判停于本页 0 行或「下一页」禁用，
  不盲信页码指示 `/N`）。
- 「方案号」是合并单元格，行以「升级策略ruleid」列为锚解析。
- 「升级成功率」「升级策略ruleid」两列需**横向滚动**才会渲染，脚本会滚动到位后再读取。
- Rule ID 两侧比较前都做**归一化**：去掉所有分隔符（半角 `,` / 全角 `，` / 空格等），
  使 `65,196` 与 `65196` 视为同一个 ID。

---

## 三、独立查成功率（只读）：`dass_perf_query`

不触发任何升级，单独查某个方案号+目标版本下指定 Rule ID 的升级成功率。

```bash
# 查指定 Rule ID（区间 65204-65212，或逗号 65204,65210，可混用）
CMS_HEADED=true ./.venv/bin/python -m automation.dass_perf_query \
    --plancode "MT9655AndrGTVLas" --target "0100UQ1015" "65204-65212"

# 列出该方案号+目标版本下的全部 Rule ID 及成功率
CMS_HEADED=true ./.venv/bin/python -m automation.dass_perf_query \
    --plancode "MT9655AndrGTVLas" --target "0100UQ1015" --all
```

- `--plancode` = 方案号（设备类型代码）、`--target` = 目标版本、位置参数 = Rule ID（区间/逗号）、`--all` = 列全部。
- 逐行打印每个 Rule ID 的成功率与是否 ≥ 50%；存在 < 50% 的行时返回码为 `1`（正常，表示这批升级会被闸门拦下）。
- **纯只读**：绝不点升级/提交/插入行。

---

## 四、CMS OTA 升级规则工作流：`./run_login.sh`

CMS 侧的登录与 OTA 导入/创建/上线流程脚本入口（见 `automation/login.py`、`automation/inspect_*.py`）。

```bash
./run_login.sh
```

上线前同样**暂停并等待人工确认 Rule ID**，不做无人值守的直接发布。

---

## 五、CMS active_ota：导入建规则 → 上线 → 汇总：`./run_active_ota.sh`

CMS 侧的 `active_ota` 模块（`automation/active_ota.py`）把「导入 OTA 模板建规则 → 逐条上线并处理冲突 → 回读状态汇总」串成三阶段一键流程。

```bash
./run_active_ota.sh                      # 取 ota_v4 输出目录里最新的 .xlsx 作为导入模板
./run_active_ota.sh "/路径/模板.xlsx"      # 指定导入模板 Excel
./run_active_ota.sh --probe-online        # 只读 dump 编排视图选择器（不做任何写操作）
./run_active_ota.sh --selftest            # 纯逻辑自测（不连浏览器）
```

三阶段：

- **阶段1 导入建规则**：走「OTA升级模板导入」入口，把规则名称/备注填入导入弹窗后导入，再回读本批连续的 Rule ID range。
  - 规则名称 = `设备类型代码 + "_" + 目标版本最后 5 个字符`。例 `MT9655AMAndrGTV` + `0100UQ1015`(末5=`Q1015`) → `MT9655AMAndrGTV_Q1015`。
  - 备注 = `今天日期YYYYMMDD + "_USA_" + <导入行数>`。例 `20261102_USA_9`。
  - 设备类型代码/目标版本从 Excel 读，假设单一值；混多个 → 报错中止提示，不猜。
  - 模板默认取 ota_v4 输出目录最新 `.xlsx`，也支持命令行指定路径。
- **阶段2 逐条上线 + 冲突处理**：全自动（不加人工 approve），阶段2 开始前打印待上线 Rule ID 清单。
  - 在「待选规则」按 Rule ID 搜出 → 点「上线」→「确认上线该规则?」点「确定」。
  - 若弹「规则替换」冲突窗（文案「当前规则和 XXXXX 冲突,是否替换?」）：解析被冲突号 XXXXX → **先不点弹窗按钮** → 去「已选规则」按 XXXXX 搜、读其 mac组 列：mac组**非空** → 回弹窗点「确定」替换上线；mac组**为空/判定不确定** → 点「取消」中止该条并记录。
- **阶段3 状态汇总**：对每个 Rule ID 回读最终状态 = `online`（已选规则能搜到、操作列为「下线」）/ `冲突中止`（含被冲突号与 mac组 判定）/ `失败`。全部 online 退出码 0，否则非 0。

安全边界：

- 所有写操作（导入确定、上线确定、替换确定）都封装成独立的二次确认 gate，条件不满足只打印不点，绝不误点「删除/替换」等危险按钮。
- mac组 判定但凡不确定（搜不到被冲突规则、读不到 mac组 列、列文本歧义）一律**保守中止该条并记录**，不猜着替换。
- Rule ID 比较统一复用 `dass_fill.norm_ruleid`（去所有非数字分隔符）。
- ⚠ 编排视图（「已选规则/待选规则」双区、上线/下线/编辑按钮、确认框、规则替换弹窗、mac组 列）的真实选择器在只读探测阶段**未实跑复现**（见 `.agents/tasks/active-ota-probe-findings.md` 第 2.2/3/5 节），这些选择器为**预期值**，首次实跑前请先 `--probe-online` 校正。
- 端到端写流程（真导入/上线/替换）改线上生产数据、不可逆，请在知情下运行 `run_active_ota.sh`。

---

## 测试与验证

- 核心模块带离线自测（不连浏览器）：
  ```bash
  ./.venv/bin/python -m automation.dass_perf --selftest
  ./.venv/bin/python -m automation.dass_perf_query --selftest
  ./.venv/bin/python -m automation.dass_fill --selftest
  ./.venv/bin/python -m automation.active_ota --selftest
  ```
  覆盖：成功率阈值判定（含等于 50%）、abort 返回码、Rule ID 去分隔符对齐、区间/逗号解析、日志格式；
  以及 active_ota 的 D1 名称拼接 / D2 备注拼接 / D3 多值校验 / D7 冲突号解析 + mac组 空判定 / D10 汇总三态格式。
- 真实页面行为（三字段查询、翻页、逐条日志）需连 DASS 实跑确认。只读的成功率查询可用
  `dass_perf_query` 安全验证；涉及写入的 `run_dass.sh` 升级，在人工 `approve` 之前不会真正提交。

## 安全与边界（重要）

- 凭据只从 `config/.env` 读取，不入库。
- 「上线 / 提交」是高影响操作：脚本在提交前**暂停并等待人工确认**，不无人值守发布。
- 升级成功率不达标时**自动中止**升级事务（返回码 2），避免把低成功率的批次推上线。
- 每个关键步骤自动**截图**到 `artifacts/`，便于核对和排错。
- 请确认自动化该内部系统符合公司规范。

## 目录结构

```
ota_cms_automation/
  automation/
    dass_login.py       # DASS 登录 + 导航
    dass_fill.py        # DASS 填报/更新主脚本（init / otaXX），含成功率闸门接入
    dass_perf.py        # 升级成功率闸门：查询 + 校验 + 逐条日志
    dass_perf_query.py  # 独立只读的成功率查询 CLI
    dass_data.py        # 读取更新后的 Excel → 填报字段
    active_ota.py       # CMS 导入建规则 → 逐条上线并处理冲突 → 状态汇总
    login.py / inspect_*.py  # CMS 侧登录与页面探测
  config/               # .env.example、运行配置
  artifacts/            # 运行时截图
  .agents/tasks/        # 只读探测结论（findings）与评审记录
  run_dass.sh           # DASS 填报/更新入口
  run_active_ota.sh     # CMS active_ota 导入/上线/汇总入口
  run_login.sh          # CMS 登录入口
  requirements.txt
```
