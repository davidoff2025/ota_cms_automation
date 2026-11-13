# OTA 部署自动化使用手册

用 Playwright 驱动真实浏览器，把海外安卓 OTA 部署的重复操作串成一个命令 `./deployota`：

```
OTA 申请模板(Excel)
   │  prep      生成标准 Excel → 严格校验 → FTP 下载升级包 → 上传 S3（地址换成 ota-dl 链接）
   ▼
标准 Excel
   │  publish   CMS 导入建规则 → 拿 Rule ID(回写 Excel) → 逐条上线 → DASS init 填报   ← 4 个检查点
   ▼
   │  update    后续放量：ota50 → ota200 → ota1k → ota5k → ota20k → ota-all，按 Rule ID 更新 DASS（ota200 起先过成功率闸门）
   ▼
   │  history   查看部署历史（每次 dass 更新后自动入库）
```

目标系统没有可用的后端 API，导入、建规则、上线、填报全是网页操作，所以用 Playwright 模拟真实用户。
本项目与 `ota_v4` 完全独立：`prep` 用的生成引擎已复制到 `automation/otagen/`，不依赖 `ota_v4`。

---

## 目录

1. [快速开始](#一快速开始)
2. [一次完整部署怎么跑](#二一次完整部署怎么跑)
3. [命令速查](#三命令速查)
4. [各命令详解](#四各命令详解)
5. [配置与环境变量](#五配置与环境变量)
6. [映射表（设备信息）](#六映射表设备信息)
7. [退出码](#七退出码)
8. [常见问题与排错](#八常见问题与排错)
9. [安全边界](#九安全边界)
10. [测试与目录结构](#十测试与目录结构)

---

## 一、快速开始

首次安装：

```bash
python3 -m venv .venv
./.venv/bin/python -m pip install -r requirements.txt
./.venv/bin/python -m playwright install chromium
cp config/.env.example config/.env      # 然后编辑 config/.env 填入真实值
```

`config/.env` 至少要填：

| 用途 | 变量 |
|---|---|
| CMS 登录（publish / online / offline / update） | `CMS_USERNAME`、`CMS_PASSWORD` |
| DASS 登录（dass / publish 的第 4 步） | `DASS_USERNAME`、`DASS_PASSWORD` |
| 下载 FTP 升级包（prep 真实上传） | `OTA_FTP_USER`（默认 ftpuser）、`OTA_FTP_PASS` |
| 部署记录的负责人（可选） | `OTA_OWNER` |

- 凭据只从 `config/.env` 读取，该文件已在 `.gitignore`，不会被提交。
- AWS 区域和密钥不写在 `.env`，走 boto3 默认链（`~/.aws`、实例角色或环境变量）。
- `CMS_HEADED=true`（默认）开可见浏览器。DASS 的人工核对环节需要有头模式。

确认安装正常：`./deployota --help`。

---

## 二、一次完整部署怎么跑

```bash
# 第 1 步：把申请模板变成标准 Excel，并把升级包传到 S3
./deployota prep --file '/路径/海外安卓OTA申请模板.xlsx'

# 第 2 步：CMS 导入 + 取 Rule ID + 上线 + DASS init 填报（一条命令，4 个检查点）
./deployota publish

# 第 3 步：之后按放量节奏，逐档更新 DASS——只要给 Rule ID（或区间、逗号分隔）
./deployota update ota50  65370-65379
./deployota update ota200 65370-65379      # ota200 起会先过成功率闸门
./deployota update ota1k  65370-65379
./deployota update ota5k  65370-65379
./deployota update ota20k 65370-65379
./deployota update ota-all 65370-65379

# 随时查看部署历史
./deployota history
```

要点：

- **不用每步都写 Excel 路径**。`prep` 通过校验后会记下它生成的那份 Excel（`.last_prep.json`），
  之后 `online` / `publish` / `dass` 不带 Excel 参数时自动用它，并打印用的是哪个文件。
  想用别的文件就在命令末尾写路径。
- **第一次建议用只有 1 条数据的测试 Excel 小批量试跑**，确认无误再整批。
- 第 3 步放量不再需要 Excel：`update` 按 Rule ID 去部署库取方案号、目标版本、设备扩展信息，再去 DASS 更新。

---

## 三、命令速查

| 命令 | 作用 | 写线上数据？ |
|---|---|---|
| `./deployota prep` | 生成 Excel → 严格校验 → FTP 下载 → 上传 S3 | 是（S3，`--no-upload` 则否） |
| `./deployota publish` | CMS 导入 + 回收 Rule ID + 上线 + dass init，4 个检查点 | 是（CMS 导入/上线，DASS 填报） |
| `./deployota online [Rule ID]` | 逐条上线 + 冲突处理 + 汇总（Excel，或直接给 Rule ID） | 是（CMS 上线/替换） |
| `./deployota offline <Rule ID>` | 按 Rule ID 下线（在线的直接下线，已下线/找不到的跳过） | 是（CMS 下线，需输入 yes） |
| `./deployota update <OPTION> <Rule ID>` | 按 Rule ID 更新 DASS 的 OTA 阶段，含成功率闸门 | 是（DASS，提交需你输入 approve） |
| `./deployota dass [OPTION]` | DASS 填报（init 插入新行）；也保留按 Excel 更新的旧用法 | 是（DASS，提交需你输入 approve） |
| `./deployota query <ID>` | 按 Rule ID 查部署信息（单个 / 区间 / 逗号） | 否 |
| `./deployota history` | 查看部署历史 | 否 |

每个子命令都支持 `--help`：`./deployota <子命令> --help`。

`update` / `online` / `offline` 结束时（成功、失败或中止）都会打印**操作总结**：逐步列出做了什么、结果如何，
最后一句结论（全部完成 / 部分完成 / 未完成：卡在哪一步）。例如：

```
======== 操作总结：update ota50 ========
  Rule ID：65345~65348（4 个）
  1. ✓ 从部署库解析 Rule ID：4 个取到方案号/目标版本/设备扩展信息
  2. · 成功率闸门（每个 Rule ID ≥ 50%）：ota50 不检查成功率
  3. ✓ CMS 灰度放开到 50阶段：放开 4 个（测试阶段 → 50阶段）
  4. ✓ 灰度阶段记录：4 个 → 50
  5. ✓ DASS 更新到 OTA50：写入 4 个（共 4 个）
  6. ✓ DASS 提交：已提交
  结论：✓ 全部完成
```

`update` 的操作总结最后生成一句可以直接转发的**同意语句**，后面跟着每个 Rule ID 的升级成功率：

```
  同意（OTA200，20261113，50032~50033）50032:83%，50033:70%
```

- 阶段写法：OTA50 / OTA200 / OTA1000 / OTA5000 / OTA20000 / OTA完全放开；日期为当天。
- Rule ID 为本次真正升级成功（CMS 放开 + DASS 写入）的那些，部分成功时只列成功的。
- 成功率：ota200 起取自成功率闸门（上一阶段的值）；ota50 不过闸门，只读查一次报表。按报表原值，只去掉末尾的 0
  （83.00% → 83%），不四舍五入；报表里没有的写「未查到」。
- DASS 还没提交时这句标黄，并提示「提交完成后再发」。

`publish` 结束时打印的是它自己的检查点报告。

常用写法：

```bash
./deployota prep --file x.xlsx --no-upload        # 只生成+校验，不碰 FTP/S3（演练用）
./deployota publish --no-dass                     # 只做 CMS 三步，不跑 DASS
./deployota publish '/路径/指定.xlsx'               # 指定 Excel
./deployota update ota200 65370,65375       # 按 Rule ID 更新
./deployota query 65370-65379,65385         # 查这些 Rule ID 属于哪个批次
./deployota history --device MT9655 --limit 50
./deployota history --owner devinw --export
```

---

## 四、各命令详解

### 4.1 `prep`：生成 Excel → 严格校验 → 上传 S3

```bash
./deployota prep --file '申请模板.xlsx'                      # 常用
./deployota prep --file 文本.txt --device-code MT9655AMAndrGTV
cat 文本.txt | ./deployota prep --device-code MT9655AMAndrGTV   # 从 stdin 读文本
```

**输入**：`--file` 可以是

- `.xlsx` / `.xlsm`：OTA 申请模板 Excel，直接解析（设备类型代码和 Feature Code 从表里的列读取）；
- `.txt`：粘贴的升级信息文本；
- 省略 `--file`：从 stdin 读文本。

旧版 `.xls` 不支持（请在 Excel 里另存为 `.xlsx`）；`~$` 开头的是 Excel 的临时锁文件，不是真正的表格，会被拒绝。

**参数**

| 参数 | 说明 |
|---|---|
| `--file FILE` | 输入文件，见上 |
| `--device-code` / `--device-name` / `--feature-code` / `--mac-group` | 非空时覆盖从输入解析出的值 |
| `--out-dir DIR` | 工作目录（生成的 Excel 和去重索引放这里） |
| `--no-upload` | 只生成+校验，打印 S3 计划和 ftp zip 清单，不碰 FTP/S3 |
| `--yes`, `-y` | 设备信息与映射表不一致时自动继续（非交互场景用） |
| `--force` | 强制重新生成并全量重传（默认复用相同内容的已有 Excel、跳过 S3 上已存在的同大小文件） |

**工作目录规则**（优先级从高到低）：`--out-dir` > `--file` 所在目录 > `OTA_WORKSPACE`。
`--file` 和 `--out-dir` 的相对路径按你运行命令时所在的目录解析，也支持 `~/xxx`。

**流程**

1. **设备校验**：设备类型代码以映射表的写法为准（仅大小写不同会自动规范化并提示）。
   代码不在映射表时：
   - 在终端里运行 → 引导你录入四项（设备类型代码、feature code、设备类型名称、MAC 组），
     可选 `y` 保存进映射表以后自动识别 / `n` 仅本次使用 / 其它取消（输入 `q` 随时退出）；
   - 非交互或加了 `--yes` → 报不一致，`--yes` 才继续。
2. **生成**：生成「升级规则 + 升级提示语」两个表的标准 Excel。同一内容已生成过则复用（`--force` 可重做）。
3. **严格校验**：表名、18 列表头、每行升级文件地址（https、.zip 结尾、无 3 个以上连续斜杠）、
   OTA 类型 ∈ {1,2,11}、设备字段等。不通过直接退出，不上传。
4. **S3 预检**：S3 上已有同名文件且大小一致的，跳过下载和上传。
5. **下载并上传**：从 `ftp://` 地址下载，传到 `s3://fam-media-andr/ota/<设备码>/<YYYYMMDD>/`，
   带进度百分比；上传后用 `head_object` 回查确认。Excel 里的地址已换成 `https://ota-dl.intsmarthub.com/...`。

**限制**：升级包只能从 `ftp://` 地址下载。申请模板里如果是 `http://` 地址，Excel 能生成，
但没有可下载的 zip，真实上传会报错。请改成 `ftp://用户@主机/路径` 的形式。

> ⚠ 真实的 FTP 下载和 S3 上传是不可逆的外部操作（升级包动辄上 GB）。演练请先加 `--no-upload`。

### 4.2 `publish`：导入 + 上线 + DASS init（带检查点）

```bash
./deployota publish               # 用 prep 刚生成的 Excel
./deployota publish --no-dass     # 只做前 3 个检查点
```

同一个浏览器会话里依次执行，**任何检查点失败就立刻停下，用红色框醒目提醒原因和下一步**，后面的步骤不再执行。

| 检查点 | 判定 | 失败退出码 |
|---|---|---|
| ① 导入上传成功 | 导入向导走完且无页面报错，模板在「OTA升级模板」列表里找到 | 11 |
| ② 得到 Rule ID | 回收到的 Rule ID 数量 == Excel 数据行数；回写 Excel 后读回一致 | 12 |
| ③ 上线成功 | 全部 Rule ID 回读为已上线，没有失败或冲突中止 | 13 |
| ④ DASS init 填报完成 | 上线全部成功后，用同一份 Excel 跑 `dass init`；退出码 0 才算通过 | 14 |

结束时总会打印一份「检查点报告」，每项显示 ✓ / ✗ / 未执行。

要点：

- **DASS 的人工确认保留**：填完后终端会问「是否批复并提交？」，输入 `approve`（或 `提交` / `yes` / `y`）才提交，
  其它输入都不提交。检查点④通过只代表填写完成。
- **防重复导入**：Excel 里已经有 Rule ID 说明导入过，`publish` 会直接拒绝（再导入会触发 CMS「目标版本已经存在」）。
  这时只需上线就用 `./deployota online`。
- **失败后不要直接重跑**：每个失败提示里都写了该怎么处理。例如检查点①失败时模板可能已在 CMS 建好，
  先去核对；检查点④失败时 CMS 已完成，单独跑 `./deployota dass init` 即可。
- 同一份 Excel 可以含**多个目标版本**，模板名称会带上全部版本的末 5 位，例如 `MT9666AndrWH80E_3.9.2_3.9.4`；
  但**设备类型代码必须只有一个**。

### 4.3 `online`：上线

单独上线（或 `publish` 中途失败后补做）时使用。导入建规则只在 `publish` 里做，不再单独提供 `upload`。

```bash
./deployota online      # 从 Excel 的 Rule ID 列读出 → 逐条上线 → 冲突处理 → 汇总
./deployota online 50000-50005               # 直接给 Rule ID 上线（单个 / 区间 / 逗号，可混用）
./deployota online 50000 --device MT9655AMAndrGTV   # 部署库里查不到设备类型代码时手工指定
```

直接给 Rule ID 时：系统按 Rule ID 自动找设备类型代码（顺序见下方「设备类型代码怎么找」，`--device` 可选），按设备类型分组，
每组的上线、冲突处理、已上线跳过都和 Excel 模式完全相同；最后打印上线汇总和部署信息汇总表
（部署时间取部署库里该批次的部署日期）。Rule ID 和 Excel 不能同时给。

**命名规则**：模板名称 = 设备类型代码 + `_` + 目标版本末 5 位（如 `MT9655AMAndrGTV_Q1015`）；
备注 = `YYYYMMDD_USA_<数据行数>`（如 `20261102_USA_9`）。

**逐条上线**

1. 在「待选规则」区的「规则ID」框输入 Rule ID 搜索，等目标行出现。
2. 点该行「上线」→ 在「确认上线该规则?」点「确定」。
3. **已经上线的会自动跳过**：开始前先在「已选规则」里查；点上线后页面提示已上线的也跳过。
4. 遇到「规则替换」冲突弹窗（「当前规则和 XXXXX 冲突，是否替换？」）：
   - 先点「取消」关掉弹窗（遮罩会挡住下面的搜索）；
   - 去「已选规则」搜被冲突的 XXXXX，读它的 **mac 组**；
   - mac 组**非空** → 重新对这条点一次上线，在第二次弹出的冲突窗里点「确定」替换；
   - mac 组**为空，或读不到** → 中止这一条并记录（保守，不猜着替换）。
5. 最后回读每个 Rule ID 的状态：`online` / `冲突中止` / `失败`，并打印**部署信息汇总**表（至少有一条上线时）。

### 4.3b `offline`：按 Rule ID 下线

```bash
./deployota offline 50000                  # 单个
./deployota offline 50000-50005            # 区间
./deployota offline 50000,50003,50010      # 逗号分隔（可与区间混用）
./deployota offline 50000-50005 --yes      # 跳过开始前的确认
./deployota offline 50000 --device MT9655AMAndrGTV   # 部署库里查不到设备类型代码时手工指定
```

流程（不做单独的状态检查，逐个直接处理）：

1. 按 Rule ID 自动找设备类型代码（见下方「设备类型代码怎么找」，`--device` 可选）。哪里都找不到的说明不在线，直接跳过。
2. 开始前确认一次：列出要下线的 Rule ID，**你输入 `yes`** 才开始（`--yes` 跳过；非交互终端不加 `--yes` 会直接取消）。
3. 打开「升级事件 → 规则上下线编排」，按设备类型分组，在「已选规则」区设一次设备类型。
4. 每个 Rule ID 在「已选规则」（在线区）按规则 ID 精确搜索：
   - 搜到且有「下 线」按钮 → 点「下 线」→ 出现「确认下线该规则?」后，确认背后搜索结果仍是目标行，才在确认框里点「确定」
     （绝不点编排弹窗底部的「确 定」）→ 重新搜索，「已选规则」里没有了才算成功；
   - 搜不到 → 已下线或不存在，**直接跳过**。
   每次搜索都核对输入框里确实是目标号，防止填错框。

最后汇总：下线成功 / 跳过（已下线或找不到）/ 查不到设备类型代码 / 失败（附原因）。
然后为下线成功的规则按设备类型代码打印**下线信息汇总**：与上线汇总同一张表，但红色标题、⛔【已下线】标识，
头部显示设备类型代码、下线规则总数、Rule ID 范围、下线时间，表里列出每条的升级类型、机型信息、机型扩展码、
源版本、目标版本（目标完整版本）和下线时间。明细优先取部署库/来源 Excel；取不到时用下线前 CMS 页面上的那一行
（CMS 页面只显示「X_to_Y」式的目标版本，这种情况下目标版本列不是完整版本）。
退出码：`0` 没有失败；`1` 有失败或查不到设备类型代码的；`2` 参数问题或你取消。

**设备类型代码怎么找**（online / offline 通用，`--device` 只在特殊情况下需要）：

1. 部署库（优先有效批次）；
2. prep 最近生成的 Excel——`publish` 的导入步骤已把 Rule ID 回写进去，刚导入、还没入库的新规则靠这一步；
3. CMS「在线规则」页按批量规则ID 查「设备类型」列（只能查到已在线的规则）。

只有「还没上线、又不在部署库和 prep Excel 里」的规则查不到（CMS「待选规则」区不先选设备类型就搜不出来），
这时 online 会列出并提示用 `--device` 指定。

### 4.4 `update`：按 Rule ID 更新 DASS 阶段（放量用）

```bash
./deployota update ota200 65370                  # 单个 Rule ID
./deployota update ota200 65370-65379            # 区间
./deployota update ota1k  65370,65375,65380      # 逗号分隔（区间可混用）
./deployota update ota50  65370-65372 65380      # 也可给多个参数
```

`OPTION` 取值：`ota50` / `ota200` / `ota1k` / `ota5k` / `ota20k` / `ota-all`（`init` 要插入新行，仍用 `./deployota dass init`）。

**它做什么**

1. 把参数解析成 Rule ID 列表。
2. 到部署库里找每个 Rule ID 所属的**有效**批次，取出该规则的方案号（设备类型代码）、完整目标版本、设备扩展信息等：
   优先用批次里存的「规则明细」；老批次没有明细，就去读它的来源 Excel。
3. 先打印一张解析结果表（Rule ID、方案号、目标版本、设备扩展信息、来源），再进入 DASS 更新：
   ota200 起先过成功率闸门；Stage1 用「方案号 + 目标版本 + 上一阶段」粗筛，Stage2 再加设备扩展信息精选；
   写入「当前阶段」和「当前阶段时间」（ota50 还会写「是否展示=1」）；
   最后仍由你在浏览器核对、在终端输入 `approve` 才提交。逻辑与下面 `dass` 的更新模式完全相同，只是输入由 Excel 改成了 Rule ID。
4. 更新后不会重复写部署库（部署记录在首次部署时已入库），也不显示汇总表。

**CMS 灰度放开**（默认开启；顺序是 成功率闸门 → CMS → DASS）：在 CMS「整机升级管理 → 在线规则」里，
按部署库取到的设备类型代码分组，「批量规则ID」填这些 Rule ID 搜索（核对结果恰好是这些号），看「累计检测数」：

判断以「阶段名称」为主，「累计检测数」交叉核对。累计检测数是**累计值**（场测 5 台，加上每档新放开的台数）：

| 档位 | 放开前应为（阶段名称 / 累计检测数） | 配置窗口选择 | 放开后（阶段名称 / 累计检测数） |
|---|---|---|---|
| ota50 | 测试阶段 / 5 | 50阶段(50) | 50阶段 / 55 |
| ota200 | 50阶段 / 55 | 200阶段(200) | 200阶段 / 255 |
| ota1k | 200阶段 / 255 | 1000阶段(1000) | 1000阶段 / 1255 |
| ota5k | 1000阶段 / 1255 | 5000阶段(5000) | 5000阶段 / 6255 |
| ota20k | 5000阶段 / 6255 | 20000阶段(20000) | 20000阶段 / 26255 |
| ota-all | 20000阶段 / 26255 | 全量(-1) | 全量 / -1 |

阶段名称对、检测数对不上时照常处理并给出提示；没有阶段名称时才只看检测数。
已经是目标阶段的跳过；其它情况（跳档、更高）不动并列出。然后批量操作类型选「批量灰度」（页面默认是**批量暂停**，
脚本会回读确认）→ 配置最大升级检测数选目标阶段、核对累计检测数自动变成目标值 → 确定 → **你输入 `yes`** →
再次确认是批量灰度后点「执行批量操作」→ 重新搜索，阶段名称变成目标阶段才算成功。
只有 CMS 放开成功（或本来已是目标阶段）的 Rule ID 才会接着更新 DASS；CMS 一个都没放开时 DASS 不动（退出码 4）。
`--yes` 跳过这次确认，`--no-cms` 只更新 DASS（旧行为）。每次最多 20 个 Rule ID 一组。

**DASS 分组提交**：当这批规则大多不在上一阶段、要按设备扩展信息逐组精选时，DASS 有未提交的编辑就不能再查询下一组。
交互模式下，每组写完会停下来让你在浏览器核对，**输入 `approve` 提交本组并自动继续下一组**；输入其它内容就停止，
剩下的记为未处理（之后用同一条命令再跑，已到目标阶段的会自动跳过）。已在目标阶段之后的规则不会被改回（不降级）。

**测试模式 `--dry-run`**：只在 CMS 把流程走一遍，什么都不改——找到规则、核对累计检测数、切到批量灰度、
打开「配置最大升级检测数」选好目标阶段并核对检测数自动填对，然后点「取 消」；不点确定、不点「执行批量操作」，
也不做成功率闸门、不更新 DASS。正式执行前建议先跑一次：

```bash
./deployota update ota50 65370-65372 65380 --dry-run
```

**成功率闸门与人工批复**（ota200 起）：闸门会逐个显示**每个 Rule ID 的成功率**。
按要求的上一阶段没查到的，会同时显示报表里该规则实际的阶段和成功率，标注「仅供参考，不算通过」。
全部 ≥ 50% 才自动放行；**不达标时，在交互终端里会问你是否人工批复**：

```
输入 override 并回车 = 人工批复继续；其它任意输入（或直接回车）= 中止：
```

- 必须原样输入 `override` 才继续（大小写、`yes` 都不行），直接回车就是中止；批复后日志会留一行「已人工批复」。
- 批复只是跳过闸门，之后**仍要你在浏览器核对、在终端输入 `approve` 才会提交**。
- 加 `--strict` 则不询问，不达标一律中止（脚本/定时任务建议加）。非交互终端本来就不会询问。

**解析不出来的 Rule ID**会逐个列出原因，默认整体中止、什么都不改；加 `--skip-missing` 则只处理能解析的：

| 原因 | 说明 |
|---|---|
| 部署库里没有 | 需先经 `publish` / `dass init` 入库 |
| 只在「作废」批次里 | 作废批次不能更新 |
| 批次没有明细且来源 Excel 不可用 | 例如 CMS 回补来的批次（来源是 `CMS:` 开头）：CMS 只给「X_to_Y」形式的版本，没有 DASS 要的完整目标版本，无法自动查询 |

> 目前部署库里约四成 Rule ID 属于上面第三种（CMS 回补的老批次）。新部署的批次会自动存规则明细，之后不依赖 Excel 是否还在。

### 4.4b `dass`：DASS 填报（init）与按 Excel 更新

日常放量请用上面的 `update`；`dass` 主要用于 `init`（插入新行，需要 Excel），`dass otaXX <Excel>` 的旧用法仍保留。

```bash
./deployota dass                       # init（默认）：插入 N 行并逐格填入
./deployota dass ota50                 # 更新已存在行的「当前阶段」「当前阶段时间」
./deployota dass ota200 '/路径/x.xlsx'
```

`OPTION` 取值：`init`（默认）/ `ota50` / `ota200` / `ota1k` / `ota5k` / `ota20k` / `ota-all`。
OPTION 和 Excel 路径都可省略，顺序不限。

**流程**：（ota200 起）先过成功率闸门 → 填入/更新并逐条打印写入明细 → 截图 `artifacts/` → 暂停，
你在浏览器里核对后，在终端输入 `approve` 才执行「验证 → 提交」，其它输入取消。

更新完成后会自动：

- 打印一张**部署汇总表**（含升级类型），
- 把本次部署写入部署库（见 4.5），幂等，同一批不会重复写。

**升级成功率闸门**

- 触发：`ota200 / ota1k / ota5k / ota20k / ota-all`；`ota50` 和 `init` 不校验。
- 规则：本批**所有** Rule ID 的升级成功率都 **≥ 50%** 才放行；任一低于 50%，或查不到，就中止
  （不更新、不进入提交环节，退出码 2）。
- 数据来源：DASS「安卓升级展示报表」的「升级成功率」列，逐条打印每个 Rule ID 的成功率、是否通过、来源阶段。
- 查询方式：先用 **方案号 + 目标版本 + 当前阶段(上一阶段)** 查（Stage1），没命中的再加 **设备扩展信息** 精准查（Stage2）。

| 升级档位 | 查询用的当前阶段（上一阶段） |
|---|---|
| `ota200` | `OTA50` |
| `ota1k` | `OTA200` |
| `ota5k` | `OTA1000` |
| `ota20k` | `OTA5000` |
| `ota-all` | `OTA20000` |

**只想看成功率、不升级**：独立的只读查询。

```bash
CMS_HEADED=true ./.venv/bin/python -m automation.dass_perf_query \
    --plancode "MT9655AndrGTVLas" --target "0100UQ1015" "65204-65212"     # 指定 Rule ID（区间/逗号可混用）
CMS_HEADED=true ./.venv/bin/python -m automation.dass_perf_query \
    --plancode "MT9655AndrGTVLas" --target "0100UQ1015" --all             # 列出该版本下全部
```

存在低于 50% 的行时返回码为 1（正常，表示这批升级会被闸门拦下）。纯只读。

### 4.4c 灰度阶段记录

每个 Rule ID 有一条灰度阶段时间线，存在 `$OTA_WORKSPACE/OTADATA/rule_stages.jsonl`（只追加的事件日志），
同时导出 `rule_stages.csv`（每个 Rule ID 一行：当前阶段、场测/50/200/1K/5K/20K/完全开放 各阶段日期、CMS 状态）。

| 事件 | 记录 |
|---|---|
| `dass init` 填写完成（含 `publish` 第 4 步） | 阶段 **场测** |
| `update otaXX` 在 CMS 灰度放开成功（回读累计检测数 = 目标值） | 阶段 50 / 200 / 1K / 5K / 20K / 完全开放 |
| `update --no-cms` 或 `dass otaXX`（不经 CMS） | 以 DASS 更新成功为准记录阶段 |
| `online` / `publish` 上线成功 | CMS 状态 **在线** |
| `offline` 下线成功 | CMS 状态 **已下线**（不改 DASS，也不动阶段时间线） |

规则：只前进不回退；同一阶段只保留第一次的日期（重跑不改日期）；跳档照写并标注。

**规则冲突记录**：`online` / `publish` 上线时出现「规则替换」冲突，会在新、旧两个 Rule ID 上各记一条：

| 情况 | 新规则 | 被冲突的旧规则 |
|---|---|---|
| 旧规则 mac 组非空 → 替换 | 「与 X 冲突 → 已替换」，附旧规则 mac 组 | 「被 新ID 替换」；回读确认已不在「已选规则」时，CMS 状态记 **已下线（被替换）** |
| 旧规则 mac 组为空 / 读不到 → 中止 | 「与 X 冲突 → 冲突中止」 | 「冲突中止（未替换）」，CMS 状态不变 |

替换后回读发现旧规则仍在线，会在汇总里提示「旧规则仍在线，请核对」。
另导出 `rule_conflicts.csv`（每次冲突一行：时间、设备类型代码、新 Rule ID、被冲突 Rule ID、结果、旧规则 mac 组、旧规则状态）。
`./deployota query <Rule ID>` 会显示该 Rule ID 的冲突记录。
`./deployota query <Rule ID>` 的结果里会显示每个 Rule ID 的当前阶段、时间线和 CMS 状态。
已有的老 Rule ID 没有补历史阶段，从现在开始记录。

### 4.5 `history`：部署历史

每次 `dass` 更新后，部署记录会自动写入部署库：`$OTA_WORKSPACE/OTADATA`（`OTA_DEPLOY_DB_DIR` 可改），
JSONL 为主存，同时导出 CSV。

```bash
./deployota history                       # 最近 20 条（按部署时间）+ 全库统计
./deployota history --device MT9655       # 按设备类型代码过滤（子串）
./deployota history --owner devinw        # 按负责人过滤
./deployota history --rule 65204          # 按 Rule ID 或范围过滤
./deployota history --limit 0             # 显示全部
./deployota history --export              # 重新导出 CSV 后退出
```

**去重规则**：以「设备类型代码 + 源版本 + 目标版本 + 设备扩展码集合」为唯一标识（不含 Rule ID）。
同标识且 Rule ID 相同 → 跳过；同标识但 Rule ID 不同 → 旧记录标「作废」、新记录为「有效」；新标识 → 写为「有效」。

### 4.6 `query`：按 Rule ID 查部署信息

给一个、一段或一组 Rule ID，查它们各自属于哪个部署批次，只读本地部署库。

```bash
./deployota query 65370                    # 单个
./deployota query 65370-65379              # 区间（闭区间）
./deployota query 65370,65375,65380        # 逗号分隔
./deployota query 65370-65372,65380        # 区间与逗号混用
./deployota query 65370 65380-65383        # 也可以给多个参数

# 按月份查：列出这段时间部署的全部批次和 Rule ID，并统计
./deployota query 2026.04                  # 一个月
./deployota query 2026.04-2026.10          # 月份区间（含首尾月；也可写 2026.04~2026.10、2026/04、2026年04月）
```

输出：

- 同一批次里命中的 Rule ID 合并成一段（如 `65370~65373`），批次信息只打印一次：设备类型代码、版本变化、
  升级类型、状态（有效/作废）、部署时间、Owner、设备扩展码、批次号、来源、入库时间。
- 同一个 Rule ID 若落在一个作废批次和一个有效批次里，两个都会列出。
- **逐条明细**：来源 Excel 还在时，再按行读出每个 Rule ID 自己的机型、设备扩展信息、源版本 → 目标版本、升级类型。
  这很重要：部署库只记批次**第一行**的版本，混有多个目标版本的批次，只看批次信息会不准，以逐条明细为准。
  来源是 `CMS:` 开头（回补得来）或文件已不在时，没有逐条明细，会给出提示。
- 末尾汇总：找到几个、未找到几个。库里没有的 Rule ID 单独列出。

按月份查时：按部署时间分月列出每个批次（与 `history` 同一行格式），每月一行小计「N 个批次，共 M 条规则」；
末尾汇总批次数（有效，作废另列）、规则条数、平台数、去重后的 Rule ID 数和范围，跨多月时再给出按月统计。
规则条数按批次累计；若同一批规则在库里被记录了两次（例如本机 Excel 入库一次、CMS 回补又入库一次），
会在汇总里提示重复条数，以「去重后 Rule ID」为准。月份和 Rule ID 不能混在一次查询里。

退出码：`0` 全部找到；`1` 有未找到的；`2` 参数有误（如区间起始大于结束）。一次最多查 5000 个。

注意：`64,700` 这种只有一组三位数字的写法会被当作带千分位逗号的单个号 64700；
要查两个号请写成 `64700,64701`。

### 4.7 旧入口

`run_prep.sh`、`run_active_ota.sh`、`run_dass.sh`、`run_login.sh` 仍可用，但**不推荐**，也不再出现在 `./deployota --help` 里；日常请只用 `./deployota`。
注意：旧入口**没有**「默认用 prep 刚生成的 Excel」这个功能，不带 Excel 时会取 `OTA_V4_OUTPUTS` 目录里最新的 `.xlsx`。
`active_ota` 还有两个调试选项：`./run_active_ota.sh --probe-online`（只读 dump 编排视图选择器）和 `--selftest`。

---

## 五、配置与环境变量

写在 `config/.env`，模板见 `config/.env.example`。

| 变量 | 说明 |
|---|---|
| `CMS_BASE_URL` / `CMS_LOGIN_URL` / `CMS_OTA_URL` | CMS 地址 |
| `CMS_USERNAME` / `CMS_PASSWORD` | CMS 凭据 |
| `DASS_URL` / `DASS_LOGIN_URL` | DASS 地址 |
| `DASS_USERNAME` / `DASS_PASSWORD` | DASS 凭据 |
| `CMS_HEADED` | `true`（默认）开可见浏览器；`false` 无头。DASS 人工核对需要有头 |
| `CMS_ARTIFACTS_DIR` | 截图和中间产物目录（默认 `artifacts/`） |
| `OTA_WORKSPACE` | 统一工作目录；prep 默认输出目录、部署库（`OTADATA`）都在这下面 |
| `OTA_DEPLOY_DB_DIR` | 单独指定部署库目录（可选） |
| `OTA_OWNER` | 写入部署记录的负责人（可选） |
| `OTA_FTP_USER` / `OTA_FTP_PASS` | prep 下载升级包用的 FTP 账号（用户名默认 `ftpuser`） |
| `OTA_S3_UPLOAD_BUCKET` | S3 桶（默认 `fam-media-andr`） |
| `OTA_AWS_PREFIX` | 下载链接前缀（默认 `https://ota-dl.intsmarthub.com`） |
| `NO_COLOR` | 设为任意值关闭彩色输出 |

运行时会生成的本地文件（均不入库）：`.last_prep.json`（prep 最近生成的 Excel）、
输出目录里的 `.prep_index.json`（prep 去重索引）、`artifacts/*.png`（每步截图）。

---

## 六、映射表（设备信息）

`prep` 用映射表补全设备信息并校验：`automation/otagen/data/device_history_seed.json`，
每条四项：`device_code`、`device_name`、`feature_code`、`mac_group`（默认 `测试` + 设备类型代码）。

- **新增设备**：直接运行 `prep`，遇到不在表里的设备会引导录入并可保存，不用手工改文件。
- 也可手工编辑该 JSON，格式同现有条目。
- 该文件是本机维护的数据；是否纳入 git 提交由你决定（当前约定不提交）。
- 设备类型代码匹配不区分大小写，以映射表的写法为准。

---

## 七、退出码

| 命令 | 退出码 |
|---|---|
| `prep` | `0` 成功；`1` 生成/校验/上传失败；`2` 参数问题或你取消 |
| `publish` | `0` 全部通过；`2` 前置检查不通过（未动 CMS）；`11` 导入失败；`12` Rule ID 问题；`13` 上线失败；`14` DASS init 失败 |
| `online` | `0` 全部成功；`1` 运行出错；`2` 无 Rule ID、存在冲突中止或未上线 |
| `dass` | `0` 成功；`1` 出错；`2` 成功率闸门未通过；`3` init 有行没填成功 |
| `dass_perf_query` | `1` 存在成功率低于 50% 的 Rule ID |
| `query` | `0` 全部找到；`1` 有未找到的；`2` 参数有误 |
| `offline` | `0` 没有失败；`1` 有失败或查不到设备类型代码的；`2` 参数问题或取消 |

---

## 八、常见问题与排错

| 现象 | 原因与处理 |
|---|---|
| `prep` 报「设备类型代码不在映射表中」 | 在终端运行会自动引导录入；非交互场景加 `--yes`，或先录入映射表 |
| `prep` 提示 zip 地址「不是 ftp://」 | 申请模板里是 `http://` 地址。改成 `ftp://用户@主机/路径`，本工具只支持 FTP 下载 |
| `prep` FTP 连接被拒绝 / 登录失败 | 检查 `OTA_FTP_PASS`、FTP 主机是否可达；地址里的主机必须真的有 FTP 服务 |
| `--file` 报找不到文件 | 路径写错（注意文件名里的空格）；路径含空格或中文要加引号 |
| `--file` 报「临时锁文件」 | `~$xxx.xlsx` 是 Excel 打开时产生的，请指定原文件 |
| `--file` 报「不支持旧版 Excel」 | `.xls` 请另存为 `.xlsx` |
| `publish` 报「目标版本已经存在」 | CMS 里已有同设备同目标版本的模板。先核对旧模板；已导入过就用 `online`，不要重跑 |
| `publish` 拒绝「Excel 已含 Rule ID」 | 说明已导入。只需上线用 `./deployota online` |
| 报「设备类型代码应为单一值」 | 一份 Excel 只能含一个设备类型代码，请拆分文件（目标版本可以有多个） |
| 检查点② 数量不符 | 到 CMS 核对模板的规则数量；不会自动上线。核对后把 Rule ID 填进 Excel 再用 `online` |
| 上线「gate 未满足」/「搜出 0 行」 | 看日志里的 ✗ 提示和 `artifacts/active_gate_fail_<RuleID>.png`；已上线的会自动跳过，处理后重跑 `online` 即可 |
| 冲突中止 | 被冲突规则的 mac 组为空或读不到，需人工判断是否替换 |
| 成功率闸门未通过（退出码 2） | 本批有 Rule ID 低于 50% 或查不到；看逐条明细里每个 ID 的成功率。确认风险可接受时，在提示处输入 `override` 人工批复继续（`--strict` 关闭此提示） |
| 闸门显示「按要求的阶段未找到；报表里该规则阶段=OTA200」 | 展示报表是按日期的快照，滞后于填报页，阶段还没同步。等报表日期推进后重试，或人工批复 |
| DASS 检查点④失败（退出码 14） | CMS 已完成，不要重跑 publish；看 dass 日志后单独 `./deployota dass init` |
| 页面选择器失效 | 站点改版所致。先 `./run_active_ota.sh --probe-online` 看现状，再让维护者校正 |

出问题时先看终端里的红色提示和 `artifacts/` 里最新的截图。

---

## 九、安全边界

- 凭据只从 `config/.env` 读取，不写进代码，不入库。
- 「提交」是高影响操作：`dass` 在提交前暂停等你输入 `approve`，不会无人值守发布。
- 「上线 / 替换」会改线上规则且不可撤销：写按钮都经过二次确认检查，条件不满足只打印不点；
  mac 组等判断只要不确定就保守中止，不猜。
- `prep` 的 FTP 下载和 S3 上传不可逆，演练先加 `--no-upload`。
- 升级成功率不达标时自动中止，避免把低成功率批次继续放量。
- 每个关键步骤自动截图到 `artifacts/`，便于核对。
- 请确认自动化内部系统的做法符合公司规范。

---

## 十、测试与目录结构

不连浏览器的离线自测：

```bash
./.venv/bin/python -m automation.prep --selftest
./.venv/bin/python -m automation.active_ota --selftest
./.venv/bin/python -m automation.active_offline --selftest
./.venv/bin/python -m automation.dass_fill --selftest
./.venv/bin/python -m automation.dass_perf --selftest
./.venv/bin/python -m automation.dass_perf_query --selftest
./.venv/bin/python -m automation.depdb --selftest
./.venv/bin/python -m automation.dass_update --selftest
./.venv/bin/python -m automation.rule_info --selftest
./.venv/bin/python -m automation.depdb_backfill --selftest
./.venv/bin/python -m automation.summary --selftest
```

真实页面行为（翻页、逐条日志、上线、冲突处理）需要连真实系统确认，建议先用单条测试 Excel 试跑。

```
ota_cms_automation/
  deployota               # 统一入口（薄分发）
  automation/
    deployota.py          # 子命令分发 + 默认使用 prep 刚生成的 Excel
    prep.py               # 生成 Excel → 校验 → FTP 下载 → 上传 S3
    last_prep.py          # 记录 prep 最近生成的 Excel（.last_prep.json）
    otagen/               # 生成引擎、传输、映射表（data/device_history_seed.json）
    active_ota.py         # CMS：导入建规则(upload) / 上线(online) / 合并(publish)
    active_offline.py     # CMS：按 Rule ID 下线(offline)
    dass_fill.py          # DASS 填报/更新（init、otaXX），含成功率闸门接入
    dass_update.py        # deployota update：按 Rule ID 解析记录后调用 dass_fill
    dass_perf.py          # 升级成功率闸门：查询 + 校验 + 逐条日志
    dass_perf_query.py    # 独立只读的成功率查询
    dass_data.py          # 读取更新后的 Excel → 填报字段
    depdb.py              # 部署库（JSONL + CSV）与 history
    rule_info.py          # 按 Rule ID 查部署信息
    depdb_backfill.py     # CMS 历史只读回补（维护用，已不在 deployota 菜单里：python -m automation.depdb_backfill）
    summary.py            # dass 更新后的部署汇总表
    _log.py               # 彩色日志
    login.py / dass_login.py / inspect_*.py   # 登录与页面探测
  config/                 # .env.example（真实 .env 不入库）
  artifacts/              # 运行截图
  .agents/tasks/          # 只读探测结论与评审记录
  run_prep.sh / run_active_ota.sh / run_dass.sh / run_login.sh   # 旧入口，仍可用
  requirements.txt
```
