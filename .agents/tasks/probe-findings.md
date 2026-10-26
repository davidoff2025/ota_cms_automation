# DASS 查询 UI 探测结果（实现者可直接复制）

- 生成时间：2026-10-26 15:17:11
- 站点可达：是

## 过滤框选择器

| 字段 | 选择器 |
| --- | --- |
| 方案号 | `input[name='PLANCODE']`  (name='PLANCODE', 当前value='') |
| 目标版本（UI 标签显示为『标版本』） | `input[name='NEXTVERSION']`  (name='NEXTVERSION', 当前value='') |
| 设备扩展信息 | `input[name='DEVICEMSG']`  (name='DEVICEMSG', 当前value='') |
| 当前阶段 | `input[name='CURRENTSTAGE']`  (name='CURRENTSTAGE', 当前value='') |
| 大区域 | `input[name='REGION']`  (name='REGION', 当前value='') |
| 国家 | `input[name='COUNTRYCODE']`  (name='COUNTRYCODE', 当前value='') |
| 机型 | `input[name='DEVICECODE']`  (name='DEVICECODE', 当前value='') |
| 源版本 | `input[name='CURRENTVERSION']`  (name='CURRENTVERSION', 当前value='') |

## 显示行数（limit）输入框
- 选择器：`input[name='LIMIT']`
- 设为 50：True；设为 100：True
- 机制：该框内容被原样拼进数据集 SQL 的 LIMIT 子句，必须写完整片段 'limit N'（不能只写数字）。设 50 => fill('limit 50')；设 100 => fill('limit 100')。只写 '50' 会报 11300001 数据集配置错误（SQL 语法错误 near '50'）。

## 查询按钮
- 选择器：`#fr-btn-BTNSUBMIT`（取可见且文本=『查询』的叶子元素；box={'x': 1085, 'y': 162, 'width': 80, 'height': 21}）

## 结果读取方式
- existing_row_info(fr, letters) 读取查询结果（方案号非空的数据行）
- 结论：查询后网格显示的行即查询结果，`dass_fill.existing_row_info(fr, letters)` 可直接读取（返回方案号非空的数据行，带 n/方案号/设备扩展信息/机型/目标版本/ruleid/当前阶段）。

## 查询后是否需重读表头布局（read_column_layout）
- 会话锚：before='1512' after='1513' -> 变化=True
- 列字母是否变化：False
- 结论：**每次查询后必须重新调用 `read_column_layout(fr)`**（会话锚/列字母会随网格刷新改变），再用新的 letters/session 去读行与拼单元格 id。

## 单击 vs 双击进入单元格编辑态
- 需双击进入编辑态（单击=False，双击=True）。dass_fill.fill_cell 的双击做法成立。
- 说明：测试时仅进入编辑态后按 Escape 退出，**未写入任何值**。

## 样例查询结果
- 结果行数（limit=50）：50；limit=100 行数：100
- （下表为前 8 行样例）

| 行号n | 方案号 | ruleid | 目标版本 | 当前阶段 |
| --- | --- | --- | --- | --- |
| 2 | MT9655AMAndrGTV | 65,232 | 0100BQ1022 |  |
| 3 | MT9666Android | 65,230 | V1.4.3.0 |  |
| 4 | MT9666Android | 65,229 | V1.4.3.0 |  |
| 5 | MT9666Android | 65,231 | V1.4.3.0 |  |
| 6 | A311D2AndrT4K | 65,224 | 0001MQ0909 |  |
| 7 | A311D2AndrT4K | 65,223 | 0001MQ0909 |  |
| 8 | A311D2AndrT4K | 65,225 | 0001MQ0909 |  |
| 9 | A311D2AndrT4K | 65,224 | 0001MQ0909 |  |

## 原始探测转储（供排错）

### inputs
- label=None id='cpTextArea' name='' placeholder='' value='' cls='' widget='' @(-1000,-1000 182x36)
- label=None id='' name='sortable0' placeholder='' value='推荐1' cls='sortable-texteditor sortable-common-item' widget='' @(30,28 90x29)
- label=None id='' name='sortable1' placeholder='' value='推荐2' cls='sortable-texteditor sortable-common-item' widget='' @(120,28 90x29)
- label=None id='' name='REGION' placeholder='' value='' cls='fr-texteditor' widget='' @(87,71 75x21)
- label=None id='' name='LIMIT' placeholder='' value='limit 10' cls='fr-texteditor' widget='' @(579,72 80x21)
- label=None id='' name='COUNTRYCODE' placeholder='' value='' cls='fr-texteditor' widget='' @(730,72 80x21)
- label=None id='' name='PLANCODE' placeholder='' value='' cls='fr-texteditor' widget='' @(331,73 100x21)
- label=None id='' name='DEVICEMSG' placeholder='' value='' cls='fr-texteditor' widget='' @(124,110 120x21)
- label=None id='' name='DEVICECODE' placeholder='' value='' cls='fr-texteditor' widget='' @(331,110 140x21)
- label=None id='' name='CURRENTSTAGE' placeholder='' value='' cls='fr-texteditor' widget='' @(579,110 92x21)
- label=None id='' name='CURRENTVERSION' placeholder='' value='' cls='fr-texteditor' widget='' @(730,110 100x21)
- label=None id='' name='NEXTVERSION' placeholder='' value='' cls='fr-texteditor' widget='' @(900,110 100x21)

### buttons
- 'save template' <div> id='' cls='fr-btn ui-state-enabled save-params-template-height' @(912,32)
- 'save template' <div> id='fr-btn-saveBtn' cls='fr-btn-up  fr-btn-noicon save-params-template-height' @(912,32)
- 'save template' <div> id='' cls='fr-btn-small fr-btn-icon-small-left' @(912,35)
- 'save template' <button> id='' cls='fr-btn-text' @(936,35)
- '查询' <div> id='' cls='fr-btn ui-state-enabled' @(855,72)
- '查询' <div> id='fr-btn-BTNSUBMIT' cls='fr-btn-up  fr-btn-noicon' @(855,72)
- '查询' <div> id='' cls='fr-btn-small fr-btn-icon-small-left' @(855,75)
- '查询' <button> id='' cls='fr-btn-text' @(880,75)
- '插入行' <div> id='' cls='fr-btn ui-state-enabled' @(2575,212)
- '插入行' <div> id='fr-btn-' cls='fr-btn-up  fr-btn-text-icon' @(2575,212)
- '插入行' <div> id='' cls='fr-btn-small fr-btn-icon-small-left' @(2575,219)
- '插入行' <button> id='' cls='fr-btn-text x-emb-add' @(2583,219)
- '删除行' <div> id='' cls='fr-btn ui-state-enabled' @(2647,212)
- '删除行' <div> id='fr-btn-' cls='fr-btn-up  fr-btn-text-icon' @(2647,212)
- '删除行' <div> id='' cls='fr-btn-small fr-btn-icon-small-left' @(2647,219)
- '删除行' <button> id='' cls='fr-btn-text x-emb-delete' @(2654,219)
- '插入行' <div> id='' cls='fr-btn ui-state-enabled' @(2575,244)
- '插入行' <div> id='fr-btn-' cls='fr-btn-up  fr-btn-text-icon' @(2575,244)
- '插入行' <div> id='' cls='fr-btn-small fr-btn-icon-small-left' @(2575,246)
- '插入行' <button> id='' cls='fr-btn-text x-emb-add' @(2583,246)
- '删除行' <div> id='' cls='fr-btn ui-state-enabled' @(2647,244)
- '删除行' <div> id='fr-btn-' cls='fr-btn-up  fr-btn-text-icon' @(2647,244)
- '删除行' <div> id='' cls='fr-btn-small fr-btn-icon-small-left' @(2647,246)
- '删除行' <button> id='' cls='fr-btn-text x-emb-delete' @(2654,246)
- '插入行' <div> id='' cls='fr-btn ui-state-enabled' @(2575,265)
- '插入行' <div> id='fr-btn-' cls='fr-btn-up  fr-btn-text-icon' @(2575,265)
- '插入行' <div> id='' cls='fr-btn-small fr-btn-icon-small-left' @(2575,267)
- '插入行' <button> id='' cls='fr-btn-text x-emb-add' @(2583,267)
- '删除行' <div> id='' cls='fr-btn ui-state-enabled' @(2647,265)
- '删除行' <div> id='fr-btn-' cls='fr-btn-up  fr-btn-text-icon' @(2647,265)
- '删除行' <div> id='' cls='fr-btn-small fr-btn-icon-small-left' @(2647,267)
- '删除行' <button> id='' cls='fr-btn-text x-emb-delete' @(2654,267)
- '插入行' <div> id='' cls='fr-btn ui-state-enabled' @(2575,286)
- '插入行' <div> id='fr-btn-' cls='fr-btn-up  fr-btn-text-icon' @(2575,286)
- '插入行' <div> id='' cls='fr-btn-small fr-btn-icon-small-left' @(2575,288)
- '插入行' <button> id='' cls='fr-btn-text x-emb-add' @(2583,288)
- '删除行' <div> id='' cls='fr-btn ui-state-enabled' @(2647,286)
- '删除行' <div> id='fr-btn-' cls='fr-btn-up  fr-btn-text-icon' @(2647,286)
- '删除行' <div> id='' cls='fr-btn-small fr-btn-icon-small-left' @(2647,288)
- '删除行' <button> id='' cls='fr-btn-text x-emb-delete' @(2654,288)
- '插入行' <div> id='' cls='fr-btn ui-state-enabled' @(2575,307)
- '插入行' <div> id='fr-btn-' cls='fr-btn-up  fr-btn-text-icon' @(2575,307)
- '插入行' <div> id='' cls='fr-btn-small fr-btn-icon-small-left' @(2575,314)
- '插入行' <button> id='' cls='fr-btn-text x-emb-add' @(2583,314)
- '删除行' <div> id='' cls='fr-btn ui-state-enabled' @(2647,307)
- '删除行' <div> id='fr-btn-' cls='fr-btn-up  fr-btn-text-icon' @(2647,307)
- '删除行' <div> id='' cls='fr-btn-small fr-btn-icon-small-left' @(2647,314)
- '删除行' <button> id='' cls='fr-btn-text x-emb-delete' @(2654,314)
- '插入行' <div> id='' cls='fr-btn ui-state-enabled' @(2575,339)
- '插入行' <div> id='fr-btn-' cls='fr-btn-up  fr-btn-text-icon' @(2575,339)
- '插入行' <div> id='' cls='fr-btn-small fr-btn-icon-small-left' @(2575,346)
- '插入行' <button> id='' cls='fr-btn-text x-emb-add' @(2583,346)
- '删除行' <div> id='' cls='fr-btn ui-state-enabled' @(2647,339)
- '删除行' <div> id='fr-btn-' cls='fr-btn-up  fr-btn-text-icon' @(2647,339)
- '删除行' <div> id='' cls='fr-btn-small fr-btn-icon-small-left' @(2647,346)
- '删除行' <button> id='' cls='fr-btn-text x-emb-delete' @(2654,346)
- '插入行' <div> id='' cls='fr-btn ui-state-enabled' @(2575,371)
- '插入行' <div> id='fr-btn-' cls='fr-btn-up  fr-btn-text-icon' @(2575,371)
- '插入行' <div> id='' cls='fr-btn-small fr-btn-icon-small-left' @(2575,378)
- '插入行' <button> id='' cls='fr-btn-text x-emb-add' @(2583,378)
- '删除行' <div> id='' cls='fr-btn ui-state-enabled' @(2647,371)
- '删除行' <div> id='fr-btn-' cls='fr-btn-up  fr-btn-text-icon' @(2647,371)
- '删除行' <div> id='' cls='fr-btn-small fr-btn-icon-small-left' @(2647,378)
- '删除行' <button> id='' cls='fr-btn-text x-emb-delete' @(2654,378)
- '插入行' <div> id='' cls='fr-btn ui-state-enabled' @(2575,403)
- '插入行' <div> id='fr-btn-' cls='fr-btn-up  fr-btn-text-icon' @(2575,403)
- '插入行' <div> id='' cls='fr-btn-small fr-btn-icon-small-left' @(2575,410)
- '插入行' <button> id='' cls='fr-btn-text x-emb-add' @(2583,410)
- '删除行' <div> id='' cls='fr-btn ui-state-enabled' @(2647,403)
- '删除行' <div> id='fr-btn-' cls='fr-btn-up  fr-btn-text-icon' @(2647,403)
- '删除行' <div> id='' cls='fr-btn-small fr-btn-icon-small-left' @(2647,410)
- '删除行' <button> id='' cls='fr-btn-text x-emb-delete' @(2654,410)
- '插入行' <div> id='' cls='fr-btn ui-state-enabled' @(2575,435)
- '插入行' <div> id='fr-btn-' cls='fr-btn-up  fr-btn-text-icon' @(2575,435)
- '插入行' <div> id='' cls='fr-btn-small fr-btn-icon-small-left' @(2575,442)
- '插入行' <button> id='' cls='fr-btn-text x-emb-add' @(2583,442)
- '删除行' <div> id='' cls='fr-btn ui-state-enabled' @(2647,435)
- '删除行' <div> id='fr-btn-' cls='fr-btn-up  fr-btn-text-icon' @(2647,435)
- '删除行' <div> id='' cls='fr-btn-small fr-btn-icon-small-left' @(2647,442)
- '删除行' <button> id='' cls='fr-btn-text x-emb-delete' @(2654,442)
- '插入行' <div> id='' cls='fr-btn ui-state-enabled' @(2575,467)
- '插入行' <div> id='fr-btn-' cls='fr-btn-up  fr-btn-text-icon' @(2575,467)
- '插入行' <div> id='' cls='fr-btn-small fr-btn-icon-small-left' @(2575,474)
- '插入行' <button> id='' cls='fr-btn-text x-emb-add' @(2583,474)
- '删除行' <div> id='' cls='fr-btn ui-state-enabled' @(2647,467)
- '删除行' <div> id='fr-btn-' cls='fr-btn-up  fr-btn-text-icon' @(2647,467)
- '删除行' <div> id='' cls='fr-btn-small fr-btn-icon-small-left' @(2647,474)
- '删除行' <button> id='' cls='fr-btn-text x-emb-delete' @(2654,474)
- 'Submit' <div> id='' cls='fr-btn ui-state-enabled' @(1,6)
- 'Submit' <div> id='fr-btn-Submit' cls='fr-btn-up  fr-btn-text-icon' @(1,6)
- 'Submit' <div> id='' cls='fr-btn-small fr-btn-icon-small-left' @(1,6)
- 'Submit' <button> id='' cls='fr-btn-text x-emb-submit' @(1,6)
- 'Data validation' <div> id='' cls='fr-btn ui-state-enabled' @(71,6)
- 'Data validation' <div> id='fr-btn-Verify' cls='fr-btn-up  fr-btn-text-icon' @(71,6)
- 'Data validation' <div> id='' cls='fr-btn-small fr-btn-icon-small-left' @(71,6)
- 'Data validation' <button> id='' cls='fr-btn-text x-emb-verify' @(71,6)
- 'Print[Client]' <div> id='' cls='fr-btn ui-state-enabled' @(184,6)
- 'Print[Client]' <div> id='fr-btn-FlashPrint' cls='fr-btn-up  fr-btn-text-icon' @(184,6)
- 'Print[Client]' <div> id='' cls='fr-btn-small fr-btn-icon-small-left' @(184,6)
- 'Print[Client]' <button> id='' cls='fr-btn-text x-emb-flashprint' @(184,6)
- 'Print' <div> id='' cls='fr-btn ui-state-enabled' @(279,6)
- 'Print' <div> id='fr-btn-' cls='fr-btn-up  fr-btn-text-icon' @(279,6)
- 'Print' <div> id='' cls='fr-btn-small fr-btn-icon-small-left' @(279,6)
- 'Print' <em> id='' cls='fr-btn-arrow' @(279,6)
- 'Print' <button> id='' cls='fr-btn-text x-emb-print' @(279,6)
- 'Export' <div> id='' cls='fr-btn ui-state-enabled' @(346,6)
- 'Export' <div> id='fr-btn-' cls='fr-btn-up  fr-btn-text-icon' @(346,6)
- 'Export' <div> id='' cls='fr-btn-small fr-btn-icon-small-left' @(346,6)
- 'Export' <em> id='' cls='fr-btn-arrow' @(346,6)
- 'Export' <button> id='' cls='fr-btn-text x-emb-export' @(346,6)
- 'Email' <div> id='' cls='fr-btn ui-state-enabled' @(424,6)
- 'Email' <div> id='fr-btn-Email' cls='fr-btn-up  fr-btn-text-icon' @(424,6)
- 'Email' <div> id='' cls='fr-btn-small fr-btn-icon-small-left' @(424,6)
- 'Email' <button> id='' cls='fr-btn-text x-emb-email' @(424,6)
- 'Add Record' <div> id='' cls='fr-btn ui-state-enabled' @(487,6)
- 'Add Record' <div> id='fr-btn-AppendColumnRow' cls='fr-btn-up  fr-btn-text-icon' @(487,6)
- 'Add Record' <div> id='' cls='fr-btn-small fr-btn-icon-small-left' @(487,6)
- 'Add Record' <button> id='' cls='fr-btn-text x-emb-appendrow' @(487,6)
- 'Delete Column Row' <div> id='' cls='fr-btn ui-state-enabled' @(716,6)
- 'Delete Column Row' <div> id='fr-btn-DeleteColumnRow' cls='fr-btn-up  fr-btn-text-icon' @(716,6)
- 'Delete Column Row' <div> id='' cls='fr-btn-small fr-btn-icon-small-left' @(716,6)
- 'Delete Column Row' <button> id='' cls='fr-btn-text x-emb-deleterow' @(716,6)
