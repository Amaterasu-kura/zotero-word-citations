# zotero-word-citations

> English version: [README.en.md](README.en.md)

> **一句话：把参考文献从「手打的死编号」变成「Zotero 的活引用」。**

手打的 `[1,2]`，删掉一段就得把后面几十个编号重排一遍；投稿前被要求换一种引用格式，更是逐条手改的噩梦。
本 skill 只要读一遍你的 Word 稿子，就能把正文每个 `[1,2]` 换成 Zotero 的活引用域、把文末那份死参考文献表换成 Zotero 自动生成的活列表。
此后在 Word 里点一次 **Zotero → Refresh**，编号全自动重排；换引用样式，一次 **Document Preferences** 点击搞定。
原稿绝不改动，结果另存为新文档。

> Turn a Word manuscript whose citations are typed plain text into one with **live, updatable Zotero citation fields**, and regenerate its reference list from Zotero.

把一个「正文引用是手打纯文本、文末参考文献表是死编号」的 Word 稿子，转换成引用由 **Zotero 管理**的文档：
正文里每个 `[1,2]` 变成 Zotero 的 `ZOTERO_ITEM` 域，文末参考文献表变成 `ZOTERO_BIBL` 域。
转换完成后在 Word 里点一次 **Zotero → Refresh**，编号自动重排；要换引用样式，**Document Preferences** 一键切换 —— 这两件事是纯文本引用表永远做不到的。

本仓库是一个 **agent skill**（`SKILL.md` 提供触发条件与操作规范），配套 3 个 Python 脚本完成「解析 → 注入 → 校验」。

---

## 效果

- 每个正文引用 = 一个 `ADDIN ZOTERO_ITEM CSL_CITATION` 域，携带所引 Zotero 条目的 URI
- 文末参考文献表 = 一个 `ADDIN ZOTERO_BIBL` 域，由 Zotero 在 Refresh 时生成
- 文档首选项（`ZOTERO_PREF_*` 自定义属性）记录所选引用样式
- **绝不覆盖原文件**：只输出一份新文档

## 前置条件

| 依赖 | 要求 / 检查方式 |
|---|---|
| Zotero | 正在运行，并已开启本地 API：**Settings → Advanced → "Allow other applications on this computer to communicate with Zotero"**。验证：`curl -s -o /dev/null -w '%{http_code}' "http://127.0.0.1:23119/api/users/0/items?limit=1"` 应返回 `200` |
| Zotero Word 插件 | 已安装 `Zotero.dotm`（Windows 路径 `%APPDATA%\Microsoft\Word\STARTUP\`）。没有它，注入的域永远无法刷新 |
| Microsoft Word | 必须在 **Word** 里编辑与刷新；**WPS 不加载该插件**（稿件由 WPS 写的没关系，只要用 Word 打开） |
| **zotero-mcp** | 让 AI agent 直接检索 / 写入 Zotero 文献库（例如按 DOI 用 `zotero_add_item` 入库）。安装见下节 |
| Python | 3.9+，**脚本只用标准库，无需 pip 安装任何依赖** |
| 稿件形式 | 正文编号方括号 `[1,2]` + 文末编号参考文献表。作者-年份 `(Smith, 2020)`、上标数字、或正文无标记的稿件**暂不支持** |

### 安装 zotero-mcp（推荐）

[zotero-mcp](https://github.com/54yyyu/zotero-mcp) 把 Zotero 文献库接入 AI 助手（Model Context Protocol）。本 skill 在「参考文献不在库里」「需要按 DOI 导入」时依赖它：

```bash
uv tool install zotero-mcp-server     # 或: pip install zotero-mcp-server
zotero-mcp setup                      # 自动配置 Claude Desktop 等客户端
```

随后在 Zotero 7+ 打开 **Settings → Advanced → Allow other applications…**；如需写入，Zotero 10+ 执行一次 `zotero-mcp authorize-local`（选择 Always Allow），旧版则配置 `ZOTERO_API_KEY` 与 `ZOTERO_LIBRARY_ID` 走 Web API。

> MCP 服务器与「agent skill」两种接法共享同一份配置，详见 zotero-mcp 仓库的文档。

## 安装本 skill

把整个目录放进你的 agent skill 目录即可（例如 Claude Code 为 `~/.claude/skills/zotero-word-citations/`，其它客户端见各自文档）。
`SKILL.md` 的 frontmatter 已写好 `name` 与触发词（"把文献插进 Word"、"引用无法更新"、"参考文献表是纯文本"、"换引用样式" 等），
agent 会在相关请求中自动加载。

## 快速开始

三个脚本按顺序运行，参数全部显式传入，不硬编码任何机器路径。建议在稿件旁边建一个临时目录存放中间产物。

```bash
# 0) 先确认稿件形式：正文 [1,2] + 文末编号参考文献表，并选定引用样式 ID
#    查看本机已安装样式:  ls "$ZOTERO_DATA/styles/"

# 1) 解析参考文献 → Zotero 条目
python scripts/resolve.py --docx "manuscript.docx" --out-dir "work"

# 2) 注入域（把 [n] 变成 Zotero 引用域，替换文末引用表，写入样式首选项）
python scripts/inject.py \
  --docx "manuscript.docx" \
  --refmap "work/refmap.json" \
  --itemmeta "work/itemmeta.json" \
  --style "http://www.zotero.org/styles/elsevier-harvard" \
  --out "manuscript_v2.docx"

# 3) 校验（位置对应 + 语义正确，两道独立检查；不要跳过）
python scripts/verify.py \
  --original "manuscript.docx" \
  --output "manuscript_v2.docx" \
  --refmap "work/refmap.json" \
  --itemmeta "work/itemmeta.json"
```

**4) 最后一步只能由 Word 完成**：打开 `manuscript_v2.docx` → 点 **Zotero → Refresh**，引用重排、参考文献表生成。

> 若 Refresh 后毫无变化：先**关闭其它所有 Word 窗口** —— Zotero 的 Word 集成同一时刻只处理一个文档，
> 另一个已挂 Zotero 域的窗口会让 Refresh 静默失败。也不要用脚本另起一个隐藏 Word 实例来代替手动刷新。

## 脚本说明

### `scripts/resolve.py` — 解析参考文献到 Zotero 条目

| 参数 | 说明 |
|---|---|
| `--docx` | 待处理的稿件 |
| `--out-dir` | 中间产物目录 |
| `--threshold` | 标题 token 包含度阈值（默认 `0.85`） |

两遍匹配：先用标题 token 包含度在库内匹配，未命中的再用条目里的 DOI 到 CrossRef 取权威标题后二次匹配
（稿件措辞常与 Zotero 标题略有出入，例如 "in **the** water-sediment system" vs "in water sediment system"）。

产物（写入 `--out-dir`）：

| 文件 | 内容 |
|---|---|
| `refmap.json` | `{文献编号: Zotero item key}` |
| `itemmeta.json` | `{item key: {itemID, csl}}`，供 `inject.py` 使用 |
| `citation-map.md` | 人工复核用的对照表 |
| `unresolved.md` | 未能匹配的条目（含 DOI），正常情况下为空 |

> **有未定位条目时先停下并向用户报告 DOI**。本 skill **不会**擅自往用户文献库里导入任何东西 —— 是否入库由用户决定；用户同意后再用 zotero-mcp 的 `zotero_add_item` 按 DOI 导入。

### `scripts/inject.py` — 注入 Zotero 域

| 参数 | 说明 |
|---|---|
| `--docx` | 原稿 |
| `--refmap` / `--itemmeta` | 上一步的产物 |
| `--style` | CSL 样式 URL |
| `--out` | 输出文件（**不允许与输入相同**，脚本会拒绝覆盖） |
| `--keep-red` | 保留直接红色字符格式（默认会清除，见下） |

除正文改写外还会：替换文末参考文献表为 `ZOTERO_BIBL` 域、写入文档首选项、清除稿件常带的红色标注残留。
正文改写不是简单查找替换 —— 引用 token 常被拆散在多个 `<w:t>` run 中（红色标记、拼写检查边界、修订标记都会切断它），
脚本把整段 run 拼接后定位 token，再重建 run 结构。

### `scripts/verify.py` — 校验（务必执行）

两道独立检查，因为它们的失效方式互不覆盖：

- **A. 位置对应**：按正文顺序，输出文档中第 N 个域所引条目，必须与原稿第 N 个 `[n]` token 映射结果完全一致 —— 捕捉 run 重建引入的错位。
- **B. 语义正确**：逐条比对手稿条目与 Zotero 条目元数据（先 DOI，再经 CrossRef 解析 DOI 比对标题，最后退化为标题包含度）—— 捕捉 A 看不见的「映射到了错误的文献」，并常常**顺带暴露原稿自身的错误**（DOI 无法解析、或指向完全不同的论文）。

校验不通过会以非 0 退出。属于稿件自身问题的（错误 DOI 等）会以 `[NOTE]` 报出但不算失败 —— 请如实转告用户，不要擅自"修正"你没被要求改动的参考文献表。

## 引用样式（CSL style ID）

样式以 ID 形式传入，格式为 `http://www.zotero.org/styles/<name>`：

| 场景 | 样式 ID |
|---|---|
| Elsevier 期刊（哈佛格式） | `elsevier-harvard` |
| 美国化学会 | `american-chemical-society` |
| 国标 GB/T 7714（顺序编码制） | `china-national-standard-gb-t-7714-2015-numeric` |
| Vancouver | `vancouver` |

样式决定每条引用长什么样，**必须向用户确认**而不是猜测。转换后仍可改：Word 中 **Zotero → Document Preferences**。

## 限制

- 只支持编号方括号引用（`[1,2]`）；作者-年份与上标数字形式尚未解析
- Windows 取向：Word 自动化提示与插件路径假设 Windows（docx 改写本身跨平台）
- 不附带 PDF，只使用文献元数据

## 故障排查

| 现象 | 原因 |
|---|---|
| Word 点 Refresh 直接卡死、无报错 | 文档是手工从零构造的（非 Word 写出），Word 进入兼容模式；务必基于真实文档改写 |
| Zotero 完全无视这些域 | 缺 `ZOTERO_PREF_*` 自定义属性（`docProps/custom.xml` 缺失或未登记进 `[Content_Types].xml` / `_rels/.rels`） |
| Refresh 正常返回但毫无变化 | Zotero 的 Word 集成正被另一个打开的文档占用；关掉其它所有 Word 窗口重试 |
| 刷新后域变成"已取消链接"的文本 | `uris` 写错（手工拼造的 URI 或他人的 user id）—— 必须从 Zotero 取 |
| 引用能渲染但指向错误的论文 | 文献→条目映射错误；由 `verify.py` 的语义检查捕捉 |
| 参考文献表还是旧文字 | 未插入 bibliography 域，或旧列表段落没被替换掉 |
| 正文出现 `&lt;16 µm`、`&gt;400,000` | run 改写时把已转义的 XML 实体二次转义；只命中同时含实体与引用 token 的 run，`verify.py` 会检查 |
| 文档原有自定义属性消失 | `docProps/custom.xml` 被整体替换而非合并（脚本会保留既有属性） |

更多字段格式细节（域语法、`ZOTERO_PREF` 布局、调试方法）见 [`references/field-format.md`](references/field-format.md)。

## 目录结构

```
zotero-word-citations/
├── SKILL.md                  # agent skill 定义（触发条件 + 操作规范）
├── README.md
├── LICENSE
├── scripts/
│   ├── resolve.py            # 1) 参考文献 → Zotero 条目
│   ├── inject.py             # 2) 注入 ZOTERO_ITEM / ZOTERO_BIBL 域
│   ├── verify.py             # 3) 位置 + 语义双重校验
│   └── zcommon.py            # 共用工具（Zotero API / docx 解析，仅标准库）
├── references/
│   └── field-format.md       # Zotero Word 域格式与踩坑记录
└── evals/
    └── evals.json            # 评测用例
```

## 工作原理（要点）

Zotero 的域格式是有明确定义的，但坑很集中，本 skill 主要就是把它们固化下来：

- **永远不要从零构造 docx**：最小化手工文档会让 Word 进入兼容模式，Zotero Refresh 随即**无报错地冻结 Word**（脚本与手工均已复现）。正确做法是在用户真实文件上原地改写 `word/document.xml`，其余部件原样复制。
- 域指令两侧的**空格**是必须的；`uris` 是 Zotero 刷新时真正解析的东西，且只能从 Zotero 取得。
- 数字 `itemID` 不在本地 API 中暴露，需从 `zotero.sqlite` 的 WAL 快照副本读取（Zotero 对活动库持排他锁）。
- `data-version` 必须为 **3**；每个自定义属性值受 Office 的 255 字符上限约束，故拆成 `ZOTERO_PREF_1, _2, …`。

## 致谢

- [zotero-mcp](https://github.com/54yyyu/zotero-mcp)（作者 [54yyyu](https://github.com/54yyyu)）— 让 agent 直接检索与写入 Zotero 文献库的 MCP 服务器，本 skill 的文献查找/导入环节依赖它
- [Zotero](https://www.zotero.org/) 及其 Word 集成（`xpcom/integration.js`）— 域格式的事实标准来源

## 许可证

MIT License — 详见 [LICENSE](LICENSE)。
