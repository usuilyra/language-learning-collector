# Language Learning Collector

**简体中文** | [English](README_EN.md)

用 AI 学外语时，学完说一句「整理本轮知识」，这个 skill 就会把值得记的词汇、语法和表达整理进你本地的 Excel 知识库，并记录每个知识点出现的频率，方便你之后按频率和掌握度复习。

无论是考前拿真题突击，还是从各种语言材料里慢慢积累，和 AI 聊过的内容都不会散落在对话里，而是留成一份属于你自己的学习库。

## 你会得到什么

每种目标语言对应一个 Excel 文件，包含四个工作表：

| 工作表 | 保存的内容 |
| --- | --- |
| `Library` | 这个文件对应的目标语言 |
| `Vocabulary` | 核心词、释义、词性、常见搭配和例句 |
| `Grammar` | 语法结构、核心用法和必要的接续说明 |
| `Expression` | 需要整体记忆、可以整体复用的表达和句型 |

两个重要字段：

- **`frequency`**：知识点出现在多少个不同的学习材料中，例如不同的阅读、听力稿或练习题。整理时自动记录和更新。
- **`mastery`**：由你自己标记的掌握度，新知识点默认为空
  - `unmastered`：未掌握
  - `unfamiliar`：有印象
  - `mastered`：已掌握

把 `frequency` 和 `mastery` 放在一起看，就能找出高频但还不熟的重点。完整字段说明见 [library-schema.md](references/library-schema.md)。

## 学习库文件放在哪里

第一次保存某种语言的知识点时，Skill 会询问学习库放在哪个文件夹，以及是否使用默认文件名（如果学习韩语，默认文件名为 korean_learning_library.xlsx）

如果你已经有学习库，可以直接告诉 Skill 文件位置，比如：

```text
使用 D:\Language Libraries\my-korean.xlsx 作为我的韩语学习库。
```

之后想移动、重命名或切换学习库，直接说明目标位置即可。Skill 会为每种语言记住当前使用的文件。


## 使用场景

### Case 1：边学边收集，学完统一整理

适合一边和 AI 学习，一边让 Skill 留意值得保存的内容。

**开始前**，告诉 Skill 你要学什么：

```text
$language-learning-collector 接下来我要学习韩语。请留意值得保存的词汇、语法和表达，等我说“整理”时再列出来。
```

**学习中**，照常贴入阅读材料、听力稿或练习题并提问。Skill 会正常回答，同时在后台记下可能值得保留的知识点，不会打断你。

**学完后**，说：

```text
整理本轮知识点。
```

Skill 会列出候选知识点，并标出每个知识点出现在几个学习材料中（×N）：

```text
1. 줄어들다 | Vocabulary | 减少 | ×1
2. -면서 | Grammar | 一边……一边…… | ×1
3. 발 벗고 나서다 | Expression | 主动出面帮忙 | ×1
```

可以修改候选知识点列表，比如：

```text
删掉 3，把 1 的释义改成“逐渐减少”，再加入 차지하다。
```

**确认保存**：

```text
按修改后的清单保存。
```

Skill 会把确认后的新知识点存入学习库，同时自动更新本轮材料中已有知识点的 frequency。

### Case 2：学完之后，回头整理

适合学习前没有想到要收集，学完才想整理的情况。比如：

```text
整理刚才那段韩语听力里值得保留的知识点。
```

Skill 会从本轮学习内容中挑出值得保留的词汇、语法和表达，先展示候选清单，再由你增删、修改和确认。

### Case 3：直接保存指定知识点

适合你已经知道要保存什么的情况。明确指定的知识点不需要候选清单，会直接保存：

```text
把 줄어들다（减少）保存到韩语学习库。
```

你也可以同时指定类型、释义、搭配或例句：

```text
把 -면서 保存为韩语语法，意思是“一边……一边……”。
```

当你只说保存某个单词或语法，没有给出释义的时候，AI 会自动查找释义或者用法，补充到个人学习库中。

### Case 4：在 Excel 中补充和调整

学习库就是一个普通的 Excel 文件，你可以自己排序、筛选，也可以增加自己的列和批注。Skill 只写入标准字段，会保留你添加的其他列、工作表和格式。

如果想修改某条记录，也可以直接告诉 Skill：

```text
把 줄어들다 的释义改成“逐渐减少”。
```

Skill 只修改你明确指出的字段，其他内容和掌握度保持不变。


## 安装

### 方式一：自动安装（推荐）

打开终端，粘贴并运行：

```bash
npx skills add usuilyra/language-learning-collector -g
```

按照提示选择你正在使用的 Agent。`-g` 表示安装到用户级目录，之后可以在不同项目中使用。

或者把下面这句话发给你的 Agent：

```text
帮我安装这个 Skill：https://github.com/usuilyra/language-learning-collector
```

### 方式二：手动安装

在 [GitHub 仓库](https://github.com/usuilyra/language-learning-collector) 点击 **Code → Download ZIP**。解压后，将文件夹重命名为 `language-learning-collector`，再放进你的 Agent 的 Skills 目录。

或者使用 Git 直接安装。例如 Claude Code：

```bash
git clone https://github.com/usuilyra/language-learning-collector.git ~/.claude/skills/language-learning-collector
```

其他 Agent 可以替换为对应的用户级路径：

| Agent | 目标路径 |
| --- | --- |
| Claude Code | `~/.claude/skills/language-learning-collector` |
| Codex | `~/.agents/skills/language-learning-collector` |
| Gemini CLI | `~/.gemini/skills/language-learning-collector` |
| OpenCode | `~/.config/opencode/skills/language-learning-collector` |
| Cursor | `~/.cursor/skills/language-learning-collector` |
| Kiro | `~/.kiro/skills/language-learning-collector` |


首次使用时，如果缺少 Python 依赖，Agent 会先征求你的允许，再安装依赖并继续。

完成后，重新打开 Agent 或新建一个会话。可以这样开始：

```text
请使用 language-learning-collector。接下来我要学习韩语，
请留意值得保存的知识点，等我说“整理”时再列出来。
```
