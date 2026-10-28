# Language Learning Collector

[简体中文](README.md) | **English**

When learning a foreign language with AI, by saying “Summarize this session”, this skill will extract noteworthy vocabulary, grammar, and expressions into your local Excel knowledge base—tracking the frequency of each point so you can easily review them later by frequency and mastery level.

Whether you are cramming with past exam papers or accumulating everyday materials, your conversations with AI won't fade away into scattered chat logs. Instead, they become a personalized study repository that is uniquely yours.

## What You Get

Each target language has its own Excel workbook with four worksheets:

| Worksheet | What it stores |
| --- | --- |
| `Library` | The target language for this workbook |
| `Vocabulary` | Core words, meanings, parts of speech, common collocations, and example sentences |
| `Grammar` | Grammar structures, their main uses, and any necessary attachment rules |
| `Expression` | Expressions and sentence patterns that are best learned and reused as a whole |

Key fields for smart review:

- **`frequency`**: Tracks how often an item appears across different sources (auto-updated).

- **`mastery`**: Your current progress (unmastered | unfamiliar | mastered).

Filter by both to quickly target high-frequency items you haven't mastered yet. See [library-schema.md](references/library-schema.md) for the full schema.

## Where the Learning Library Is Stored

The first time you save an item for a language, the Skill will ask which folder should hold the learning library and whether to use the default filename. For Korean, for example, the default filename is `korean_learning_library.xlsx`.

If you already have a workbook, you can give the Skill its path:

```text
Use D:\Language Libraries\my-korean.xlsx as my Korean learning library.
```

If you later want to move, rename, or switch the workbook, tell the agent the new destination. It remembers the selected workbook for each language.

## Usage

### Case 1: Learn First, Organize Later

Use this workflow to study naturally without interrupting your flow. The Skill tracks useful items in the background and only writes to your workbook upon confirmation.

**Initialize**: Tell the Skill your target language before you start.

```text
$language-learning-collector I'm studying Korean. Please keep track of useful vocabulary, grammar, and expressions, and show them when I say “organize.”
```

**Study normally**: Paste your readings, transcripts, or exercises and ask questions. The Skill passively collects items without saving them yet.

**Review & Refine**: When finished, ask to organize your session. The Skill generates a candidate list with item frequency (`×N`):

Sample Output:
```text
1. 줄어들다 | Vocabulary | to decrease | ×1
2. -면서 | Grammar | while doing; at the same time | ×1
3. 발 벗고 나서다 | Expression | to step up and help proactively | ×1
```

Optional: Edit the candidate list directly in the chat (e.g., "Remove 3, update the definition of 1, and add 차지하다").

**Save**: Confirm to update your workbook.

```text
Save the revised list.
```
The Skill writes new items to the workbook and auto-increments frequencies for existing ones.

### Case 2: Organize the Current Study Session

Use this when you did not start collecting items beforehand but want to organize what you just studied:

```text
Organize the useful items from the Korean listening exercise we just studied.
```

The Skill extracts key vocabulary, grammar, and expressions from your recent conversation, presents a candidate list, and awaits your approval.

### Case 3: Save a Specific Item Directly

Skip candidate review and save specific items directly:

```text
Save 줄어들다 (to decrease) to my Korean learning library.
```

You can also specify its type, meaning, collocation, or example sentence:

```text
Save -면서 as Korean grammar meaning “while doing; at the same time.”
```

If you ask to save a word or grammar point without providing a meaning, the AI can look up its meaning or usage and add the necessary information to your learning library.

### Case 4: Add to or Adjust the Excel Workbook

The learning library is a regular Excel workbook. You can sort and filter it, add your own columns or comments, and adjust the formatting. The Skill writes only to its standard fields and preserves other columns, worksheets, and formatting you add.

You can also ask the Skill to update a record:

```text
Change the meaning of 줄어들다 to “to gradually decrease.”
```

The Skill changes only the field you explicitly request and leaves other content and the mastery rating unchanged.

## Installation

### Option 1: Automatic Installation (Recommended)

Open a terminal and run:

```bash
npx skills add usuilyra/language-learning-collector -g
```

Follow the prompts to select the Agent you use. The `-g` flag installs the Skill at the user level so it is available across different projects.

Alternatively, send this message to your Agent:

```text
Install this Skill for me: https://github.com/usuilyra/language-learning-collector
```

### Option 2: Manual Installation

Open the [GitHub repository](https://github.com/usuilyra/language-learning-collector), then select **Code → Download ZIP**. Extract the archive, rename the folder to `language-learning-collector`, and place it in your Agent's Skills directory.

Or install it directly with Git. For example, for Claude Code:

```bash
git clone https://github.com/usuilyra/language-learning-collector.git ~/.claude/skills/language-learning-collector
```

For another Agent, replace the destination with the corresponding user-level path:

| Agent | Destination |
| --- | --- |
| Claude Code | `~/.claude/skills/language-learning-collector` |
| Codex | `~/.agents/skills/language-learning-collector` |
| Gemini CLI | `~/.gemini/skills/language-learning-collector` |
| OpenCode | `~/.config/opencode/skills/language-learning-collector` |
| Cursor | `~/.cursor/skills/language-learning-collector` |
| Kiro | `~/.kiro/skills/language-learning-collector` |

If a required Python dependency is missing the first time you use the Skill, the Agent will ask for your permission before installing it and continuing.

After installation, restart your Agent or open a new session. You can begin with:

```text
Please use language-learning-collector. I'm going to study Korean.
Keep track of useful items and show them to me when I say “organize.”
```
