---
name: language-learning-collector

description: Builds and maintains a personal per-language Excel library of vocabulary, grammar, and expressions from the user's foreign-language study. Use when the user asks, in any language, to organize what was learned in a study session, to analyze material and collect knowledge points for later organization (confirm the language-study context when the material arrives), to save a specific word, grammar point, or expression, or to choose, switch, or bind the library's workbook location. Do not use for ordinary language questions or practice, one-off translation, review plans, quizzes, flashcards, or note and file saving unrelated to language study, unless the user also asks to organize or save language points.
---

# Language Learning Collector

Turn language-study conversations into a structured personal library (one Excel file per language). Use the model for language understanding, classification, and candidate selection. Use `scripts/excel_library.py` for every binding-registry and standard-schema workbook operation. Handle explicitly requested non-schema changes under Explicit Workbook Extensions. If the script reports a missing dependency, follow Dependency Setup. If the script is missing or still cannot run afterward, report that the Skill installation is incomplete and stop; do not substitute ad hoc logic for standard operations.

Read [Workbook schema](references/library-schema.md) before the first step in a task that creates, validates, reads, looks up, inserts, or updates a workbook.

## Core Principles

1. **Teach first, track silently** — Answer the learning question normally. Do not ask whether to save items or mention Pending items, frequency tracking, or other library operations unless relevant to the user's request.
2. **Confirm candidates before committing** — Items identified from the conversation enter Pending. Commit only those the user confirms in the Candidate List. A Direct Save request skips this confirmation.
3. **Be selective** — Collect what the user asked about, struggled with, or what is clearly reusable. Do not extract every unfamiliar word from a source text.
4. **Never modify mastery** — Only the user sets it.
5. **Report verified writes only** — Report an item as saved or updated only after the write succeeds.
6. **One Language, One Workbook** — Use a separate workbook for each target language. Detect the target language from the current context and write to its corresponding workbook; ask only if the language is ambiguous.

## Execution Workflow
### Phase 0: Detect Mode

User instructions may be in any language. Interpret meaning, not keywords.

| User input | Mode | Go to |
| --- | --- | --- |
| Asks to organize what was covered, or asks what is worth keeping ("organize what we covered", "整理本轮知识") | A: Organize | Phase 2 |
| Says before studying that language points should be collected as they go | B: Declared session | Phase 1 |
| Explicitly asks to save identifiable items, including to Excel ("save 차지하다", "add -는 가운데 to the library") | C: Direct save | Phase 3 |
| Chooses or switches the folder or file of a language library, or points to an existing workbook to use as one | D: Bind (supporting) | Route the Workbook |

If a message contains multiple explicit intents, complete any required Bind first, Direct Save the specifically named items, then Organize the remaining in-scope material. Do not re-propose directly saved items.

A message that does not match any mode above is ordinary conversation. Answer it normally. If it is a study message and a declared session is active, also update the Phase 1 Pending Candidates silently. Otherwise, the skill does nothing.

For an advance request to analyze material and later organize “knowledge points,” answer without promising language-point collection; wait for the material before entering Phase 1. Treat material in another language, or an explicit study target, as language study. If the material is in the conversation language with no stated study target, answer the analysis and ask whether the user is studying that language before collecting; if not, organize ordinary content.

For other ambiguous intents, or when “save” or “record” is used outside a clear language-learning context, ask once: “Would you like to save this to your language library, or as a general note/file?” Route the message after the user clarifies.

If a save request points only to a passage and does not identify the learning items, ask which learning points to save.

Determine two languages for the round:

- **Target language** — the user's explicit choice, then the material under discussion, then the language already established in this round. If a mixed-language item is still ambiguous, ask before assigning it to a library.
- **Explanation language** — the user's stated language, else the conversation language. Use it for meanings and translations.

### Phase 1: Declared Session

Phase 1 begins when the user declares that upcoming study material should be collected. After each study message, answer normally, then silently:

1. Record the source unit (see Frequency) and the items it contains.
2. Add worth-keeping items to Pending Candidates (see Candidate Selection). Merge repeats by `language + type + normalized item`, keeping useful meaning, usage, collocations, and examples.

Do not read or write the workbook, or mention the Pending Candidates. Pending Candidates exist only in the conversation.
When the user asks to organize or indicates that the declared study session has ended, end Phase 1 and enter Phase 2 step 1 using the accumulated Pending Candidates.

### Phase 2: Organize

1. **Scope** — Use study material added since the previous completed Organize, together with any failed writes and uncounted source units kept for retry. For the first Organize, use the conversation from the start. If required context is unavailable, ask the user to repost it; do not infer.
2. **Prepare candidates** — Do not touch the workbook.
- If Pending Candidates already exists, start with them and scan the in-scope material only to fill gaps.
- Otherwise, identify source units and candidates from the in-scope material now.
Merge repeats by `language + type + normalized item`. For each candidate, count the distinct in-scope source units in which it appears (Frequency rules). If a new collocation belongs to a core word, include it in that Vocabulary candidate.
3. **Show** the Candidate List with each item's count and ask whether it covers what the user wants to keep. If there are no candidates, say so in one line and go to step 6.
4. **Apply the user's edits and selection** using the table below. Do not continue to step 5 until the user clearly asks to save or submit the list, or confirms that editing is complete.
5. **Commit** the selected items through Phase 4. Resolve the workbook there through Route the Workbook. If nothing is selected, skip to step 6.
6. **Passive count** —  Use `scripts/excel_library.py` for each target language in scope with a usable bound workbook. Find stored items that appear in the in-scope source units but were not submitted in step 5. Apply their eligible frequency increments under the Frequency rules in one workbook update batch without asking the user. Skip this step when the required source material is unavailable.
7. **Report** with the Save Result format and close the current Candidate List. Remove committed or user-removed candidates from Pending. Mark a source unit as counted for an item only after the corresponding write succeeds. Keep failed writes and skipped counts, together with their uncounted source units, for retry. 

| User says | Action |
| --- | --- |
| save all | Commit all displayed items |
| only 1, 2, 4 | Keep 1, 2, and 4; remove the rest. Commit only if the message also expresses completion |
| drop 3 | Remove 3 and keep the rest Pending. Commit only if the message also expresses completion |
| all except 2 | Remove 2 and keep the rest. Commit only if the message also expresses completion |
| only Grammar and Expression | Keep those types and remove the rest. Commit only if the message also expresses completion |
| also include X (not in list) | Add X as a candidate with its in-scope count. Keep it Pending unless the same message also asks to save or submit the list, or says the editing is complete |
| change 5 to ~로 나타나다 | Edit only. Reclassify if the type changes. Item stays Pending |
| edit and add in one message | Apply the edits and additions. Commit only if the message also expresses completion |
| change category, form, or example | Edit only |

- Adding, removing, or editing items only updates the Candidate List. Keep existing item numbers unchanged and give added items the next available numbers. 
- Commit only when the user clearly asks to save or submit the list, or confirms that editing is complete. If this is unclear, ask whether the list is ready to save. Until then, keep the remaining items Pending.
- If a number is out of range, duplicated, or ambiguous, ask one short question and change nothing.

### Phase 3: Direct Save

An explicit request to save an identifiable item skips candidate confirmation and the passive-count step. If the referent is unclear ("save this expression"), ask which item. Use Phase 2 step 1 to determine the scope, calculate the item's frequency change under the Frequency rules, then continue to Phase 4. After a successful commit, remove any matching item (same language + type + normalized item) from Pending and from any displayed Candidate List, so a later Organize does not re-propose or recount it.

### Phase 4: Commit

Commit the selected items (Organize) or the named items (Direct Save), grouped by target language:

1. Use `scripts/excel_library.py` for the standard operations below. Resolve the language's bound workbook and validate it against the schema. Follow Route the Workbook if user input is required. Handle an explicitly requested non-schema change only after the applicable standard operation succeeds.
2. Classify and normalize the item, then look up `language + type + normalized item`.
3. **Not found** → insert a row. `frequency` is the item's in-scope count (minimum 1). Leave mastery blank unless the user explicitly supplied a schema-valid rating. Fill the other fields per Record Writing.
4. **Found** → apply only the Existing Record Rules. `frequency` += in-scope count.

Use one `read_workbook()` result per language to look up exact sheet + `normalize_item(item)` identities and prepare explicit `insert` or `commit_update` operations. Submit them with `batch_execute_operations()`; the standard batch saves each workbook at most once. Languages commit independently. Report only successful per-item results from the completed batch.

**Before returning the result, verify:** the workbook matches the target language, schema validation passed, every item was looked up before writing, discarded items were not written or counted, no source unit was counted twice for the same item, mastery was preserved, and source units were marked counted only after a successful write. Treat an item as added or updated only after its write succeeds. Keep failed items unsaved.

## Rules

### Candidate Selection

Prioritize in this order:
1. Items the user asked about or wanted analyzed in detail
2. Items the user clearly found unfamiliar
3. Structures the user asked to have explained
4. Items with high reuse value

Collect only what the user is actually studying. Do not add every word or phrase in a passage. Do not add an item that appears only in assistant messages unless the user later asks about or studies it.

### Classification

| Type | Definition | Examples |
| --- | --- | --- |
| **Vocabulary** | Core word or lexical item with independent meaning, in standard dictionary form | 높아지다, 차지하다, 원동력 |
| **Grammar** | Structure with a grammatical function, attachment rule, or syntactic effect | -면서, N(으)로는, -(으)ㄹ 예정이다 |
| **Expression** | Conventional idioms, proverbs, or fixed sayings learned as a whole; not ordinary collocations or freely composed sentence frames | 발 벗고 나서다, 식은 죽 먹기, 티끌 모아 태산 |

- For an ordinary collocation, identify the core word being studied and put the collocation in that Vocabulary candidate's `collocations` field. Do not create a separate row for the collocation.
- Choose the core word according to the user's question or explicit selection.
- Grammar covers grammatical structures; Expression covers conventional fixed expressions. Do not classify a phrase as Expression merely because it is complete or reusable.
- Extract component words only when they are learning targets. Add derived words only when actually studied; do not treat them as inflected forms of the core word.
- Example sentences belong in fields. A round may contain no Expression candidates.

### Frequency

`frequency` counts distinct source units in which an item appears. A source unit is self-contained material introduced for study, such as an exercise, exam question, reading passage, listening transcript, or example set. Treat the complete unit (stem, passage, options) as one source. A sentence quoted only to ask about a word, grammar point, or expression is question context, not a source unit. Messages and follow-up questions are not counting units.

| Situation | Effect |
| --- | --- |
| Selected candidate or directly saved item not in the library | Insert with `frequency = N`, where N is the number of in-scope source units containing it, excluding units already counted for it (minimum 1) |
| Selected candidate or directly saved item already in the library | `frequency += N` (N may be 0: frequency unchanged) |
| Stored item that is not a candidate in the current Organize appears in N distinct in-scope source units not yet counted for it | frequency += N |
| An inflected or conjugated form of an item appears (높아지면서 → 높아지다) | Restore it to its standard form (높아지다) for the Candidate List and for matching, and count it as that item when clearly the same. Judge each item independently: one surface form can count for several items (높아지면서 → 높아지다 and -면서) |
| Follow-up question about the current source, or the same source reposted in this conversation | No additional count |
| Assistant output, Pending items, review or edit commands, and save or confirm commands | No count |
| Workbook unresolved or write fails | Do not count or claim a count. The unit stays in scope for the next organize |

Match items to source units using the Normalize rules. Do not count coincidental substrings, homographs with another meaning, or uncertain analyses.

Track source units and the items already counted for them in conversation context. Reposting the same source in the same conversation does not create another count. A new conversation starts fresh source tracking, so the same material counts again if the user presents it there as study material. Treat materially different material, or material the user explicitly identifies as a new exercise or passage, as a new source unit. 

### Normalize and Deduplicate

- Identity is `language + type + normalized item`. The same normalized string may legitimately exist under different types.
- Reduce Vocabulary to its dictionary form only when clearly the same word (높아졌습니다, 높아지고, 높아지면 → 높아지다). Otherwise keep the user's form.
- Treat a normalized match as the same record only when it has the same lexical meaning or grammatical function. Do not match by substring alone, merge unrelated homographs, or assume an uncertain morphological analysis.
- No synonym merging, fuzzy matching, or embeddings.

### Existing Record Rules

- Apply the update constraints in [references/library-schema.md](references/library-schema.md).
- Increment `frequency` only when the Frequency table permits it.
- Merge a collocation only from a confirmed Pending update or an explicit save request.
- Preserve `mastery` unless the user explicitly supplies a schema-valid rating.
- Let the user choose `mastery` from the workbook dropdown or explicitly name a schema-valid rating. Leave it blank if unrated; do not require a rating to save.

### Record Writing

Follow the schema field rules. Keep fields short and searchable. Never store full AI explanations.

Store ordinary word combinations in `collocations`. Use `example` for a complete sentence showing the item in context, not a bare word or collocation. If the learning context provides no reliable sentence, leave both `example` and `example_translation` blank.

Use the schema's default display for new workbooks and borders for new rows. Preserve existing display during normal operations; call `format_workbook()` with only the requested options when the user explicitly asks to change it. Use `mastery_dropdown=True` to add rating choices to an existing workbook without changing its ratings.

### Explicit Workbook Extensions

When the user explicitly requests a field or workbook change outside the standard schema, complete the applicable standard operation with `scripts/excel_library.py` first, then apply only the requested extension with workbook-preserving spreadsheet tooling. For an extension-only change to an existing record, validate and locate it without changing `frequency` or other standard fields. Verify both stages and report any partial failure accurately.

### Route the Workbook

Use `scripts/excel_library.py` to maintain one persistent workbook binding per canonical language key. Each binding stores the workbook's validated absolute path; rebinding replaces the previous path. `<language>_learning_library.xlsx` is only the default filename offered during first-time creation. After binding, the stored path, including its filename, is authoritative.

Before writing, check the available permission scope. Request missing workbook and registry directory access through the host's approval mechanism; saving needs temporary-file creation and destination replacement. If denied, keep affected items unsaved until access is granted. Do not retry unchanged permission failures, change paths to bypass them, or alter system permissions.

1. Normalize the target language to one lowercase English key (Korean, 한국어, 韩语 → korean).

2. Check the binding for that language:
| State | Action |
| --- | --- |
| The user selects an existing workbook | Validate it. If it is valid and writable, bind its absolute path and use it. Otherwise, ask the user to repair it or select another workbook |
| The bound file is available and the user asks to move or rename it | Locate it through the binding, move or rename it as requested, validate it at the new path, then update the binding |
| A binding exists and the bound file is available | Validate and use the bound file |
| A binding exists but the bound file is missing | Search the bound path's parent directory and subdirectories for the bound filename, then handle the result in step 3 |
| No binding exists | Ask for the destination folder and whether to use the default filename or a user-provided filename. Search that folder and its subdirectories for the selected filename before creating anything, then handle the result in step 3 |

3. Handle the results of the searches for a missing binding or first-time binding:

- Exactly one valid, writable workbook: bind its absolute path and use it.
- Several valid, writable workbooks: show their absolute paths and ask the user to choose, then bind and use the selected workbook.
- Matching files exist but none are valid and writable: ask the user to repair one or choose another workbook.
- No match for a missing binding: ask the user to locate the original workbook or confirm creation of a replacement.
- No match on first use: create `<chosen folder>/<chosen filename>`, then validate, bind, and use it.

Never overwrite a workbook or move, rename, or replace one without the user's explicit request or confirmation.

For a Bind-only request, report the target language and absolute bound path after the binding is persisted, then stop.

## Output Formats

**Workbook Location Request** (render in the user's language):

```
No <target language> learning library was found. Which folder should it be saved in? It will be named <default filename> unless you give another name.
```

**Candidate List** (header and closing line in the user's language; ×N is the in-scope source-unit count):
```
4 knowledge points extracted from this round:

1. 차지하다 | Vocabulary | To occupy; to take up | ×2
2. N(으)로는 | Grammar | As for...; in terms of... | ×1
3. 그 뒤를 잇다 | Expression | To follow closely behind | ×1
4. 높아지다 | Vocabulary | To rise; + collocation 가능성이 높아지다 | ×3

Does this cover what you want to keep? Tell me what to add, drop, or change.
```

**Save Result** (report only verified outcomes):

```
Workbook result for <file>:
- Added (2): 차지하다 ×2 (Vocabulary), -면서 ×1 (Grammar)
- Updated (1): 높아지다 +3, added collocation 가능성이 높아지다
- Also counted (2): 원동력 +1, 나타나다 +1
- Discarded (1): 4
- Unsaved (0)
```

## Dependency Setup

Use the same Python interpreter to check dependencies, install them, and run `scripts/excel_library.py`.

If importing the script reports that `openpyxl` is missing:

1. Show the user the exact install command reported by the script and ask for permission to run it. Do not install packages silently.
2. If the user approves, run that command once, then retry the original workbook operation once.
3. If the user declines, the installation fails, or the retry still fails, report the problem and stop the workbook operation.

If `requirements.txt` is missing, report that the complete Skill folder must be reinstalled. Do not guess dependency versions or use a different Python interpreter.

## Supporting Files

| File | Purpose | When to Use |
| --- | --- | --- |
| [Workbook schema](references/library-schema.md) | Authoritative workbook structure: required sheets, columns, field rules, record identity, validation rules, and permitted updates | Read before a task creates, validates, reads, looks up, inserts, or updates any workbook |
| `scripts/excel_library.py` | Manage the persistent language-to-workbook registry and perform standard-schema workbook operations: bind, resolve, search, move or rename, create, validate, read, look up, insert, update, and format | Run for every binding-registry or standard-schema operation required by Route the Workbook, Phase 2 step 6, or Phase 4 |
