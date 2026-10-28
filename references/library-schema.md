# Excel Library Schema

Use one separate workbook per target language. Each workbook requires one metadata sheet and three knowledge sheets with the exact names and row-1 headers below. Column order is not significant. Preserve any other sheets and columns. They remain outside the standard schema and may be changed only when the user explicitly requests a supplemental workbook change under `SKILL.md`.

| Sheet | Required columns |
| --- | --- |
| `Library` | `language` |
| `Vocabulary` | `id, item, meaning, pos, collocations, example, example_translation, frequency, mastery` |
| `Grammar` | `id, item, meaning, usage, example, example_translation, frequency, mastery` |
| `Expression` | `id, item, meaning, usage, example, example_translation, frequency, mastery` |

## Default display

- On creation, hide the `id` column in each knowledge sheet. Keep the identifiers stored and readable by the script.
- On creation, fill all required row-1 headers with solid `#DDEBF7`. Preserve users' later color changes.
- On creation, add a `mastery` dropdown to each knowledge sheet from row 2 onward: `unmastered`, `unfamiliar`, `mastered`. Allow blank and use a Stop error alert for other typed values; never change stored ratings when adding the dropdown.
- Use thin black borders on all four sides of required header cells, the `Library!language` value, and required-field cells in newly inserted knowledge rows, including blank optional cells. Do not format unused rows, extra columns, or unrelated sheets.
- Ordinary reads, binding, and existing-record updates preserve cell formatting and column visibility; passive count changes only `frequency`. Apply selected display defaults to an existing workbook only on the user's explicit request.

Display choices are not data-validation requirements.

## Workbook identity

`Library!language` has exactly one nonblank data cell, in row 2. Its value is the canonical lowercase English language key chosen by the Skill (for example, `korean`). It must match the language key used to bind or open the workbook. This field identifies the workbook's target language; the binding itself remains outside the workbook.

## Shared fields

| Field | Type | Required | Rule |
| --- | --- | --- | --- |
| `id` | Text | Yes | Workbook-unique identifier generated for a new row. Never change or reuse it. No specific format is required. |
| `item` | Text | Yes | Canonical form used for lookup. Use a dictionary form for Vocabulary, a standard structure for Grammar, or a reusable form or pattern for Expression. |
| `meaning` | Text | Yes | Short meaning in the user's explanation language. |
| `example` | Text | No | A complete sentence in the target language showing the item in context, not an isolated word or collocation. Leave blank when no reliable sentence is available. |
| `example_translation` | Text | Conditional | Required when `example` is present; otherwise blank. Use the explanation language. |
| `frequency` | Integer | Yes | Value must be at least `1`. It counts distinct source units under the Frequency rules in `SKILL.md`. The Skill supplies the initial value or increment; workbook operations never infer a count. |
| `mastery` | Text | No | Blank means unrated. Allowed values are the exact lowercase strings `unmastered`, `unfamiliar`, and `mastered`. |

## Sheet-specific fields

| Sheet | Field | Type | Required | Rule |
| --- | --- | --- | --- | --- |
| Vocabulary | `pos` | Text | No | Short part-of-speech label in the user's explanation language when clear; otherwise blank. Use the same language as `meaning`. |
| Vocabulary | `collocations` | Text | No | Ordinary combinations tied to the core word, serialized with `; `. Preserve existing entries and omit duplicates. |
| Grammar | `usage` | Text | Yes | Concise grammatical function and necessary attachment rule. |
| Expression | `usage` | Text | Yes | Concise use case or communicative function. |

Keep entries suitable for search and review; do not store a full assistant explanation. Blank optional fields are empty cells, not placeholder text.

## Normalization

Normalize before lookup or write:

- Apply Unicode NFC normalization.
- Trim leading and trailing whitespace. Collapse consecutive internal whitespace to one space.
- Preserve the canonical display form in `item`. Lookup is case-sensitive after Unicode and whitespace normalization. The Skill chooses conventional casing before lookup.
- Write Vocabulary in dictionary form when the inflected form can be identified with confidence.
- Follow the target language's conventional notation for grammatical structures. For bound endings conventionally shown with a leading hyphen, retain that hyphen. Use slots such as `N`, `V`, or `A` only when they are part of the conventional reusable form, such as `N(으)로는`.
- Write Expressions in a reusable canonical form. Use `~` only when it marks a variable part of the expression.

Normalization must not merge synonyms, unrelated homographs, different grammatical functions, or entries from different sheets.

## Workbook state boundary

Store only committed library records in the workbook. Pending candidates, Candidate List numbering and edits, source-unit history, counted-item state, workbook bindings, and write-retry state remain outside the workbook. The script receives explicit operations only after the Skill has made those decisions. Direct Save and Organize use the same stored fields; they differ only in how the Skill authorizes and prepares an operation. A passive-count update (`SKILL.md` Phase 2 step 6) is a distinct, narrower operation: it may only increase `frequency` on existing rows and must not insert rows, merge collocations, fill blank fields, or change any other field.

## Identity and updates

- `type` maps directly to the `Vocabulary`, `Grammar`, or `Expression` sheet. Identify a record by `target language + sheet + normalized item`. The Skill performs linguistic normalization; the workbook operation applies the deterministic normalization above and performs exact lookup. If the context shows a different meaning or grammatical function for the same key, the Skill must not request an automatic Frequency or field update.
- If the key is absent, add one row with a new `id` and the Skill-supplied initial `frequency`. That value is the number of eligible in-scope source units, with a minimum of `1`; eligible source units collected during a Declared Session are included.
- If the key exists, do not add a row. Increase `frequency` by exactly the nonnegative integer supplied by the Skill. A delta of `0` leaves it unchanged. Filling fields and merging collocations must not implicitly increment it.
- Merge a new ordinary collocation into the matching Vocabulary row only when the Skill sends it from a confirmed Pending update or explicit save request. Do not duplicate existing values. A collocation alone does not create another Vocabulary row or its own Frequency count.
- On an existing row, fill a blank descriptive field (`meaning`, `pos`, `usage`, `example`, `example_translation`) only as part of a confirmed Pending update (Organize) or an explicit save request (Direct Save), and only when the current learning context supplies an unambiguous value. A frequency-only passive-count update (Phase 2 step 6 in `SKILL.md`) must not fill, replace, or otherwise touch any descriptive field. Replace a nonblank value only when the user explicitly requests that field change. Preserve `id` and all untouched fields.
- Never infer or automatically change `mastery`. Preserve its existing value during lookup, Frequency, collocation, and descriptive-field updates. Set it only when the user explicitly provides one of `unmastered`, `unfamiliar`, or `mastered`.

## Validation and write safety

Before any lookup or change, validate the affected workbook against the expected target-language key:

- All four required sheets and required row-1 headers exist.
- `Library!language` contains exactly one nonblank data value, and it equals the expected canonical language key.
- Every nonblank record has nonempty text values for `id`, `item`, and `meaning`.
- Every Grammar and Expression record has a nonempty text value for `usage`.
- Optional text fields are either blank or text. `example` and `example_translation` are either both blank or both nonblank.
- Each `id` is unique across the workbook.
- Each `frequency` is a non-Boolean integer greater than or equal to `1`.
- Each `mastery` is blank or exactly one of the text values `unmastered`, `unfamiliar`, or `mastered`.
- No two rows share the same `sheet + normalized item` identity within the workbook.

Ignore fully blank rows. If duplicate identities exist, block writes to the affected sheet and report the conflicting row numbers during a Commit operation. Preserve unrelated sheets, extra columns, formulas, cell formatting, filters, and existing row order. Write `frequency` as an integer and a rated `mastery` as one of the exact text values above.
