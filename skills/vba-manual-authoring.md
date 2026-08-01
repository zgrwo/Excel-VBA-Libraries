---
name: vba-manual-authoring
description: >
  Documentation architecture and user manual authoring rules for Excel VBA Libraries.
  Use when writing or editing any file under docs/, or when adding a new Public
  function/UDF that requires documentation updates.
last_updated: 2026-06-17
---

# VBA Libraries — Documentation Architecture & Manual Authoring

## 0. Documentation Architecture

This project uses a **layered documentation system with five distinct roles**.
Each role answers one question for one audience. Information lives in exactly
one place — all others link to it.

### 0.1 The Five Roles

| # | Document | Audience | Core Question | Type |
|---|----------|----------|---------------|------|
| 1 | `README.md` | Humans (all) | "What is this? Can it solve my problem?" | Project storefront + nav hub |
| 2 | `docs/VBA_LIB_User_Manual.md` | Human users | "I want to do X — what formula?" | Recipe-driven user guide |
| 3 | `docs/VBA_LIB_Documentation.md` | Humans + AI | "What functions exist? What are their signatures?" | **Single source of truth** for function signatures |
| 4 | `skills/*/SKILL.md` | AI assistant | "How should I write VBA/Python/SQL code?" | On-demand SOP, loaded per task type |
| 5 | `agents.md` | AI assistant | "How is the project organized? What are the red lines?" | AI constitution, **always in context** |

### 0.2 The Golden Rule: Single Source of Truth

```
Every fact has exactly ONE canonical home.
All other documents reference it by link — never by copy.
```

| Fact type | Canonical home | Referenced by |
|-----------|---------------|---------------|
| Function signature | `docs/VBA_LIB_Documentation.md` | User Manual (anchor link), SKILL.md (§ ref) |
| Coding rule | `skills/vba-SKILL.md` | agents.md (§ ref) |
| Project structure, deps (含导入顺序), test commands | `agents.md` | README (link) |
| Usage recipe, parameter behavior | `docs/VBA_LIB_User_Manual.md` | API Index (cross-link) |

### 0.3 Collaboration Flow

**When adding a new Public function:**
```
Source (.bas)
  → API Index (add signature row — canonical)
    → User Manual (add § with anchor + recipe reference)
      → Python cross-validation (tests/crossval/build_<module>.py)
        → README (formula quick-ref, only if high-value)
```
> Note: VBA `Test_ModuleName` self-tests (except SqlUtils) have been migrated to Python crossval.
```

**When AI works on a task:**
```
agents.md (always loaded — tells AI where to go)
  → task = "edit .bas"     → load skills/vba-SKILL.md
  → task = "edit SqlUtils" → load skills/sql-SKILL.md
  → task = "write test"    → load skills/python-SKILL.md
  → task = "lookup API"    → read rules/api-reference.md
  → task = "explain usage" → search rules/user-manual.md
  → task = "update manual" → load this file (manual-authoring.md)
```

### 0.4 Boundaries: What NOT to Put Where

| Don't put... | In... | Because... |
|-------------|-------|------------|
| Full function signatures | User Manual | Canonical source is API Index; link instead |
| Coding rules | agents.md | Too long for always-in-context; reference SKILL §N |
| Usage recipes | API Index | That's the User Manual's job |
| Project structure | SKILL.md | That's agents.md's job |
| Business "why" | agents.md | AI needs "how to build", not product rationale |
| Technical implementation | User Manual | Users don't need to see code internals |
| VBA rules | python/SKILL.md or sql/SKILL.md | Each SKILL file covers its own domain; cross-reference by link |
| Python/COM rules | vba/SKILL.md or sql/SKILL.md | Same — one SKILL per domain |

### 0.5 Sync Checklist (complements §12)

After any change that touches multiple docs, run the automated validation pipeline:

```bash
python tests/run_all_validation.py
```

This runs consistency checks across 4 layers:
1. **Function signatures** — counts, UDF params (all `As Variant`), no `Err.Raise`, no `Volatile`
2. **Cross-references** — `§` refs, anchors, stale markers, EN-Chinese
3. **Metadata** — module counts, `last_updated` stamps, routing table existence

All checks must pass before committing. For quick pre-commit gate (Layer 1 only, <1s):
```bash
python tests/run_all_validation.py --quick
```

---

The user manuals (`docs/VBA_LIB_User_Manual.md` and `docs/VBA_LIB_User_Manual_EN.md`)
are chapter-per-module with a uniform function-entry template. All rules below
apply to both Chinese and English editions equally.

## 1. Chapter Structure

Each chapter mirrors the `.bas` module's public-function ordering with sections
grouped by category. Chapter header format:

```
## Chapter N: ModuleName Description

One-line purpose. Key conventions.
**Module**: `ModuleName.bas`
```

Both the quick-reference table (§3) and the recipe index table (§4)
appear immediately after the module metadata line, before any function detail.

## 2. Function Entry Templates

Every linked function must have a standalone `####` heading — 🔴 combined headings
(`#### FuncA / FuncB`) are **forbidden** because they produce a single anchor
that cannot serve multiple quick-reference links. Functions appearing only in
code blocks have no anchor at all and must also receive their own `####`.

Each entry contains:
1. A one-line description
2. **VBA Usage** — signature + parameter table + example code
3. **UDF Usage** (UDF functions only) — formula + parameter table + input/output

**Full template (UDF function):**

````markdown
#### ArrayUnique

Deduplicates, preserving first-occurrence order.

**VBA Usage**

```vb
ArrayUnique(arr) As Variant
```

| Parameter | Type | Description |
|-----------|------|-------------|
| arr | Variant | 1D array |

```vb
arr = Array(3, 1, 4, 1, 5)
result = ArrayUnique(arr)
' result → Array(3, 1, 4, 5)
```

**UDF Usage**

```
=UDF_ARR_UNIQUE(array)
```

| Parameter | Type | Description |
|-----------|------|-------------|
| array | Range/Array | 1D range or array |

**Input**
|   | A |
|---|---|
| 1 | 3 |
| 2 | 1 |

`=UDF_ARR_UNIQUE(A1:A3)` ⤳

**Output**
|   | A |
|---|---|
| 1 | 1 |
| 2 | 3 |
````

---

**Compact template (related functions sharing a description):**

````markdown
#### ArrayMin

#### ArrayMax

Return the minimum/maximum. Returns Empty when no numeric values present.

```vb
ArrayMin(arr) As Variant
ArrayMax(arr) As Variant
```

| Parameter | Type | Description |
|-----------|------|-------------|
| arr | Variant | 1D array |

```vb
ArrayMin(Array(3, 1, 2))   ' → 1
ArrayMax(Array(3, 1, 2))   ' → 3
```
````

## 3. Formatting Rules

- **No `Dim`** — omit `Dim` in examples: `arr = Array(1, 2, 3)`.
- **Result arrows** — VBA results use `→`; UDF formulas use `⤳`.
- **Row/column headers** — data tables use row numbers in col 1 (`|   |`) and
  column letters in the header row (`A/B/C`).
- **Optional parameters** — shown as `[param]` in signatures; table notes "Optional"
  and the default value.
- **Separate parameter tables** — VBA and UDF each get their own table.
- **Output tables** — required for array/matrix results; scalar results
  (number, boolean, string) may use inline `⤳ 3` or `⤳ TRUE`.
- **VBA-only** — functions returning Dictionary/Object are annotated `**VBA-only**`.

## 4. Quick-Reference Table

Every chapter has a 4-column table immediately after the module metadata line:

```markdown
**Quick Reference**

| Function | Parameters | Description | Returns |
|----------|-----------|-------------|---------|
| [`ArraySort`](#arraysort) | `(arr, [ascending])` | Sort | Variant |
```

Rules:
- Function name links to its `####` heading (anchor = function name in lowercase).
- Parameters show variable names only; strip `ByRef`/`ByVal`/type annotations;
  wrap optional params in `[ ]`.

## 5. Recipe Index Table

After the quick-reference table, a 2-column table lists the chapter's three recipes:

```markdown
| Recipe | Functions Used |
|--------|---------------|
| [Score Analysis](#recipe-1-1) | `UDF_ARR_SORT`, `UDF_ARR_UNIQUE`, `UDF_ARR_COUNTIF` |
| [Dynamic Filter](#recipe-1-2) | `UDF_ARR_FILTER`, `UDF_ARR_SLICE` |
```

Rules:
- Recipe links use pure-ASCII anchors: `#recipe-{chapter}-{number}`.
- Each recipe heading is preceded by `<a id="recipe-X-Y"></a>` so the anchor
  resolves identically in GitHub and VS Code (see §6).
- Function names in the second column are backtick-wrapped, comma-separated.
- **Forbidden**: the old single-line format `**Recipes**: [A](#a) · [B](#b)`.

## 6. Anchor Safety Rules

These rules eliminate anchor breakage across GitHub and VS Code — drawn from
fixing 560 broken links (2026-06-08).

**Heading character whitelist.** Only these characters are safe in headings:

| Allowed | Behaviour across renderers |
|---------|---------------------------|
| `a-z A-Z 0-9` | Preserved |
| ` ` (space) | Converted to `-` |
| `-` (hyphen) | Preserved |
| `.` (period) | Stripped by all renderers |
| `:` (colon) | Stripped by all renderers |

**Banned characters** (use the replacement instead):

| Character | Problem | Use instead |
|-----------|---------|-------------|
| `—` (em dash U+2014) | Encoded vs stripped — renderer-dependent | `-` |
| `&` | May be URL-encoded by VS Code | `and` |
| `+` | GitHub preserves, VS Code may encode to `%2B` — anchor mismatch | space (becomes `-`) |
| `=` | GitHub strips, VS Code preserves — `#a=1` vs `#a1` | remove |
| CJK / Unicode punctuation | VS Code strips all non-ASCII | `<a id>` anchor (below) |

**Exception:** Em dashes (`—`) are permitted in `## Chapter` headings as bilingual
separators (e.g. `## Chapter 5: LinearUtils — 矩阵与线性代数计算`). Chapter headings
are not used as link targets — function entries and recipes use their own anchors.

**Recipe anchors.** Recipe headings contain Chinese/Japanese characters.
GitHub keeps CJK in anchors (`#配方-11-成绩单快速分析`); VS Code strips them
(leaving just `#11`). The only universal fix is an explicit `<a id>` before
every recipe heading:

```markdown
<a id="recipe-1-1"></a>

### Recipe 1.1 - Score Analysis
```

All recipe links then target `#recipe-X-Y` — pure ASCII, works everywhere.

**Pre-release validation.** After any manual edit, run:
```bash
py tests/utils/validate_manual_anchors.py  # ✅ 已实现
```
Target: **0 broken links across all 3 doc files** (API Index + User Manual CN + User Manual EN).

## 7. Recipe Body Format

```markdown
### Recipe N.1 — Title

**Scenario**: One-line description of the data and goal.

**Input**
|   | A | B |
|---|---|---|
| 1 | Alice | 85 |
| 2 | Bob   | 92 |

`=UDF_XXX(A1:B2)` ⤳ result
```

Rules:
- Every recipe must have **Scenario** + **Input** (data table) + formula + output.
- Input data must match the formula's column references precisely — no
  generic placeholder data like `{3,1,4,1,5}`.
- VBA recipes may use ` ```vb ` code blocks.

## 8. Source-Driven Authoring

**Trust the source, not the current manual.** Before writing a chapter, extract
the authoritative function list:
```bash
grep -n "^Public Function\|^Public Sub" src/ModuleName.bas
```

Write function-by-function against this list. Then verify coverage:
- Every `Public Function` has a standalone entry.
- Every `UDF_*` has a UDF Usage subsection.
- Old manual content is never accepted as evidence of completeness.

## 9. Runtime Validation

All numeric examples must pass cross-validation before release — 🔴 **never
hand-calculate output values**.

Three-way check:
1. **VBA output** — via Excel COM (handled by `tests/crossval/build_manual_examples.py` ✅ 已实现)
2. **Python reference** — numpy/scipy independent computation (handled by `tests/run_all_crossval.py` ✅ 已实现)
3. **Manual recorded value** — the `⤳` result in the example

A discrepancy > `1E-10` (numeric) or any string mismatch is a defect.
*Real examples:* PolyFit data mismatches, StdDev value drift, IQR calculation
errors — all invisible to manual review, caught only by runtime validation.

## 10. Bilingual Maintenance

Both manuals must stay in sync. Any change to the Chinese manual must be
reflected in the English edition simultaneously.

**Translate** (functional text): function descriptions, parameter tables,
chapter overviews, recipe scenarios, inline formula comments.

**Preserve** (do not translate): sample data values (Chinese names, product
names), VBA code comments (must match the `.bas` source), Chinese content
inside JSON/SQL/formula strings, data-content column labels like `1月_A`.

After updating the EN manual, verify no functional Chinese remains:
```bash
rg -c '[一-龥]' docs/VBA_LIB_User_Manual_EN.md
# or: grep -cP '[一-龥]' docs/VBA_LIB_User_Manual_EN.md
```
Remaining Chinese lines should all fall into the "preserve" category above.

**Recipe scenarios must be written independently per recipe** — never
copy-paste-tweak. Check for duplicates:
```bash
grep -n '^\*\*Scenario\*\*:' docs/VBA_LIB_User_Manual_EN.md | sort
```

## 11. Writing Order

Work in rounds, completing a batch of modules before starting the next:

| Round | Modules | Layer |
|-------|---------|-------|
| 1 | ArrayUtils, DictSetUtils, PivotUtils | Data |
| 2 | SqlUtils, LinearUtils, StatsUtils, RegressUtils | SQL + Math |
| 3 | StringUtils, RegexUtils, JsonUtils, XmlUtils | Text |
| 4 | DateTimeUtils, RangeUtils, FileSystemUtils | Excel/File |
| 5 | PhyChemUtils | PhysChem |

## 12. New Public Function / UDF Checklist

When adding a function, update every location below — missing even one breaks
consistency:

| # | File | What to update |
|---|------|---------------|
| 1 | `src/<Module>.bas` | Module header function list |
| 2 | `docs/VBA_LIB_Documentation.md` | Signature table + function count（唯一计数信源） |
| 3 | `docs/VBA_LIB_User_Manual.md` | Add function entry with anchor + recipe reference |
| 4 | `docs/VBA_LIB_User_Manual_EN.md` | Same as CN manual |
| 5 | `tests/crossval/build_<module>.py` ✅ 已实现 | Crossval test cases + module registration (also `build_common.py` + `build_manual_examples.py`) |
| 6 | All three doc files | Run `validate_manual_anchors.py` → 0 broken links |

**Quick verification after adding a function:**
```bash
grep -c "^Public Function\|^Public Sub" src/<Module>.bas
```

*Common misses (from 2026-06-07/08):* CN/EN editions out of sync;
anchor links broken after heading changes.
