# Contributing to Excel VBA Libraries

Thanks for your interest in contributing! This document outlines the process.

## Getting Started

### Prerequisites

- **Windows** with Microsoft Excel (2016 or later, 64-bit recommended)
- **Python 3.10+** for running tests
- **Git** for version control

### Setup

```bash
git clone https://github.com/zgrwo/Excel-VBA-Libraries.git
cd Excel-VBA-Libraries

# Install Python test dependencies
pip install numpy scipy comtypes openpyxl

# Install the pre-commit hook
bash scripts/install_hooks.sh
```

### Importing into Excel

Open `docs/VBA_Libraries.xlsm`, or import modules manually from `src/` in this order:

1. **VBA-Core**: `VariantKit.cls` → `ArrayOps.cls` → `DictProxy.cls`
2. **src/ modules**: Any order after VBA-Core is loaded

## Development Workflow

### 1. Pick an issue

Check the [issues](https://github.com/zgrwo/Excel-VBA-Libraries/issues) for something to work on, or open a new issue to discuss your idea first.

### 2. Create a branch

```bash
git checkout -b feature/your-feature-name
# or: git checkout -b fix/your-bug-fix
```

### 3. Write code

Follow the coding standards in [`skills/vba-SKILL.md`](skills/vba-SKILL.md). Key rules:

- **All Public function parameters must be `As Variant`** — otherwise Range inputs cause `#VALUE!`
- **Every Public function needs a UDF wrapper** (`UDF_*`) for worksheet use
- **Error handling**: UDF wrappers → `CVErr()`; internal VBA functions → `Err.Raise`
- **Use `VariantKit.NormalizeInput()`** to handle Range vs Array inputs uniformly
- **New doc page**: For new UDFs, add user-facing docs to [User Manual CN](rules/user-manual.md), [EN](docs/VBA_LIB_User_Manual_EN.md), and [API Index](rules/api-reference.md)

### 4. Test

Run tests in order of speed:

```bash
# Fast: signature & metadata validation (<2s)
python tests/run_all_validation.py --quick

# Full: all 4 layers (~10s)
python tests/run_all_validation.py

# Cross-validation: numpy/scipy reference comparison (requires Excel, ~5-10 min)
python tests/run_all_crossval.py

# Integration: COM Range-path tests (requires Excel, ~2-3 min)
python tests/utils/integration_test_all_modules.py
```

**Always run `run_all_validation.py --quick` before committing.**

### 5. Commit

Follow [conventional commits](https://www.conventionalcommits.org/):

```
feat: add FOO function to ArrayUtils
fix: handle empty array in SORT_ASC
docs: update user manual for new ArrayUtils functions
test: add crossval cases for ArrayUtils.FOO
refactor: extract shared helper from StringUtils
```

### 6. Open a Pull Request

Push your branch and open a PR. Use the PR template — it guides you through the checklist.

A maintainer will review your PR. CI checks must pass before merge.

## Project Structure

See [agents.md](agents.md) for the full module map and dependency graph.

## Questions?

Open a [discussion](https://github.com/zgrwo/Excel-VBA-Libraries/discussions) or ask in an issue.
