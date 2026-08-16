## Summary

<!-- Brief description of what this PR does -->

## Checklist

- [ ] VBA code follows [SKILL.md](skills/vba-SKILL.md) conventions
- [ ] All parameters are `As Variant` (UDF compatibility — no `#VALUE!` from Range inputs)
- [ ] New Public functions have UDF wrappers (naming: `UDF_*`)
- [ ] Error handling: UDF → `CVErr()`, VBA functions → `Err.Raise`
- [ ] `python tests/run_all_validation.py --quick` passes locally
- [ ] Full validation (`python tests/run_all_validation.py`) passes locally
- [ ] Cross-validation tests updated (if new functions added)
- [ ] Integration tests updated (if Range-path functions added)
- [ ] [API Index](rules/api-reference.md) updated (if signatures changed)
- [ ] [User Manual (CN)](rules/user-manual.md) updated (if new UDFs)
- [ ] [User Manual (EN)](docs/VBA_LIB_User_Manual_EN.md) updated (if new UDFs)
- [ ] Manual anchors validated (`python tests/utils/validate_manual_anchors.py`)

## Type of change

- [ ] Bug fix
- [ ] New feature / function
- [ ] Performance improvement
- [ ] Documentation update
- [ ] Test improvement
- [ ] Refactoring (no functional change)

## Related issues

<!-- Link to issues this PR fixes: Fixes #123 -->
