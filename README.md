# Scorecard

## Apply indicator input messages

To update `Endpoint Security Test1.xlsx` so indicator rows (`A2:A57`) show an input message:

```bash
python apply_indicator_input_messages.py
```

This script sets:
- **Input title** = row's **Measure** (column `B`)
- **Input message** = row's **Reference** (column `C`)

It also preserves existing list validations on `D2:D57` and `E2`.
