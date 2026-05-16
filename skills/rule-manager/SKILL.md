---
name: rule-manager
description: Manages scheduling logic, constraints, and formulas. Updates formulas in the active Excel file and records persistent rules in the global memory file.
---

# Rule Manager

## Operating Scope

This skill modifies the formulas and logic in the `TÍNH TOÁN` sheet of the **Active Project File** and updates the persistent memory.

- Applies changes to the scheduling algorithm.
- Updates daily patient caps or doctor workload limits.

## Workflow

1. **Update Active File:** When a user requests a rule change (e.g., "Ghép gan chỉ tối đa 1 ca 1 ngày"), use the Excel MCP to update the corresponding formula in the `TÍNH TOÁN` sheet of the active Excel file.
2. **Update Persistent Memory:** Append the new rule to the memory file located at `~/.codex/memories/medmate-scheduler.md`. This ensures that when new projects are created, the agent can reference these established rules.
3. **Reference Materials:** Refer to `rules.md` and `algorithm.md` in this directory for the baseline logic and formulas.
