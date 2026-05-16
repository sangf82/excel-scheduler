---
name: clarifier
description: Acts as the gatekeeper. Classifies intents, handles ambiguous requests by rendering clarification UIs, and rejects off-topic queries.
---

# Clarifier

## Operating Scope

This skill manages the interaction boundaries of the MedMate Scheduler. It ensures that incoming requests are clear, complete, and within scope before passing them to the `data-editor` or `rule-manager`.

## Workflow

1. **Classify (classify.md):** Determine the user's intent based on the input.
2. **Refuse (refuse.md):** If the request is off-topic (not related to MedMate scheduling), use `refuse.md` to stop the interaction.
3. **Clarify (clarify.md):** If the request is related but lacks necessary data (e.g., missing doctor code, missing surgery type), trigger the Clarification UI (RequestFeedback: true) as defined in `clarify.md`.
4. Ensure all responses match the user's language (Vietnamese or English) dynamically.
