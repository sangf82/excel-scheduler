# Clarification Workflow

When a user's request is ambiguous, missing critical information (e.g., missing doctor ID, missing surgery type, or unclear intent), or could be interpreted in multiple ways, you MUST pause the write flow and use Codex's native "Ask User" UI to get clarification.

**How to trigger the native Ask User UI:**
Instead of just asking the question in the chat, you must create a new Artifact (using the `write_to_file` tool with `IsArtifact: true`) to present your questions. 
Crucially, you MUST set `RequestFeedback: true` in the `ArtifactMetadata` when creating this file.

1. **Artifact Name:** `clarification_needed.md`
2. **Artifact Type:** `task` or `other`
3. **Artifact Content:** 
   Detect the user's language (Vietnamese or English) and write the explanation and bullet points **ONLY** in that language. Do not mix languages.
   Example (if user speaks Vietnamese):
   ```markdown
   **Yêu cầu chưa rõ ràng**
   Bạn chưa cung cấp mã Bác sĩ cho lịch trực.
   
   Vui lòng làm rõ:
   - Bác sĩ nào sẽ phụ trách?
   ```
4. **Metadata:** Ensure `RequestFeedback` is set to `true`.

**Rules:**
1. Do NOT proceed with any modifications to the Excel file or `build_template.py` until the user has provided the required information through the feedback UI.
2. Only ask the absolute minimum required to safely execute the write flow.
