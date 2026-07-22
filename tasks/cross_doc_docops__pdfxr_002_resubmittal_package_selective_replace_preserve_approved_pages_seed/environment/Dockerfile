FROM harbor-codex-base:2.1.114

COPY pdfxr_002_resubmittal_package_selective_replace_preserve_approved_pages_seed.pdf /root/pdfxr_002_resubmittal_package_selective_replace_preserve_approved_pages_seed.pdf
COPY cover_values.xlsx /root/cover_values.xlsx
COPY revision_summary.docx /root/revision_summary.docx
COPY revision_note.pdf /root/revision_note.pdf
COPY ["Revised Sheet B.pdf", "/root/Revised Sheet B.pdf"]
COPY bookmark_plan.docx /root/bookmark_plan.docx
COPY original_task_description.txt /root/original_task_description.txt
COPY skills/ /skills/

COPY skills/ /root/.claude/skills/
COPY skills/ /root/.codex/skills/
COPY skills/ /root/.opencode/skill/
COPY skills/ /root/.goose/skills/
COPY skills/ /root/.factory/skills/
COPY skills/ /root/.agents/skills/
COPY skills/ /root/.gemini/skills/
