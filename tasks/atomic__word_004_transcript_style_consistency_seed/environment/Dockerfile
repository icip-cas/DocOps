FROM harbor-claude-code-base:2.1.114

COPY word_004_transcript_style_consistency_seed.docx /root/word_004_transcript_style_consistency_seed.docx
COPY original_task_description.txt /root/original_task_description.txt
COPY skills/ /skills/

# Copy skills to common agent-specific locations.
COPY skills/ /root/.claude/skills/
COPY skills/ /root/.codex/skills/
COPY skills/ /root/.opencode/skill/
COPY skills/ /root/.goose/skills/
COPY skills/ /root/.factory/skills/
COPY skills/ /root/.agents/skills/
COPY skills/ /root/.gemini/skills/
