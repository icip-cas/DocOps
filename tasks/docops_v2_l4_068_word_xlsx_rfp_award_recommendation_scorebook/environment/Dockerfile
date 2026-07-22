FROM harbor-codex-base:2.1.114

COPY docops_v2_l4_068_word_xlsx_rfp_award_recommendation_scorebook_input.docx /root/docops_v2_l4_068_word_xlsx_rfp_award_recommendation_scorebook_input.docx
COPY raw_scoring_export.xlsx /root/raw_scoring_export.xlsx
COPY rfp_evaluation_rules.pdf /root/rfp_evaluation_rules.pdf
COPY vendor_proposal_excerpts.pdf /root/vendor_proposal_excerpts.pdf
COPY evaluator_conflict_note.docx /root/evaluator_conflict_note.docx
COPY original_task_description.txt /root/original_task_description.txt
COPY skills/ /skills/

COPY skills/ /root/.claude/skills/
COPY skills/ /root/.codex/skills/
COPY skills/ /root/.opencode/skill/
COPY skills/ /root/.goose/skills/
COPY skills/ /root/.factory/skills/
COPY skills/ /root/.agents/skills/
COPY skills/ /root/.gemini/skills/
