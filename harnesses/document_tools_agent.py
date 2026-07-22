#!/usr/bin/env python3
from __future__ import annotations

import json
import shlex
from pathlib import Path
from typing import Any, Literal

from pydantic import BaseModel, Field

from harbor.agents.base import BaseAgent
from harbor.environments.base import BaseEnvironment
from harbor.llms.lite_llm import LiteLLM
from harbor.models.agent.context import AgentContext

RUNTIME_TARGET = '/root/document_tools_runtime.py'


class AgentAction(BaseModel):
    thought: str
    action: Literal['call_tool', 'finish']
    tool_name: str | None = None
    arguments: dict[str, Any] = Field(default_factory=dict)
    final_message: str | None = None


class DocumentToolsAgent(BaseAgent):
    @staticmethod
    def name() -> str:
        return 'document-tools-agent'

    def __init__(
        self,
        logs_dir: Path,
        model_name: str | None = None,
        api_base: str | None = None,
        max_steps: int = 12,
        max_tokens: int = 32768,
        thinking_mode: bool | str | None = False,
        temperature: float = 0.0,
        reasoning_effort: str | None = None,
        **kwargs,
    ):
        super().__init__(logs_dir=logs_dir, model_name=model_name, **kwargs)
        self.api_base = api_base
        self.max_steps = max_steps
        self.max_tokens = max_tokens
        self.thinking_mode = thinking_mode
        if isinstance(self.thinking_mode, str):
            lowered = self.thinking_mode.strip().lower()
            if lowered in {'', 'none', 'null'}:
                self.thinking_mode = None
            elif lowered in {'1', 'true', 'yes'}:
                self.thinking_mode = True
            elif lowered in {'0', 'false', 'no'}:
                self.thinking_mode = False
        self.temperature = temperature
        self.reasoning_effort = reasoning_effort or None
        self.llm = LiteLLM(
            model_name=model_name or 'anthropic/claude-sonnet-4-6',
            api_base=api_base,
            temperature=temperature,
            reasoning_effort=self.reasoning_effort,
            model_info={
                'max_input_tokens': 262144,
                'max_output_tokens': 32768,
                'input_cost_per_token': 0,
                'output_cost_per_token': 0,
            },
        )
        self._usage = {'input': 0, 'output': 0, 'cache': 0, 'cost': 0.0}

    def version(self) -> str | None:
        return '0.1.1'

    async def setup(self, environment: BaseEnvironment) -> None:
        local_runtime = Path(__file__).with_name('document_tools_runtime.py')
        await environment.upload_file(local_runtime, RUNTIME_TARGET)
        await environment.exec(command=f'chmod +x {shlex.quote(RUNTIME_TARGET)}')
        await environment.exec(command='mkdir -p /root/submission')

    def _tool_specs(self) -> list[dict[str, Any]]:
        return [
            {'name': 'get_task_metadata', 'description': 'Args: {}. Read /root/task_metadata.json and return structured metadata including input_path and output_path.'},
            {'name': 'copy_file', 'description': 'Args: {src: string, dest: string}. Copy the source file to the required output path before mutating it.'},
            {'name': 'inspect_document', 'description': 'Args: {path: string, max_items?: integer}. Inspect a document and return a structured summary.'},
            {'name': 'word_replace_text', 'description': 'Args: {path: string, replacements: [{old: string, new: string}]}. Replace text in a DOCX paragraph/table.'},
            {'name': 'word_insert_table_column', 'description': 'Args: {path: string, table_index: integer, insert_after_column: integer, header_text?: string, fill_text?: string}. Insert a DOCX table column after a specified column index.'},
            {'name': 'xlsx_read_sheet', 'description': 'Args: {path: string, sheet_name: string, min_row?: integer, max_row?: integer, min_col?: integer, max_col?: integer}. Read a rectangular region from a workbook sheet.'},
            {'name': 'xlsx_write_cells', 'description': 'Args: {path: string, sheet_name: string, cells: [{cell: string, value: any}]}. Write literal values into workbook cells.'},
            {'name': 'xlsx_set_formula', 'description': 'Args: {path: string, sheet_name: string, cell: string, formula: string}. Set a formula in a workbook cell.'},
            {'name': 'xlsx_highlight_cells', 'description': 'Args: {path: string, sheet_name: string, cells: [string], fill_color?: string}. Apply a fill color to workbook cells.'},
            {'name': 'xlsx_delete_sheet', 'description': 'Args: {path: string, sheet_name: string}. Delete a sheet from a workbook.'},
            {'name': 'xlsx_reorder_sheets', 'description': 'Args: {path: string, sheet_order: [string]}. Reorder workbook sheets.'},
            {'name': 'ppt_list_slides', 'description': 'Args: {path: string}. List slide text content from a PPTX.'},
            {'name': 'ppt_replace_text', 'description': 'Args: {path: string, replacements: [{old: string, new: string}]}. Replace text across PPTX text runs.'},
            {'name': 'ppt_reorder_slides', 'description': 'Args: {path: string, order: [integer]}. Reorder slides using a 1-based list of slide indices.'},
            {'name': 'ppt_delete_slides', 'description': 'Args: {path: string, slide_indices: [integer]}. Delete slides using 1-based slide indices.'},
            {'name': 'ppt_set_bullets', 'description': 'Args: {path: string, slide_index: integer, bullets: [string], shape_index?: integer}. Replace the target text box content on a slide with bullet items.'},
            {'name': 'pdf_extract_text', 'description': 'Args: {path: string, page_numbers?: [integer], max_chars?: integer}. Extract plain text from PDF pages.'},
            {'name': 'pdf_reorder_pages', 'description': 'Args: {path: string, order: [integer]}. Reorder PDF pages in place using 1-based page indices.'},
            {'name': 'pdf_delete_pages', 'description': 'Args: {path: string, page_numbers: [integer]}. Delete PDF pages in place using 1-based page indices.'},
        ]

    def _system_prompt(self) -> str:
        tool_lines = '\n'.join(f"- {tool['name']}: {tool['description']}" for tool in self._tool_specs())
        return (
            'You are a structured document-editing agent for benchmarking. '
            'You must solve the task using only the provided tools, not shell commands.\n\n'
            'Important rules:\n'
            '1. Read the task metadata first.\n'
            '2. Never modify the source file in place; copy it to the required output path first.\n'
            '3. Use inspect tools to verify the result before finishing.\n'
            '4. Focus on the requested atomic operation only.\n'
            '5. Return exactly one JSON object matching the schema each turn.\n'
            '6. Do not wrap the action inside keys like "parameter" or "response".\n\n'
            '7. Do not include hidden reasoning, <think> tags, markdown fences, prose, or multiple candidate actions.\n\n'
            'Available tools:\n'
            f'{tool_lines}\n\n'
            'JSON schema:\n'
            '{"thought": str, "action": "call_tool"|"finish", "tool_name": str|null, "arguments": object, "final_message": str|null}'
        )

    def _parse_action(self, raw: str) -> AgentAction:
        try:
            data = json.loads(raw)
        except json.JSONDecodeError:
            decoder = json.JSONDecoder()
            data = None
            for index, char in enumerate(raw):
                if char != '{':
                    continue
                try:
                    data, _ = decoder.raw_decode(raw[index:])
                    break
                except json.JSONDecodeError:
                    continue
            if data is None:
                raise
        if isinstance(data, dict) and isinstance(data.get('parameter'), dict):
            data = data['parameter']
        if isinstance(data, dict) and isinstance(data.get('response'), dict):
            data = data['response']
        if isinstance(data, dict) and 'arguments' not in data:
            data['arguments'] = {}
        return AgentAction.model_validate(data)

    def _record_usage(self, response) -> None:
        if response.usage is None:
            return
        self._usage['input'] += response.usage.prompt_tokens or 0
        self._usage['output'] += response.usage.completion_tokens or 0
        self._usage['cache'] += response.usage.cache_tokens or 0
        self._usage['cost'] += response.usage.cost_usd or 0.0

    async def _exec_tool(self, environment: BaseEnvironment, tool_name: str, arguments: dict[str, Any]) -> dict[str, Any]:
        args_json = json.dumps(arguments, ensure_ascii=False)
        command = (
            f"python3 {shlex.quote(RUNTIME_TARGET)} "
            f"--tool {shlex.quote(tool_name)} "
            f"--args-json {shlex.quote(args_json)}"
        )
        result = await environment.exec(command=command)
        stdout = (result.stdout or '').strip()
        stderr = (result.stderr or '').strip()
        payload: dict[str, Any]
        try:
            payload = json.loads(stdout.splitlines()[-1]) if stdout else {'ok': False, 'error': 'empty stdout'}
        except Exception:
            payload = {'ok': False, 'error': 'failed to parse tool output', 'stdout': stdout, 'stderr': stderr}
        if stderr:
            payload.setdefault('stderr', stderr)
        return payload

    async def run(self, instruction: str, environment: BaseEnvironment, context: AgentContext) -> None:
        history: list[dict[str, str]] = [
            {'role': 'system', 'content': self._system_prompt()},
            {'role': 'user', 'content': instruction},
        ]

        bootstrap = await self._exec_tool(environment, 'get_task_metadata', {})
        history.append({'role': 'user', 'content': f'Bootstrap tool result (get_task_metadata):\n{json.dumps(bootstrap, ensure_ascii=False)}'})
        if bootstrap.get('ok') and isinstance(bootstrap.get('metadata'), dict):
            meta = bootstrap['metadata']
            history.append({'role': 'user', 'content': (
                'Important path reminder: '
                f"input_path={meta.get('input_path')} ; output_path={meta.get('output_path')} ; "
                'call inspect_document with a real path string, then copy_file from input_path to output_path before editing.'
            )})
            input_path = meta.get('input_path')
            if isinstance(input_path, str) and input_path:
                initial_inspect = await self._exec_tool(environment, 'inspect_document', {'path': input_path})
                history.append({'role': 'user', 'content': (
                    'Bootstrap tool result (inspect_document on input_path):\n'
                    f'{json.dumps(initial_inspect, ensure_ascii=False)}'
                )})

        transcript_path = self.logs_dir / 'document-tools-agent.txt'
        transcript_path.write_text('')

        final_message = 'Task ended without an explicit finish.'
        for step in range(1, self.max_steps + 1):
            response = await self.llm.call(
                prompt=(
                    'Choose the next action. Return only valid JSON matching this shape: '
                    '{"thought":"...","action":"call_tool|finish","tool_name":null|string,'
                    '"arguments":{},"final_message":null|string}.'
                ),
                message_history=history,
                logging_path=self.logs_dir / f'llm_call_{step}.json',
                extra_headers={'Accept-Encoding': 'identity'},
                extra_body={'thinking_mode': self.thinking_mode} if self.thinking_mode is not None else None,
                max_tokens=self.max_tokens,
                stop=['</think>', '</thinking>', '<action>', '```'],
            )
            self._record_usage(response)
            raw = response.content.strip()
            try:
                action = self._parse_action(raw)
            except Exception as exc:
                history.append({'role': 'assistant', 'content': raw})
                history.append({'role': 'user', 'content': f'Your previous response was invalid JSON for the schema: {exc}. Return corrected JSON only.'})
                transcript_path.write_text(transcript_path.read_text() + f'\nSTEP {step} INVALID\n{raw}\n')
                continue

            transcript_path.write_text(transcript_path.read_text() + f'\nSTEP {step}\n{action.model_dump_json(indent=2)}\n')

            if action.action == 'finish':
                final_message = action.final_message or action.thought
                break

            if not action.tool_name:
                history.append({'role': 'assistant', 'content': raw})
                history.append({'role': 'user', 'content': 'tool_name is required when action=call_tool. Return corrected JSON only.'})
                continue

            tool_result = await self._exec_tool(environment, action.tool_name, action.arguments)
            history.append({'role': 'assistant', 'content': raw})
            history.append({'role': 'user', 'content': f'Tool result for {action.tool_name}:\n{json.dumps(tool_result, ensure_ascii=False)}\nDecide the next action.'})

        transcript_path.write_text(transcript_path.read_text() + f'\nFINAL\n{final_message}\n')
        context.n_input_tokens = self._usage['input']
        context.n_output_tokens = self._usage['output']
        context.n_cache_tokens = self._usage['cache']
        context.cost_usd = self._usage['cost']
        context.metadata = {'max_steps': self.max_steps}
