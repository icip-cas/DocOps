# DocOps

[Project Page](https://jiangxiaohuan3.github.io/docops.github.io/) | [License](LICENSE)

DocOps is a Harbor-formatted benchmark for evaluating document-operation
agents across Excel, Word, PowerPoint, and PDF files. Unlike document QA
benchmarks that mainly inspect textual answers, DocOps evaluates whether an
agent can edit native documents into the requested state while preserving
format validity, document structure, formulas, styles, outlines, bookmarks, and
other task-relevant out-of-scope state.

This repository contains the complete 210-task release, deterministic
artifact-level verifiers, execution harness wrappers, document-operation
skills, and Docker base images needed to reproduce DocOps benchmark runs.

<p align="center">
  <img src="assets/main_figure.png" alt="DocOps benchmark overview" width="95%">
</p>

## ✨ Highlights

- **210 executable Harbor tasks** covering native document operations.
- **Four document formats**: Excel, Word, PowerPoint, and PDF.
- **Four difficulty levels**:
  - L1: localized atomic operations, 50 tasks.
  - L2: compositional same-document operations, 40 tasks.
  - L3: single-document workflows, 60 tasks.
  - L4: cross-document workflows, 60 tasks.
- **Deterministic artifact verification** through task-level tests that inspect
  native file state rather than only visible text.
- **Multiple execution harnesses** for DocumentTools, Terminus-2, Codex, and
  Claude Code.
- **Skill-on and skill-off evaluation** for Codex and Claude Code.
- **Bundled Docker base images** for the supported Linux amd64/x86_64 runtime
  profiles.

## 📁 Repository Layout

```text
.
|-- assets/                # README and project-page images
|-- docker/                # Docker image archives and image-loading script
|-- harnesses/             # Harness-specific runtime wrappers
|-- scripts/               # Setup, validation, materialization, and run scripts
|-- skills/                # Shared document-operation skills
|-- tasks/                 # Complete 210 Harbor task directories
`-- third_party/harbor/    # Vendored Harbor runner source
```

Generated local files are written to `results/` and `runtime/`; both are
ignored by git.

Each task follows the Harbor task format:

```text
task_name/
|-- instruction.md
|-- task.toml
|-- environment/
|   |-- Dockerfile
|   |-- task_metadata.json
|   `-- input documents
`-- tests/
    |-- test.sh
    |-- test_outputs.py
    `-- verifier utilities
```

## 🚀 Quick Start

### 1. Install Python Dependencies

Python 3.12 or newer is recommended.

```bash
python3.12 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

You can also use the helper script:

```bash
./scripts/setup_env.sh
```

Verify that the packaged benchmark is complete:

```bash
python scripts/verify_package.py
```

The verifier should report 210 tasks with the expected L1-L4 distribution.

### 2. Load Docker Base Images

DocOps task Dockerfiles depend on these Harbor base image tags:

```text
harbor-claude-code-base:2.1.114
harbor-codex-base:2.1.114
```

Two Docker image profiles are provided:

```text
docker/images/x86/    # Linux x86_64 server profile; recommended for vLLM + Codex
docker/images/amd/    # Generic Linux amd64 profile
```

Use the server x86_64 profile:

```bash
./docker/load_images.sh x86
```

Use the generic amd64 profile:

```bash
./docker/load_images.sh amd
```

Both profiles contain Linux containers for the amd64/x86_64 CPU family. They
are separated because the benchmark was packaged with two runtime image
profiles. ARM64 machines require a separately exported image set.

For Codex runs against self-hosted models served through vLLM or another
OpenAI-compatible endpoint, the `x86` profile is recommended. It includes the
Codex 0.80 chat base image used in our server-side runs and tags it as
`harbor-codex-base:2.1.114`, so the task Dockerfiles do not need to be edited.

### 3. Configure API Credentials

Create a local environment file:

```bash
cp .env.example .env
```

Fill in only the fields required by the harness you plan to run.

For OpenAI-compatible endpoints used by DocumentTools, Terminus-2, and Codex:

```bash
OPENAI_API_KEY=your_key
OPENAI_BASE_URL=http://host:port/v1
DOCOPS_DOCTOOLS_MODEL=openai/served-model-name
DOCOPS_TERMINUS2_MODEL=openai/served-model-name
DOCOPS_CODEX_MODEL=served-model-name
```

For Claude Code:

```bash
ANTHROPIC_BASE_URL=https://your-anthropic-compatible-endpoint
ANTHROPIC_AUTH_TOKEN=your_key
DOCOPS_CLAUDE_MODEL=claude-sonnet-4-6
```

Credentials and run outputs are ignored by git.

## 🧠 Self-Hosted vLLM

For open-weight model evaluation, serve the model through an OpenAI-compatible
vLLM endpoint:

```bash
vllm serve /path/to/model \
  --host 0.0.0.0 \
  --port 8000 \
  --served-model-name served-model-name \
  --tensor-parallel-size 8 \
  --enable-prefix-caching \
  --enable-chunked-prefill \
  --gpu-memory-utilization 0.90
```

Then point DocOps to the endpoint in `.env`:

```bash
OPENAI_API_KEY=dummy
OPENAI_BASE_URL=http://server:8000/v1
DOCOPS_CODEX_MODEL=served-model-name
```

For Codex with vLLM-hosted models, use Codex 0.80 and load the `x86` Docker
profile:

```bash
./docker/load_images.sh x86
```

## 🧪 Running DocOps

Run any harness wrapper with the default model configured in `.env`:

```bash
./scripts/run_doctools.sh
./scripts/run_terminus2.sh
./scripts/run_codex_with_skill.sh
./scripts/run_codex_no_skill.sh
./scripts/run_claude_code_with_skill.sh
./scripts/run_claude_code_no_skill.sh
```

All wrapper scripts call the unified runner:

```bash
./scripts/run_harness.sh doctools
./scripts/run_harness.sh terminus2
./scripts/run_harness.sh codex --skill on
./scripts/run_harness.sh codex --skill off
./scripts/run_harness.sh claude-code --skill on
./scripts/run_harness.sh claude-code --skill off
```

You can override the model or output directory per run:

```bash
./scripts/run_harness.sh codex --skill on \
  --model served-model-name \
  --output results/codex_served_model_with_skill
```

To run every packaged harness script sequentially:

```bash
./scripts/run_all_harnesses.sh
```

## 🧩 Skills and No-Skill Runs

Skill-enabled runs use the task directories in `tasks/` directly. These tasks
copy `environment/skills/` into common agent skill locations inside the
container.

No-skill runs are materialized under:

```text
runtime/no_skill_tasks/
```

The materialization step removes skill references from `task.toml`,
`instruction.md`, task Dockerfiles, and per-task `environment/skills/`
directories. This keeps skill-on and skill-off settings comparable while
preserving the same task inputs and verifiers.

## 📊 Outputs

Runs are written under `results/` by default. Harbor stores task-level logs,
agent outputs, verifier outputs, and result metadata in the selected output
directory.

The default run controls are configured in `.env`:

```bash
HARBOR_RUN_COUNT=1
HARBOR_N_CONCURRENT=1
DOCOPS_OUTPUT_ROOT=results
```

Increase `HARBOR_N_CONCURRENT` only after confirming that the selected model
endpoint and Docker host can handle the workload.

## ✅ Package Check

Use the validation script after cloning, moving, or editing the package:

```bash
python scripts/verify_package.py
```

It checks the expected task count, task-source split, difficulty distribution,
and required release files.

## 📄 License

This repository is released under the Apache-2.0 license. Third-party
components and bundled skills retain their own license notices where included.
