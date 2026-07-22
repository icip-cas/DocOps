# Docker Images

DocOps tasks use the Harbor base images below:

- `harbor-claude-code-base:2.1.114`
- `harbor-codex-base:2.1.114`

Docker images are identified by platform as `os/architecture`. Strictly
speaking, both bundled profiles are Linux containers for the x86_64/amd64 CPU
family. In this package, `x86` and `amd` are release profiles rather than two
incompatible CPU architectures.

This release includes two image profiles:

- `images/x86/`: server profile matching the images observed on `a800-1`
  (`x86_64`). Codex uses `harbor-codex080-chat-base:0.80.0-amd64` and is tagged
  as `harbor-codex-base:2.1.114` during loading.
- `images/amd/`: generic AMD64 profile with `harbor-codex-base:2.1.114`.

## x86 Server Profile

On an x86_64 Linux server, load the server profile before
running DocOps:

```bash
./docker/load_images.sh x86
```

The script loads the archives and tags them as the default image names expected
by the task Dockerfiles.

## AMD Profile

For a generic AMD64 Docker environment, load the AMD profile:

```bash
./docker/load_images.sh amd
```

The bundled images are still Linux containers for the amd64/x86_64 instruction
set; this profile simply uses the generic Codex base image.

## Codex and Self-Hosted vLLM

When running Codex against self-hosted models served through vLLM or another
OpenAI-compatible endpoint, we recommend the `x86` profile. This profile uses
the Codex 0.80 chat base image observed in the server runs, then tags it as
`harbor-codex-base:2.1.114` so the existing Harbor task Dockerfiles remain
unchanged.
