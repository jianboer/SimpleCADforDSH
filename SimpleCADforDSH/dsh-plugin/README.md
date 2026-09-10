# SimpleCADforDSH dsh plugin

Official DeepSeek Harness plugin contract: host tools + client slots + skill.
Does not fork or patch dsh source.

- Host: `simplecadfordsh_*` tools and `/simplecadfordsh` routes
- Client: `tool.call.toolview` cards and `shell.overlay` CAD pane
- Skill: `.dsh/skills/text-to-cad`

## Native dsh

1. Make package name `simplecadfordsh` resolvable from `$DSH_HOME/profiles` (add a dependency, or a link in `$DSH_HOME/profiles/node_modules/simplecadfordsh`).
2. Patch one row: `{ id: simplecadfordsh, name: simplecadfordsh }`.
3. Activate the Conda environment containing build123d. `start-dsh-web.ps1` uses its `CONDA_PREFIX`; native dsh without the script needs `SIMPLECADFORDSH_PYTHON` set explicitly (PowerShell: `$env:SIMPLECADFORDSH_PYTHON = (Get-Command python).Source`).

SimpleCADforDSH's start script does the local junction + `--patch` for this repo:

```powershell
# Run once from the DeepSeek Harness checkout:
pnpm install
pnpm run build

# Then activate Conda; the script detects CONDA_PREFIX.
powershell -File scripts\start-dsh-web.ps1
```

Open http://127.0.0.1:3080
