$root = Split-Path $PSScriptRoot -Parent
$projectRoot = Join-Path $root 'SimpleCADforDSH'
$harness = Join-Path $root '..\deepseek-harness'

# Prefer the active Conda environment; no installation path or environment name
# is guessed. SIMPLECADFORDSH_PYTHON still takes precedence for explicit configuration.
if (-not $env:SIMPLECADFORDSH_PYTHON -and $env:CONDA_PREFIX) {
  $env:SIMPLECADFORDSH_PYTHON = Join-Path $env:CONDA_PREFIX 'python.exe'
}

$dshHome = if ($env:DSH_HOME) { $env:DSH_HOME } else { Join-Path $HOME '.dsh' }
$env:DSH_HOME = $dshHome

$modules = Join-Path $dshHome 'profiles\node_modules'
$link = Join-Path $modules 'simplecadfordsh'
New-Item -ItemType Directory -Force -Path $modules | Out-Null
if (-not (Test-Path $link)) {
  New-Item -ItemType Junction -Path $link -Target (Join-Path $projectRoot 'dsh-plugin') | Out-Null
}

$plugin = ([IO.Path]::GetFullPath((Join-Path $projectRoot 'dsh-plugin\index.js'))).Replace('\', '/')
$skills = ([IO.Path]::GetFullPath((Join-Path $root '.dsh\skills'))).Replace('\', '/')
$patch = Join-Path $projectRoot '.generated-cordis.patch.yml'
@"
- insert:
    - id: simplecadfordsh
      name: '$plugin'
- id: skill-filesystem
  disabled: false
  config:
    customSkillDirs:
      - '$skills'
"@ | Set-Content $patch -Encoding utf8

Push-Location $harness
pnpm dsh web --patch $patch
Pop-Location
