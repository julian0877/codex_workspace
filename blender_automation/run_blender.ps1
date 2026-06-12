$ErrorActionPreference = "Stop"

$BlenderExe = "E:\Program Files\Blender Foundation\Blender 4.4\blender.exe"
$ScriptPath = Join-Path $PSScriptRoot "scripts\model_scene.py"

if (-not (Test-Path -LiteralPath $BlenderExe)) {
    throw "Blender executable not found: $BlenderExe"
}

if (-not (Test-Path -LiteralPath $ScriptPath)) {
    throw "Model script not found: $ScriptPath"
}

& $BlenderExe --background --python $ScriptPath

