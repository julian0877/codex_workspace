# Blender Python portal steel frame

This project lets Codex edit Python geometry and lets Blender generate the model automatically.

## Files

- Blender: `E:\Program Files\Blender Foundation\Blender 4.4\blender.exe`
- Model script: `scripts\model_scene.py`
- Parameters: `config\portal_frame_params.json`
- Output: `output\codex_auto_model.blend` and `output\codex_auto_model.png`

## Run

```powershell
cd E:\codex_workspace\blender_automation
powershell -ExecutionPolicy Bypass -File .\run_blender.ps1
```

## Parameterization

Edit `config\portal_frame_params.json` to change:

- portal frame span, eave height, ridge height
- frame count and bay spacing
- whether roof purlins and wall girts are shown
- H-section sizes for columns and rafters
- H-section size and elevation for crane beams
- round pipe brace radius
- C-section sizes for optional purlins and wall girts

The current default model focuses on the primary structure: H-section columns, H-section rafters, longitudinal crane beams, roof horizontal round-pipe bracing, base plates, anchor bolts, and concrete pads. Set `show_secondary_members` to `true` if you want to display optional lipped C-section purlins and wall girts.
