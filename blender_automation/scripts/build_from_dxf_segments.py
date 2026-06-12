from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import bpy
from mathutils import Vector


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_SEGMENTS_PATH = ROOT / "output" / "dxf" / "dxf_segments.json"
OUTPUT_DIR = ROOT / "output"
BLEND_PATH = OUTPUT_DIR / "dxf_line_model.blend"
RENDER_PATH = OUTPUT_DIR / "dxf_line_model.png"


def clear_scene() -> None:
    bpy.ops.object.select_all(action="SELECT")
    bpy.ops.object.delete()


def make_material(name: str, color: tuple[float, float, float, float]) -> bpy.types.Material:
    material = bpy.data.materials.new(name)
    material.use_nodes = True
    principled = next(node for node in material.node_tree.nodes if node.type == "BSDF_PRINCIPLED")
    principled.inputs["Base Color"].default_value = color
    principled.inputs["Roughness"].default_value = 0.45
    return material


def add_pipe_between(
    name: str,
    start: tuple[float, float, float],
    end: tuple[float, float, float],
    radius: float,
    material: bpy.types.Material,
) -> None:
    start_vec = Vector(start)
    end_vec = Vector(end)
    direction = end_vec - start_vec
    if direction.length == 0:
        return

    bpy.ops.mesh.primitive_cylinder_add(vertices=16, radius=radius, depth=direction.length, location=(start_vec + end_vec) / 2)
    obj = bpy.context.object
    obj.name = name
    obj.rotation_euler = direction.to_track_quat("Z", "Y").to_euler()
    obj.data.materials.append(material)
    bpy.ops.object.shade_smooth()


def load_segments(path: Path, limit: int | None) -> list[dict]:
    data = json.loads(path.read_text(encoding="utf-8"))
    segments = data["segments"]
    return segments[:limit] if limit else segments


def normalize_point(point: list[float], scale: float, origin: tuple[float, float, float]) -> tuple[float, float, float]:
    return (
        (float(point[0]) - origin[0]) * scale,
        (float(point[1]) - origin[1]) * scale,
        (float(point[2]) - origin[2]) * scale,
    )


def bounding_origin(segments: list[dict]) -> tuple[float, float, float]:
    xs = [coord for segment in segments for coord in (segment["start"][0], segment["end"][0])]
    ys = [coord for segment in segments for coord in (segment["start"][1], segment["end"][1])]
    zs = [coord for segment in segments for coord in (segment["start"][2], segment["end"][2])]
    return ((min(xs) + max(xs)) / 2, (min(ys) + max(ys)) / 2, min(zs))


def setup_camera() -> None:
    bpy.ops.object.light_add(type="AREA", location=(0, -12, 10))
    light = bpy.context.object
    light.data.energy = 800
    light.data.size = 7

    bpy.ops.object.camera_add(location=(9, -12, 8))
    camera = bpy.context.object
    bpy.context.scene.camera = camera
    direction = Vector((0, 0, 2.2)) - camera.location
    camera.rotation_euler = direction.to_track_quat("-Z", "Y").to_euler()
    camera.data.lens = 28


def configure_render() -> None:
    bpy.context.scene.render.engine = "CYCLES"
    bpy.context.scene.cycles.samples = 48
    bpy.context.scene.render.resolution_x = 1400
    bpy.context.scene.render.resolution_y = 1000


def main() -> None:
    parser = argparse.ArgumentParser(description="Build a simple Blender pipe model from extracted DXF line segments.")
    parser.add_argument("--segments", type=Path, default=DEFAULT_SEGMENTS_PATH)
    parser.add_argument("--scale", type=float, default=0.001, help="Coordinate scale. Use 0.001 for mm to m.")
    parser.add_argument("--radius", type=float, default=0.035, help="Pipe radius in Blender units after scaling.")
    parser.add_argument("--limit", type=int, default=2000, help="Maximum number of segments to model. Use 0 for all.")
    argv = sys.argv[sys.argv.index("--") + 1:] if "--" in sys.argv else []
    args = parser.parse_args(argv)

    segments = load_segments(args.segments.resolve(), None if args.limit == 0 else args.limit)
    origin = bounding_origin(segments)

    clear_scene()
    material = make_material("dxf_round_pipe", (0.78, 0.12, 0.08, 1.0))

    for index, segment in enumerate(segments, start=1):
        start = normalize_point(segment["start"], args.scale, origin)
        end = normalize_point(segment["end"], args.scale, origin)
        add_pipe_between(f"dxf_segment_{index}", start, end, args.radius, material)

    setup_camera()
    configure_render()
    OUTPUT_DIR.mkdir(exist_ok=True)
    bpy.ops.wm.save_as_mainfile(filepath=str(BLEND_PATH))
    bpy.context.scene.render.filepath = str(RENDER_PATH)
    bpy.ops.render.render(write_still=True)


if __name__ == "__main__":
    main()
