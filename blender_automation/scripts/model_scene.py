from __future__ import annotations

import json
from dataclasses import dataclass
from math import atan2, cos, radians, sin, sqrt
from pathlib import Path

import bpy
from mathutils import Vector


ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = ROOT / "output"
BLEND_PATH = OUTPUT_DIR / "codex_auto_model.blend"
RENDER_PATH = OUTPUT_DIR / "codex_auto_model.png"
CONFIG_PATH = ROOT / "config" / "portal_frame_params.json"


@dataclass(frozen=True)
class HSection:
    depth: float
    flange_width: float
    flange_thickness: float
    web_thickness: float


@dataclass(frozen=True)
class CSection:
    depth: float
    flange_width: float
    thickness: float
    lip: float


@dataclass(frozen=True)
class PortalFrameParams:
    span: float = 8.0
    eave_height: float = 4.1
    ridge_height: float = 5.35
    frame_count: int = 3
    bay_spacing: float = 4.0
    show_secondary_members: bool = False
    roof_purlin_count: int = 7
    wall_girt_count: int = 3
    crane_beam_height: float = 2.85
    pipe_brace_radius: float = 0.055
    column_section: HSection = HSection(0.50, 0.42, 0.075, 0.07)
    rafter_section: HSection = HSection(0.42, 0.38, 0.065, 0.06)
    crane_beam_section: HSection = HSection(0.34, 0.24, 0.055, 0.045)
    purlin_section: CSection = CSection(0.22, 0.09, 0.025, 0.045)
    girt_section: CSection = CSection(0.18, 0.075, 0.022, 0.035)


def load_params() -> PortalFrameParams:
    if not CONFIG_PATH.exists():
        return PortalFrameParams()

    data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    return PortalFrameParams(
        span=float(data.get("span", PortalFrameParams.span)),
        eave_height=float(data.get("eave_height", PortalFrameParams.eave_height)),
        ridge_height=float(data.get("ridge_height", PortalFrameParams.ridge_height)),
        frame_count=int(data.get("frame_count", PortalFrameParams.frame_count)),
        bay_spacing=float(data.get("bay_spacing", PortalFrameParams.bay_spacing)),
        show_secondary_members=bool(data.get("show_secondary_members", PortalFrameParams.show_secondary_members)),
        roof_purlin_count=int(data.get("roof_purlin_count", PortalFrameParams.roof_purlin_count)),
        wall_girt_count=int(data.get("wall_girt_count", PortalFrameParams.wall_girt_count)),
        crane_beam_height=float(data.get("crane_beam_height", PortalFrameParams.crane_beam_height)),
        pipe_brace_radius=float(data.get("pipe_brace_radius", PortalFrameParams.pipe_brace_radius)),
        column_section=HSection(**data.get("column_section", {})) if "column_section" in data else PortalFrameParams.column_section,
        rafter_section=HSection(**data.get("rafter_section", {})) if "rafter_section" in data else PortalFrameParams.rafter_section,
        crane_beam_section=HSection(**data.get("crane_beam_section", {})) if "crane_beam_section" in data else PortalFrameParams.crane_beam_section,
        purlin_section=CSection(**data.get("purlin_section", {})) if "purlin_section" in data else PortalFrameParams.purlin_section,
        girt_section=CSection(**data.get("girt_section", {})) if "girt_section" in data else PortalFrameParams.girt_section,
    )


def clear_scene() -> None:
    bpy.ops.object.select_all(action="SELECT")
    bpy.ops.object.delete()


def make_material(name: str, color: tuple[float, float, float, float], roughness: float = 0.45) -> bpy.types.Material:
    material = bpy.data.materials.new(name)
    material.use_nodes = True
    principled = next(
        node for node in material.node_tree.nodes
        if node.type == "BSDF_PRINCIPLED"
    )
    principled.inputs["Base Color"].default_value = color
    principled.inputs["Roughness"].default_value = roughness
    return material


def add_cube(
    name: str,
    location: tuple[float, float, float],
    scale: tuple[float, float, float],
    material: bpy.types.Material,
) -> bpy.types.Object:
    bpy.ops.mesh.primitive_cube_add(size=1, location=location)
    obj = bpy.context.object
    obj.name = name
    obj.dimensions = scale
    bpy.ops.object.transform_apply(location=False, rotation=False, scale=True)
    obj.data.materials.append(material)
    return obj


def add_cylinder(
    name: str,
    location: tuple[float, float, float],
    radius: float,
    depth: float,
    material: bpy.types.Material,
    vertices: int = 64,
) -> bpy.types.Object:
    bpy.ops.mesh.primitive_cylinder_add(vertices=vertices, radius=radius, depth=depth, location=location)
    obj = bpy.context.object
    obj.name = name
    obj.data.materials.append(material)
    return obj


def shade_smooth(objects: list[bpy.types.Object]) -> None:
    for obj in objects:
        bpy.context.view_layer.objects.active = obj
        obj.select_set(True)
        bpy.ops.object.shade_smooth()
        obj.select_set(False)


def rotated_location(
    center: tuple[float, float, float],
    offset: tuple[float, float, float],
    angle_y: float,
) -> tuple[float, float, float]:
    ox, oy, oz = offset
    return (
        center[0] + ox * cos(angle_y) + oz * sin(angle_y),
        center[1] + oy,
        center[2] - ox * sin(angle_y) + oz * cos(angle_y),
    )


def add_rotated_cube(
    name: str,
    location: tuple[float, float, float],
    scale: tuple[float, float, float],
    material: bpy.types.Material,
    rotation_y: float = 0,
) -> bpy.types.Object:
    obj = add_cube(name, location, scale, material)
    obj.rotation_euler[1] = rotation_y
    return obj


def add_i_beam_between(
    name: str,
    start: tuple[float, float, float],
    end: tuple[float, float, float],
    section: HSection,
    material: bpy.types.Material,
) -> list[bpy.types.Object]:
    dx = end[0] - start[0]
    dz = end[2] - start[2]
    length = sqrt(dx * dx + dz * dz)
    angle_y = -atan2(dz, dx)
    center = ((start[0] + end[0]) / 2, start[1], (start[2] + end[2]) / 2)
    flange_offset = section.depth / 2 - section.flange_thickness / 2

    parts = [
        ("web", (length, section.web_thickness, section.depth - section.flange_thickness * 2), (0, 0, 0)),
        ("top_flange", (length, section.flange_width, section.flange_thickness), (0, 0, flange_offset)),
        ("bottom_flange", (length, section.flange_width, section.flange_thickness), (0, 0, -flange_offset)),
    ]

    objects: list[bpy.types.Object] = []
    for suffix, scale, offset in parts:
        objects.append(
            add_rotated_cube(
                f"{name}_{suffix}",
                rotated_location(center, offset, angle_y),
                scale,
                material,
                angle_y,
            )
        )
    return objects


def add_i_column(
    name: str,
    x: float,
    y: float,
    height: float,
    section: HSection,
    material: bpy.types.Material,
) -> list[bpy.types.Object]:
    z = height / 2
    flange_offset = section.depth / 2 - section.flange_thickness / 2
    return [
        add_cube(
            f"{name}_web",
            (x, y, z),
            (section.depth - section.flange_thickness * 2, section.web_thickness, height),
            material,
        ),
        add_cube(
            f"{name}_outer_flange",
            (x - flange_offset, y, z),
            (section.flange_thickness, section.flange_width, height),
            material,
        ),
        add_cube(
            f"{name}_inner_flange",
            (x + flange_offset, y, z),
            (section.flange_thickness, section.flange_width, height),
            material,
        ),
    ]


def add_i_beam_y(
    name: str,
    x: float,
    y_center: float,
    z: float,
    length: float,
    section: HSection,
    material: bpy.types.Material,
) -> list[bpy.types.Object]:
    flange_offset = section.depth / 2 - section.flange_thickness / 2
    return [
        add_cube(
            f"{name}_web",
            (x, y_center, z),
            (section.web_thickness, length, section.depth - section.flange_thickness * 2),
            material,
        ),
        add_cube(
            f"{name}_top_flange",
            (x, y_center, z + flange_offset),
            (section.flange_width, length, section.flange_thickness),
            material,
        ),
        add_cube(
            f"{name}_bottom_flange",
            (x, y_center, z - flange_offset),
            (section.flange_width, length, section.flange_thickness),
            material,
        ),
    ]


def add_pipe_between(
    name: str,
    start: tuple[float, float, float],
    end: tuple[float, float, float],
    radius: float,
    material: bpy.types.Material,
) -> bpy.types.Object:
    start_vec = Vector(start)
    end_vec = Vector(end)
    center = (start_vec + end_vec) / 2
    direction = end_vec - start_vec

    bpy.ops.mesh.primitive_cylinder_add(vertices=32, radius=radius, depth=direction.length, location=center)
    obj = bpy.context.object
    obj.name = name
    obj.rotation_euler = direction.to_track_quat("Z", "Y").to_euler()
    obj.data.materials.append(material)
    bpy.ops.object.shade_smooth()
    return obj


def add_c_channel_y(
    name: str,
    center: tuple[float, float, float],
    length: float,
    section: CSection,
    material: bpy.types.Material,
    rotation_y: float = 0,
) -> list[bpy.types.Object]:
    web_x = -section.flange_width / 2 + section.thickness / 2
    free_edge_x = section.flange_width / 2 - section.thickness / 2
    flange_x = section.thickness / 2
    flange_z = section.depth / 2 - section.thickness / 2
    lip_z = section.depth / 2 - section.lip / 2

    parts = [
        ("web", (section.thickness, length, section.depth), (web_x, 0, 0)),
        ("top_flange", (section.flange_width, length, section.thickness), (flange_x, 0, flange_z)),
        ("bottom_flange", (section.flange_width, length, section.thickness), (flange_x, 0, -flange_z)),
        ("top_lip", (section.thickness, length, section.lip), (free_edge_x, 0, lip_z)),
        ("bottom_lip", (section.thickness, length, section.lip), (free_edge_x, 0, -lip_z)),
    ]

    objects: list[bpy.types.Object] = []
    for suffix, scale, offset in parts:
        objects.append(
            add_rotated_cube(
                f"{name}_{suffix}",
                rotated_location(center, offset, rotation_y),
                scale,
                material,
                rotation_y,
            )
        )
    return objects


def frame_positions(params: PortalFrameParams) -> list[float]:
    start = -(params.frame_count - 1) * params.bay_spacing / 2
    return [start + index * params.bay_spacing for index in range(params.frame_count)]


def add_bolt_group(
    name: str,
    x: float,
    y: float,
    material: bpy.types.Material,
) -> None:
    for index, (ox, oy) in enumerate([(-0.28, -0.18), (-0.28, 0.18), (0.28, -0.18), (0.28, 0.18)], start=1):
        bolt = add_cylinder(f"{name}_anchor_bolt_{index}", (x + ox, y + oy, 0.18), 0.045, 0.12, material, 24)
        bolt.rotation_euler[0] = radians(0)


def build_portal_steel_frame(params: PortalFrameParams) -> None:
    steel = make_material("galvanized_blue_gray_steel", (0.42, 0.49, 0.55, 1.0), 0.32)
    dark_steel = make_material("dark_connection_plates", (0.12, 0.14, 0.16, 1.0), 0.35)
    concrete = make_material("pale_concrete", (0.62, 0.61, 0.57, 1.0), 0.65)
    floor_mat = make_material("matte_site_floor", (0.42, 0.47, 0.43, 1.0), 0.7)
    safety_red = make_material("red_bracing_rods", (0.82, 0.08, 0.05, 1.0), 0.42)

    y_positions = frame_positions(params)
    building_length = (params.frame_count - 1) * params.bay_spacing
    purlin_length = building_length + 0.9
    left_x = -params.span / 2
    right_x = params.span / 2
    crane_left_x = left_x + params.column_section.depth / 2 + params.crane_beam_section.flange_width / 2
    crane_right_x = right_x - params.column_section.depth / 2 - params.crane_beam_section.flange_width / 2
    roof_rise = params.ridge_height - params.eave_height
    half_span = params.span / 2
    left_roof_angle = -atan2(roof_rise, half_span)
    right_roof_angle = atan2(roof_rise, half_span)

    add_cube("site_floor", (0, 0, -0.06), (params.span + 2.8, building_length + 3.4, 0.12), floor_mat)

    for frame_index, y in enumerate(y_positions, start=1):
        add_cube(f"concrete_pad_left_y{y}", (left_x, y, 0.12), (1.15, 0.9, 0.24), concrete)
        add_cube(f"concrete_pad_right_y{y}", (right_x, y, 0.12), (1.15, 0.9, 0.24), concrete)
        add_cube(f"base_plate_left_y{y}", (left_x, y, 0.31), (0.82, 0.58, 0.08), dark_steel)
        add_cube(f"base_plate_right_y{y}", (right_x, y, 0.31), (0.82, 0.58, 0.08), dark_steel)
        add_bolt_group(f"left_y{y}", left_x, y, dark_steel)
        add_bolt_group(f"right_y{y}", right_x, y, dark_steel)

        add_i_column(f"left_column_frame_{frame_index}", left_x, y, params.eave_height, params.column_section, steel)
        add_i_column(f"right_column_frame_{frame_index}", right_x, y, params.eave_height, params.column_section, steel)

        add_i_beam_between(
            f"left_roof_rafter_frame_{frame_index}",
            (left_x, y, params.eave_height),
            (0, y, params.ridge_height),
            params.rafter_section,
            steel,
        )
        add_i_beam_between(
            f"right_roof_rafter_frame_{frame_index}",
            (0, y, params.ridge_height),
            (right_x, y, params.eave_height),
            params.rafter_section,
            steel,
        )

        add_cube(f"ridge_connection_plate_frame_{frame_index}", (0, y, params.ridge_height), (0.44, 0.58, 0.52), dark_steel)
        add_cube(f"left_knee_plate_frame_{frame_index}", (left_x, y, params.eave_height), (0.58, 0.56, 0.5), dark_steel)
        add_cube(f"right_knee_plate_frame_{frame_index}", (right_x, y, params.eave_height), (0.58, 0.56, 0.5), dark_steel)

        add_rotated_cube(f"left_knee_haunch_frame_{frame_index}", (left_x + 0.45, y, params.eave_height + 0.28), (1.05, 0.42, 0.18), steel, radians(-18))
        add_rotated_cube(f"right_knee_haunch_frame_{frame_index}", (right_x - 0.45, y, params.eave_height + 0.28), (1.05, 0.42, 0.18), steel, radians(18))

    add_i_beam_y(
        "left_crane_beam",
        crane_left_x,
        0,
        params.crane_beam_height,
        building_length + 0.35,
        params.crane_beam_section,
        dark_steel,
    )
    add_i_beam_y(
        "right_crane_beam",
        crane_right_x,
        0,
        params.crane_beam_height,
        building_length + 0.35,
        params.crane_beam_section,
        dark_steel,
    )

    for bay_index, (y0, y1) in enumerate(zip(y_positions, y_positions[1:]), start=1):
        low_left = (left_x + 0.72, y0, params.eave_height + 0.18)
        low_right = (right_x - 0.72, y0, params.eave_height + 0.18)
        high_left = (-0.34, y1, params.ridge_height - 0.18)
        high_right = (0.34, y1, params.ridge_height - 0.18)

        add_pipe_between(f"roof_horizontal_pipe_brace_left_a_bay_{bay_index}", low_left, high_left, params.pipe_brace_radius, safety_red)
        add_pipe_between(f"roof_horizontal_pipe_brace_left_b_bay_{bay_index}", (left_x + 0.72, y1, params.eave_height + 0.18), (-0.34, y0, params.ridge_height - 0.18), params.pipe_brace_radius, safety_red)
        add_pipe_between(f"roof_horizontal_pipe_brace_right_a_bay_{bay_index}", high_right, low_right, params.pipe_brace_radius, safety_red)
        add_pipe_between(f"roof_horizontal_pipe_brace_right_b_bay_{bay_index}", (0.34, y0, params.ridge_height - 0.18), (right_x - 0.72, y1, params.eave_height + 0.18), params.pipe_brace_radius, safety_red)

    if params.show_secondary_members:
        purlin_step = params.span / (params.roof_purlin_count + 1)
        roof_x_positions = [left_x + purlin_step * index for index in range(1, params.roof_purlin_count + 1)]
        for index, x in enumerate(roof_x_positions, start=1):
            z = params.ridge_height - abs(x) / half_span * roof_rise
            rotation_y = left_roof_angle if x < 0 else right_roof_angle
            add_c_channel_y(f"roof_c_purlin_{index}", (x, 0, z + 0.22), purlin_length, params.purlin_section, dark_steel, rotation_y)

        for index in range(1, params.wall_girt_count + 1):
            z = params.eave_height * index / (params.wall_girt_count + 1)
            add_c_channel_y(f"left_wall_c_girt_{index}", (left_x - 0.22, 0, z), purlin_length, params.girt_section, dark_steel)
            add_c_channel_y(f"right_wall_c_girt_{index}", (right_x + 0.22, 0, z), purlin_length, params.girt_section, dark_steel, radians(180))

    add_cube("scale_reference_person", (left_x - 1.05, y_positions[0] - 0.2, 0.85), (0.22, 0.22, 1.7), dark_steel)


def setup_camera_and_lights() -> None:
    bpy.ops.object.light_add(type="AREA", location=(0, -5.5, 8.0))
    key = bpy.context.object
    key.name = "large_softbox_key"
    key.data.energy = 780
    key.data.size = 6.0

    bpy.ops.object.light_add(type="POINT", location=(-5.5, 4.0, 5.2))
    rim = bpy.context.object
    rim.name = "small_rim_light"
    rim.data.energy = 160

    bpy.ops.object.camera_add(location=(9.6, -11.5, 7.4))
    camera = bpy.context.object
    bpy.context.scene.camera = camera
    direction = Vector((0, 0, 2.7)) - camera.location
    camera.rotation_euler = direction.to_track_quat("-Z", "Y").to_euler()
    camera.data.lens = 28


def configure_render() -> None:
    bpy.context.scene.render.engine = "CYCLES"
    bpy.context.scene.cycles.samples = 64
    bpy.context.scene.render.resolution_x = 1400
    bpy.context.scene.render.resolution_y = 1000
    bpy.context.scene.view_settings.view_transform = "Filmic"
    bpy.context.scene.view_settings.look = "Medium High Contrast"


def main() -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)
    clear_scene()
    params = load_params()
    build_portal_steel_frame(params)
    setup_camera_and_lights()
    configure_render()

    bpy.ops.wm.save_as_mainfile(filepath=str(BLEND_PATH))
    bpy.context.scene.render.filepath = str(RENDER_PATH)
    bpy.ops.render.render(write_still=True)


if __name__ == "__main__":
    main()
