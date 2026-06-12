from __future__ import annotations

import argparse
import csv
import json
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable

import ezdxf


ROOT = Path(__file__).resolve().parents[2]
DEFAULT_OUTPUT_DIR = ROOT / "blender_automation" / "output" / "dxf"


@dataclass(frozen=True)
class Segment:
    source_type: str
    layer: str
    start: tuple[float, float, float]
    end: tuple[float, float, float]


def point3(point) -> tuple[float, float, float]:
    return (float(point[0]), float(point[1]), float(point[2] if len(point) > 2 else 0.0))


def default_dxf_path() -> Path:
    candidates = sorted(ROOT.glob("*.dxf"))
    if not candidates:
        raise FileNotFoundError(f"No DXF file found in {ROOT}")
    exact = ROOT / "xxx.dxf"
    return exact if exact.exists() else candidates[0]


def lwpolyline_segments(entity) -> Iterable[Segment]:
    elevation = entity.dxf.get("elevation", 0.0)
    z = float(elevation.z if hasattr(elevation, "z") else elevation)
    points = [(float(x), float(y), z) for x, y, *_ in entity.get_points()]
    if len(points) < 2:
        return
    pairs = zip(points, points[1:])
    for start, end in pairs:
        yield Segment(entity.dxftype(), entity.dxf.layer, start, end)
    if entity.closed:
        yield Segment(entity.dxftype(), entity.dxf.layer, points[-1], points[0])


def polyline_segments(entity) -> Iterable[Segment]:
    points = [point3(vertex.dxf.location) for vertex in entity.vertices]
    if len(points) < 2:
        return
    for start, end in zip(points, points[1:]):
        yield Segment(entity.dxftype(), entity.dxf.layer, start, end)
    if entity.is_closed:
        yield Segment(entity.dxftype(), entity.dxf.layer, points[-1], points[0])


def face_edges(entity) -> Iterable[Segment]:
    vertices = [point3(entity.dxf.vtx0), point3(entity.dxf.vtx1), point3(entity.dxf.vtx2), point3(entity.dxf.vtx3)]
    edges = [(vertices[0], vertices[1]), (vertices[1], vertices[2]), (vertices[2], vertices[3]), (vertices[3], vertices[0])]
    for start, end in edges:
        if start != end:
            yield Segment(entity.dxftype(), entity.dxf.layer, start, end)


def extract_segments(dxf_path: Path, include_3dface_edges: bool = False) -> list[Segment]:
    doc = ezdxf.readfile(dxf_path)
    segments: list[Segment] = []

    for entity in doc.modelspace():
        entity_type = entity.dxftype()
        if entity_type == "LINE":
            segments.append(Segment(entity_type, entity.dxf.layer, point3(entity.dxf.start), point3(entity.dxf.end)))
        elif entity_type == "LWPOLYLINE":
            segments.extend(lwpolyline_segments(entity))
        elif entity_type == "POLYLINE":
            segments.extend(polyline_segments(entity))
        elif include_3dface_edges and entity_type == "3DFACE":
            segments.extend(face_edges(entity))

    return segments


def write_json(path: Path, source: Path, segments: list[Segment]) -> None:
    payload = {
        "source": str(source),
        "segment_count": len(segments),
        "segments": [asdict(segment) for segment in segments],
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def write_csv(path: Path, segments: list[Segment]) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow(["source_type", "layer", "x1", "y1", "z1", "x2", "y2", "z2"])
        for segment in segments:
            writer.writerow([segment.source_type, segment.layer, *segment.start, *segment.end])


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract LINE/LWPOLYLINE/POLYLINE coordinates from a DXF file.")
    parser.add_argument("dxf", nargs="?", type=Path, default=None, help="DXF path. Defaults to xxx.dxf or the first DXF in the workspace.")
    parser.add_argument("--include-3dface-edges", action="store_true", help="Also export every 3DFACE boundary edge as a segment.")
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR)
    args = parser.parse_args()

    dxf_path = args.dxf or default_dxf_path()
    dxf_path = dxf_path.resolve()
    output_dir = args.output_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    segments = extract_segments(dxf_path, include_3dface_edges=args.include_3dface_edges)
    json_path = output_dir / "dxf_segments.json"
    csv_path = output_dir / "dxf_segments.csv"
    write_json(json_path, dxf_path, segments)
    write_csv(csv_path, segments)

    layers = sorted({segment.layer for segment in segments})
    print(f"DXF: {dxf_path}")
    print(f"Segments: {len(segments)}")
    print(f"Layers: {len(layers)}")
    print(f"JSON: {json_path}")
    print(f"CSV: {csv_path}")


if __name__ == "__main__":
    main()
