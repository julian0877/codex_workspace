import json
from pathlib import Path


def write_report(
    path: Path,
    *,
    operation_count: int,
    warnings: list[str],
    paragraph_count: int,
    heading_count: int = 0,
    confirmed_count: int = 0,
    template_name: str = "",
) -> Path:
    payload = {
        "operation_count": operation_count,
        "paragraph_count": paragraph_count,
        "warnings": warnings,
        "heading_count": heading_count,
        "confirmed_count": confirmed_count,
        "template_name": template_name,
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path = path.with_suffix(path.suffix + ".tmp")
    temporary_path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    temporary_path.replace(path)
    return path
