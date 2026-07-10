import argparse
import os
import shutil
import subprocess
from pathlib import Path

from pdf2image import convert_from_path


LIBREOFFICE_PROGRAM = Path(r"C:\Program Files\LibreOffice\program")
SOFFICE = LIBREOFFICE_PROGRAM / "soffice.com"
POPPLER_BIN = Path(
    r"C:\Users\Administrator\.cache\codex-runtimes\codex-primary-runtime\dependencies\native\poppler\Library\bin"
)


def file_uri(path: Path) -> str:
    return "file:///" + path.resolve().as_posix().replace(":", "|", 1)


def render_docx(docx_path: Path, output_dir: Path, dpi: int) -> None:
    if not SOFFICE.exists():
        raise FileNotFoundError(f"LibreOffice command not found: {SOFFICE}")
    if not (POPPLER_BIN / "pdfinfo.exe").exists():
        raise FileNotFoundError(f"Poppler not found: {POPPLER_BIN}")

    output_dir.mkdir(parents=True, exist_ok=True)
    profile_dir = output_dir / "_lo_profile"
    pdf_dir = output_dir / "_pdf"
    if profile_dir.exists():
        shutil.rmtree(profile_dir)
    if pdf_dir.exists():
        shutil.rmtree(pdf_dir)
    profile_dir.mkdir(parents=True, exist_ok=True)
    pdf_dir.mkdir(parents=True, exist_ok=True)

    env = os.environ.copy()
    env["SAL_USE_VCLPLUGIN"] = "svp"
    env["PATH"] = f"{LIBREOFFICE_PROGRAM};{POPPLER_BIN};{env.get('PATH', '')}"

    cmd = [
        str(SOFFICE),
        f"-env:UserInstallation={file_uri(profile_dir)}",
        "--headless",
        "--nologo",
        "--nofirststartwizard",
        "--nolockcheck",
        "--nodefault",
        "--norestore",
        "--convert-to",
        "pdf",
        "--outdir",
        str(pdf_dir.resolve()),
        str(docx_path.resolve()),
    ]
    completed = subprocess.run(cmd, env=env, capture_output=True, text=True, timeout=120)
    if completed.returncode != 0:
        raise RuntimeError(
            "LibreOffice conversion failed\n"
            f"exit={completed.returncode}\nstdout={completed.stdout}\nstderr={completed.stderr}"
        )

    pdfs = list(pdf_dir.glob("*.pdf"))
    if not pdfs:
        raise RuntimeError(f"LibreOffice did not create a PDF in {pdf_dir}")
    pdf_path = pdfs[0]

    pages = convert_from_path(str(pdf_path), dpi=dpi, poppler_path=str(POPPLER_BIN))
    for index, page in enumerate(pages, start=1):
        page.save(output_dir / f"page-{index}.png")
    print(f"Rendered {len(pages)} page(s) to {output_dir}")


def main():
    parser = argparse.ArgumentParser(description="Render DOCX to PNG on this Windows workspace.")
    parser.add_argument("docx_path")
    parser.add_argument("--output_dir", required=True)
    parser.add_argument("--dpi", type=int, default=160)
    args = parser.parse_args()
    render_docx(Path(args.docx_path), Path(args.output_dir), args.dpi)


if __name__ == "__main__":
    main()
