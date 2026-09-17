"""add_slide.py가 필요로 하는 최소 zip 유틸리티.

원래 Anthropic pptx 스킬(office.helpers)에서 가져오던 rezip/safe_extract를
외부 스킬 의존성 없이 이 프로젝트 안에서 자체 구현한 것.
"""
from __future__ import annotations

import zipfile
from pathlib import Path


def safe_extract(zf: zipfile.ZipFile, dest: Path) -> None:
    """zip slip(경로 조작)을 방지하며 압축을 해제."""
    dest = Path(dest).resolve()
    for member in zf.infolist():
        target = (dest / member.filename).resolve()
        if target != dest and dest not in target.parents:
            raise ValueError(f"Unsafe path in zip entry: {member.filename}")
    zf.extractall(dest)


def rezip(src_dir: Path, out_path: Path) -> None:
    """디렉터리를 다시 pptx(zip)로 묶는다. 원자적 교체를 위해 임시 파일에 먼저 쓴다."""
    src_dir = Path(src_dir)
    out_path = Path(out_path)
    tmp_path = out_path.with_name(out_path.name + ".tmp")
    with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for file in sorted(src_dir.rglob("*")):
            if file.is_file():
                zf.write(file, file.relative_to(src_dir).as_posix())
    tmp_path.replace(out_path)
