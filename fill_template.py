#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
템플릿(Template.pptx)에 학생 이름+내용 데이터를 자동으로 채워 넣는 스크립트.
- 슬라이드 1장당 12개(표 4개 x 3칸)까지 채움
- 12개 초과 시 슬라이드를 자동으로 복제하여 다음 페이지에 이어서 채움
- 12개 미만이면 남은 칸은 빈칸으로 둠 ({{이름}}/{{내용}} 마커만 제거)
"""
import os
import re
import sys
import shutil
import subprocess
from pathlib import Path
from pptx import Presentation

PER_SLIDE = 12  # 표 4개 x 3칸

# 슬라이드 안에서 데이터를 채울 순서: (표 이름, 시작 열번호)
SLOT_ORDER = [
    ("표 1", 0), ("표 1", 7), ("표 1", 14),
    ("표 6", 0), ("표 6", 7), ("표 6", 14),
    ("표 8", 0), ("표 8", 7), ("표 8", 14),
    ("표 10", 0), ("표 10", 7), ("표 10", 14),
]

NAME_ROW = 0
CONTENT_ROW = 3


def parse_records(raw_text):
    """'■ 이름 내용...' 형태의 레코드가 여러 개 붙은 텍스트를 리스트로 분리.

    API는 Make.com에서 이미 구조화된 JSON 배열을 받으므로 보통 이 함수는 쓰이지 않는다.
    텍스트 블록을 그대로 보내는 경우를 대비한 폴백 파서로만 남겨둠.
    """
    chunks = re.split(r"(?=■)", raw_text.strip())
    records = []
    for chunk in chunks:
        chunk = chunk.strip()
        if not chunk:
            continue
        chunk = chunk.lstrip("■").strip()
        m = re.match(r"^(\S+)\s+(.*)$", chunk, re.DOTALL)
        if not m:
            continue
        name, content = m.group(1), m.group(2).strip()
        records.append({"name": name, "content": content})
    return records


def set_cell_text(cell, new_text):
    """플레이스홀더가 여러 run으로 쪼개져 있어도 첫 run의 서식을 유지하며 텍스트 교체."""
    tf = cell.text_frame
    para = tf.paragraphs[0]
    runs = para.runs
    if not runs:
        para.text = new_text
        return
    runs[0].text = new_text
    for r in runs[1:]:
        r.text = ""
    # 혹시 플레이스홀더 아래 추가 문단이 있으면 텍스트 비우기 (서식은 유지)
    for extra_p in tf.paragraphs[1:]:
        for r in extra_p.runs:
            r.text = ""


def fill_slide(slide, records_chunk):
    """records_chunk: 최대 12개의 dict({'name':..,'content':..}) 또는 None(빈칸)"""
    tables = {}
    for shape in slide.shapes:
        if shape.has_table:
            tables[shape.name] = shape.table

    padded = list(records_chunk) + [None] * (PER_SLIDE - len(records_chunk))

    for (table_name, col_start), record in zip(SLOT_ORDER, padded):
        table = tables[table_name]
        name_cell = table.cell(NAME_ROW, col_start)
        content_cell = table.cell(CONTENT_ROW, col_start)
        if record is None:
            set_cell_text(name_cell, "")
            set_cell_text(content_cell, "")
        else:
            set_cell_text(name_cell, record["name"])
            set_cell_text(content_cell, record["content"])


def build_pptx(template_path, records, output_path, scripts_dir=None, work_dir=None):
    """records가 비어 있으면 빈칸투성이 1장을 생성한다.

    work_dir: 슬라이드 복제 작업용 임시 디렉터리. 동시 요청이 서로의 임시 파일을
    덮어쓰지 않도록 호출자(API 서버)가 요청마다 고유한 디렉터리를 넘겨야 한다.
    """
    scripts_dir = Path(scripts_dir) if scripts_dir else Path(__file__).parent / "pptx_scripts"
    work_dir = Path(work_dir) if work_dir else Path(output_path).parent
    work_dir.mkdir(parents=True, exist_ok=True)

    n_slides_needed = max(1, -(-len(records) // PER_SLIDE))  # ceil division

    # 1) 필요한 만큼 슬라이드 복제 (템플릿은 slide1.xml 한 장짜리)
    work_path = work_dir / "_work.pptx"
    shutil.copy(template_path, work_path)
    env = {**os.environ, "PYTHONIOENCODING": "utf-8"}
    for _ in range(1, n_slides_needed):
        result = subprocess.run(
            [sys.executable, str(scripts_dir / "add_slide.py"), str(work_path),
             "slide1.xml", "-o", str(work_path)],
            capture_output=True, text=True, encoding="utf-8", env=env
        )
        if result.returncode != 0:
            raise RuntimeError(f"add_slide.py failed: {result.stderr}")

    # 2) 각 슬라이드에 데이터 채우기
    prs = Presentation(work_path)
    chunks = [records[i:i + PER_SLIDE] for i in range(0, len(records), PER_SLIDE)]
    if not chunks:
        chunks = [[]]
    for slide, chunk in zip(prs.slides, chunks):
        fill_slide(slide, chunk)

    prs.save(output_path)
    return n_slides_needed


if __name__ == "__main__":
    sample_text = """■ 배장우 수업 시간 동안 양호한 자세와 태도를 유지하며 수업에 임하는 모습을 보였습니다. 타인에게 방해가 되는 행동은 관찰되지 않았으며, 이는 주변을 배려하며 안정적으로 수업에 참여하는 태도를 잘 보여주고 있습니다. 전반적으로 차분하고 무난한 수업 참여 자세를 유지한 것으로 확인되었습니다"""

    records = parse_records(sample_text)
    print("파싱된 레코드:", records)

    n = build_pptx("Template.pptx", records, "test_1record.pptx")
    print(f"생성된 슬라이드 수: {n}")
