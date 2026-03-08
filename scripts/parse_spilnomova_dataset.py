#!/usr/bin/env python3
"""
Parse SPILNOMOVA Excel metadata + transcript books into a clean NLP dataset.

No third-party dependencies required.
"""

from __future__ import annotations

import argparse
import json
import re
import unicodedata
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from zipfile import ZipFile

NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
NS_PKGREL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
SCRIPT_DIR = Path(__file__).resolve().parent


def normalize_key(value: str) -> str:
    value = unicodedata.normalize("NFKC", value or "")
    value = value.lower().strip()
    return re.sub(r"[\W_]+", "", value, flags=re.UNICODE)


def normalize_filename(value: str) -> str:
    value = unicodedata.normalize("NFKC", value or "")
    value = value.lower().strip()
    stem = str(Path(value).stem)
    return re.sub(r"[\W_]+", "", stem, flags=re.UNICODE)


def clean_text(value: Any) -> str:
    text = str(value or "").replace("\xa0", " ").strip()
    return re.sub(r"\s+", " ", text, flags=re.UNICODE)


def parse_int_like(value: Any) -> int | None:
    raw = clean_text(value)
    if not raw:
        return None
    normalized = raw.replace(",", ".")
    try:
        number = float(normalized)
    except ValueError:
        return None
    if number.is_integer():
        return int(number)
    return int(round(number))


def excel_col_to_index(ref: str) -> int:
    col = re.match(r"([A-Z]+)", ref)
    if not col:
        return -1
    index = 0
    for char in col.group(1):
        index = index * 26 + (ord(char) - 64)
    return index - 1


def as_sheet_name(value: str) -> str:
    raw = clean_text(value)
    if re.fullmatch(r"\d+\.0+", raw):
        return raw.split(".", 1)[0]
    return raw


def is_child_speaker(value: str) -> bool:
    speaker = normalize_key(value)
    return any(token in speaker for token in ("дитин", "ребен", "child"))


class XlsxBook:
    def __init__(self, path: Path):
        self.path = path
        self._zip = ZipFile(path)
        self._shared_strings = self._read_shared_strings()
        self._sheet_targets = self._read_sheet_targets()
        self.sheet_names = list(self._sheet_targets.keys())

    def close(self) -> None:
        self._zip.close()

    def _read_shared_strings(self) -> list[str]:
        if "xl/sharedStrings.xml" not in self._zip.namelist():
            return []
        root = ET.fromstring(self._zip.read("xl/sharedStrings.xml"))
        out: list[str] = []
        for si in root.findall(f"{NS_MAIN}si"):
            text = "".join((t.text or "") for t in si.iterfind(f".//{NS_MAIN}t"))
            out.append(text)
        return out

    def _read_sheet_targets(self) -> dict[str, str]:
        wb = ET.fromstring(self._zip.read("xl/workbook.xml"))
        rels = ET.fromstring(self._zip.read("xl/_rels/workbook.xml.rels"))
        rel_by_id = {
            rel.attrib["Id"]: rel.attrib["Target"]
            for rel in rels.findall(f"{NS_PKGREL}Relationship")
        }
        out: dict[str, str] = {}
        for sheet in wb.findall(f"{NS_MAIN}sheets/{NS_MAIN}sheet"):
            name = sheet.attrib["name"]
            rel_id = sheet.attrib[f"{NS_REL}id"]
            target = rel_by_id[rel_id]
            if not target.startswith("xl/"):
                target = f"xl/{target}"
            out[name] = target
        return out

    def read_sheet_rows(self, sheet_name: str) -> list[list[str]]:
        if sheet_name not in self._sheet_targets:
            available = ", ".join(self.sheet_names)
            raise KeyError(
                f"Worksheet '{sheet_name}' not found in '{self.path.name}'. "
                f"Available: {available}"
            )
        root = ET.fromstring(self._zip.read(self._sheet_targets[sheet_name]))
        rows: list[list[str]] = []
        for row in root.findall(f"{NS_MAIN}sheetData/{NS_MAIN}row"):
            cells: dict[int, str] = {}
            for cell in row.findall(f"{NS_MAIN}c"):
                idx = excel_col_to_index(cell.attrib.get("r", ""))
                if idx < 0:
                    continue
                cell_type = cell.attrib.get("t")
                value_node = cell.find(f"{NS_MAIN}v")
                if cell_type == "s" and value_node is not None and value_node.text is not None:
                    shared_index = int(value_node.text)
                    value = (
                        self._shared_strings[shared_index]
                        if shared_index < len(self._shared_strings)
                        else ""
                    )
                elif cell_type == "inlineStr":
                    value = "".join((t.text or "") for t in cell.findall(f".//{NS_MAIN}t"))
                else:
                    value = value_node.text if value_node is not None and value_node.text else ""
                cells[idx] = value
            if cells:
                width = max(cells) + 1
                materialized = [""] * width
                for idx, value in cells.items():
                    materialized[idx] = value
                rows.append(materialized)
        return rows


@dataclass
class FileIndex:
    by_name: dict[str, list[Path]]
    by_normalized_name: dict[str, list[Path]]

    @classmethod
    def build(cls, root: Path, patterns: tuple[str, ...]) -> "FileIndex":
        by_name: dict[str, list[Path]] = {}
        by_normalized: dict[str, list[Path]] = {}
        for pattern in patterns:
            for path in sorted(root.rglob(pattern)):
                if not path.is_file():
                    continue
                by_name.setdefault(path.name, []).append(path)
                by_normalized.setdefault(normalize_filename(path.name), []).append(path)
        return cls(by_name=by_name, by_normalized_name=by_normalized)

    def resolve(self, file_name: str) -> Path | None:
        exact = self.by_name.get(file_name)
        if exact:
            return exact[0]
        fuzzy = self.by_normalized_name.get(normalize_filename(file_name))
        if fuzzy and len(fuzzy) == 1:
            return fuzzy[0]
        return None


def find_header_map(row: list[str]) -> dict[str, int]:
    return {normalize_key(v): idx for idx, v in enumerate(row) if clean_text(v)}


def parse_main_table(book: XlsxBook) -> list[dict[str, str]]:
    rows = book.read_sheet_rows(book.sheet_names[0])
    if not rows:
        return []
    header = find_header_map(rows[0])
    required_columns = {
        "місто": "city",
        "навчальнийзаклад": "school",
        "група": "group",
        "дитина": "child_name",
        "аудіофайлname": "audio_filename",
        "транскриптfilename": "transcript_filename",
        "транскриптworksheetname": "transcript_sheet",
    }
    missing = [src for src in required_columns if src not in header]
    if missing:
        raise ValueError(f"Missing required columns in main table: {missing}")

    def get(row: list[str], key: str) -> str:
        idx = header[key]
        return clean_text(row[idx]) if idx < len(row) else ""

    out: list[dict[str, str]] = []
    for row in rows[1:]:
        audio_name = get(row, "аудіофайлname")
        if not audio_name:
            continue
        out.append(
            {
                "city": get(row, "місто"),
                "school": get(row, "навчальнийзаклад"),
                "group": get(row, "група"),
                "child_name": get(row, "дитина"),
                "audio_filename": audio_name,
                "transcript_filename": get(row, "транскриптfilename"),
                "transcript_sheet": as_sheet_name(get(row, "транскриптworksheetname")),
            }
        )
    return out


def parse_transcript_utterances(rows: list[list[str]]) -> list[dict[str, Any]]:
    if not rows:
        return []
    header_map = find_header_map(rows[0])
    speaker_idx = header_map.get(normalize_key("Хто говорить"))
    text_idx = header_map.get(normalize_key("Стенограма"))
    comment_idx = header_map.get(normalize_key("Коментар"))
    ts_idx = header_map.get(normalize_key("Час коментаря (хвилина:секунда)"))
    word_count_idx = header_map.get(normalize_key("Кількість слів"))

    if speaker_idx is None or text_idx is None:
        raise ValueError("Transcript sheet does not contain expected header columns")

    utterances: list[dict[str, Any]] = []
    for row in rows[1:]:
        speaker = clean_text(row[speaker_idx]) if speaker_idx < len(row) else ""
        text = clean_text(row[text_idx]) if text_idx < len(row) else ""
        if not text:
            continue
        utterances.append(
            {
                "speaker": speaker,
                "text": text,
                "comment": clean_text(row[comment_idx]) if comment_idx is not None and comment_idx < len(row) else "",
                "timestamp": clean_text(row[ts_idx]) if ts_idx is not None and ts_idx < len(row) else "",
                "word_count": (
                    parse_int_like(row[word_count_idx])
                    if word_count_idx is not None and word_count_idx < len(row)
                    else None
                ),
            }
        )
    return utterances


def write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def find_project_root() -> Path:
    candidates: list[Path] = []
    for base in (Path.cwd(), SCRIPT_DIR):
        candidates.extend([base, *base.parents])
    for candidate in candidates:
        input_dir = candidate / "01_INPUT FILES"
        if (input_dir / "01_AUDIO").exists() and (input_dir / "02_TRANSCRIPTS").exists():
            return candidate
    return SCRIPT_DIR.parent


def resolve_path_with_root(path: Path, root: Path) -> Path:
    if path.is_absolute():
        return path
    direct = Path.cwd() / path
    if direct.exists():
        return direct
    rooted = root / path
    if rooted.exists():
        return rooted
    scripted = SCRIPT_DIR / path
    if scripted.exists():
        return scripted
    return rooted


def to_rel(path: Path, root: Path) -> str:
    try:
        return path.resolve().relative_to(root.resolve()).as_posix()
    except Exception:
        return path.as_posix()


def autodetect_db_xlsx(root: Path) -> Path:
    candidates = [
        root / "SPILNOMOVA_DATABASE SAMPLE_TEST_11122025.xlsx",
        root / "01_INPUT FILES" / "SPILNOMOVA_DATABASE SAMPLE_TEST_11122025.xlsx",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    for candidate in sorted((root / "01_INPUT FILES").glob("*.xlsx")):
        if "database" in candidate.name.lower() or "sample_test" in candidate.name.lower():
            return candidate
    for candidate in sorted(root.glob("*.xlsx")):
        if "database" in candidate.name.lower() or "sample_test" in candidate.name.lower():
            return candidate
    for candidate in sorted(root.rglob("*.xlsx")):
        if "database" in candidate.name.lower():
            return candidate
    return candidates[0]


def autodetect_audio_root(root: Path) -> Path:
    candidate = root / "01_INPUT FILES" / "01_AUDIO"
    if candidate.exists():
        return candidate
    for found in root.rglob("01_AUDIO"):
        if found.is_dir():
            return found
    return candidate


def autodetect_transcripts_root(root: Path) -> Path:
    candidate = root / "01_INPUT FILES" / "02_TRANSCRIPTS"
    if candidate.exists():
        return candidate
    for found in root.rglob("02_TRANSCRIPTS"):
        if found.is_dir():
            return found
    return candidate


def build_dataset(
    db_xlsx: Path, audio_root: Path, transcripts_root: Path, output_dir: Path, project_root: Path
) -> dict[str, Any]:
    main_book = XlsxBook(db_xlsx)
    main_records = parse_main_table(main_book)
    main_book.close()

    audio_index = FileIndex.build(audio_root, ("*.m4a", "*.wav", "*.mp3"))
    transcript_index = FileIndex.build(transcripts_root, ("*.xlsx",))

    sessions: list[dict[str, Any]] = []
    turns_total = 0
    report_warnings: list[str] = []
    transcript_books: dict[Path, XlsxBook] = {}

    try:
        for record in main_records:
            audio_path = audio_index.resolve(record["audio_filename"])
            transcript_path = transcript_index.resolve(record["transcript_filename"])

            if not audio_path:
                report_warnings.append(f"Audio not found: {record['audio_filename']}")
                continue
            if not transcript_path:
                report_warnings.append(f"Transcript file not found: {record['transcript_filename']}")
                continue

            if transcript_path not in transcript_books:
                transcript_books[transcript_path] = XlsxBook(transcript_path)
            transcript_book = transcript_books[transcript_path]

            try:
                rows = transcript_book.read_sheet_rows(record["transcript_sheet"])
            except Exception as exc:
                report_warnings.append(
                    f"Sheet not found ({record['transcript_filename']} -> {record['transcript_sheet']}): {exc}"
                )
                continue

            try:
                parsed_utterances = parse_transcript_utterances(rows)
            except Exception as exc:
                report_warnings.append(
                    f"Transcript parse error ({record['transcript_filename']} -> {record['transcript_sheet']}): {exc}"
                )
                continue

            session_id = f"{record['city']}|{record['school']}|{record['group']}|{record['child_name']}"
            dialogue: list[dict[str, Any]] = []
            for utterance in parsed_utterances:
                is_child = int(is_child_speaker(utterance["speaker"]))
                dialogue.append(
                    {
                        "speaker": utterance["speaker"],
                        "role": "child" if is_child else "moderator",
                        "text": utterance["text"],
                        "word_count": utterance["word_count"],
                    }
                )

            word_count_total = sum((u["word_count"] or 0) for u in parsed_utterances)
            word_count_child_total = sum(
                (u["word_count"] or 0)
                for u in parsed_utterances
                if is_child_speaker(u["speaker"])
            )
            turns_total += len(dialogue)
            session_payload = {
                "session_id": session_id,
                "city": record["city"],
                "school": record["school"],
                "group": record["group"],
                "child_name": record["child_name"],
                "audio": {
                    "filename": record["audio_filename"],
                    "relpath": to_rel(audio_path, project_root),
                },
                "transcript": {
                    "filename": record["transcript_filename"],
                    "sheet": record["transcript_sheet"],
                    "relpath": to_rel(transcript_path, project_root),
                },
                "stats": {
                    "turns_total": len(dialogue),
                    "turns_child": sum(1 for u in parsed_utterances if is_child_speaker(u["speaker"])),
                    "word_count_annotated_total": word_count_total,
                    "word_count_annotated_child_total": word_count_child_total,
                },
                "dialogue": dialogue,
            }
            sessions.append(session_payload)
    finally:
        for book in transcript_books.values():
            book.close()

    output_dir.mkdir(parents=True, exist_ok=True)
    legacy_outputs = [
        "sessions.csv",
        "utterances.csv",
        "sessions.jsonl",
        "sessions_structured.jsonl",
        "utterances.jsonl",
    ]
    for file_name in legacy_outputs:
        legacy = output_dir / file_name
        if legacy.exists():
            legacy.unlink()

    dataset = {
        "meta": {
            "schema_version": "2.0",
            "generated_at_utc": datetime.now(timezone.utc).isoformat(),
            "project_root": to_rel(project_root, project_root),
            "db_xlsx": to_rel(db_xlsx, project_root),
            "audio_root": to_rel(audio_root, project_root),
            "transcripts_root": to_rel(transcripts_root, project_root),
            "sessions_count": len(sessions),
            "turns_count": turns_total,
        },
        "sessions": sessions,
    }
    write_json(output_dir / "dataset.json", dataset)

    report = {
        "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "project_root": to_rel(project_root, project_root),
        "db_xlsx": to_rel(db_xlsx, project_root),
        "audio_root": to_rel(audio_root, project_root),
        "transcripts_root": to_rel(transcripts_root, project_root),
        "output_dir": to_rel(output_dir, project_root),
        "input_records": len(main_records),
        "output_sessions": len(sessions),
        "output_turns": turns_total,
        "output_dataset_file": "dataset.json",
        "warnings_count": len(report_warnings),
        "warnings": report_warnings,
    }
    write_json(output_dir / "report.json", report)
    return report


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Parse SPILNOMOVA audio/transcript metadata into NLP-ready tables."
    )
    parser.add_argument(
        "--project-root",
        default=None,
        type=Path,
        help="Project root directory. If omitted, will be auto-detected.",
    )
    parser.add_argument(
        "--db-xlsx",
        default=None,
        type=Path,
        help="Path to main database XLSX. If omitted, will be auto-detected.",
    )
    parser.add_argument(
        "--audio-root",
        default=None,
        type=Path,
        help="Root directory with audio files. If omitted, will be auto-detected.",
    )
    parser.add_argument(
        "--transcripts-root",
        default=None,
        type=Path,
        help="Root directory with transcript XLSX files. If omitted, will be auto-detected.",
    )
    parser.add_argument(
        "--output-dir",
        default=None,
        type=Path,
        help="Directory for generated dataset files. Default: 02_OUTPUT/SPILNOMOVA_PARSED under project root.",
    )
    return parser


def main() -> None:
    args = build_arg_parser().parse_args()
    root = resolve_path_with_root(args.project_root, Path.cwd()) if args.project_root else find_project_root()
    db_xlsx = resolve_path_with_root(args.db_xlsx, root) if args.db_xlsx else autodetect_db_xlsx(root)
    audio_root = resolve_path_with_root(args.audio_root, root) if args.audio_root else autodetect_audio_root(root)
    transcripts_root = (
        resolve_path_with_root(args.transcripts_root, root)
        if args.transcripts_root
        else autodetect_transcripts_root(root)
    )
    output_dir = (
        resolve_path_with_root(args.output_dir, root)
        if args.output_dir
        else (root / "02_OUTPUT" / "SPILNOMOVA_PARSED")
    )

    report = build_dataset(
        db_xlsx=db_xlsx,
        audio_root=audio_root,
        transcripts_root=transcripts_root,
        output_dir=output_dir,
        project_root=root,
    )
    print(
        f"Done. Sessions: {report['output_sessions']}, "
        f"turns: {report['output_turns']}, "
        f"warnings: {report['warnings_count']}"
    )

if __name__ == "__main__":
    main()
