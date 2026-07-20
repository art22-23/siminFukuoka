from __future__ import annotations

import base64
import html
import json
import re
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parent
EXCEL_PATH = ROOT / "2026福岡県議一覧.xlsx"
TOPO_PATH = ROOT / "fukuoka_city_20230101.topojson"
OUTPUT_PATH = ROOT / "fukuoka-kenmap.html"
PHOTO_DATA_PATH = ROOT / "candidate_photos.json"
TRAVEL_COST_IMAGE_PATH = ROOT / "nishinippon-article-20260709.jpeg"

FULLWIDTH_DIGITS = str.maketrans("０１２３４５６７８９", "0123456789")
BY_ELECTION_LABELS = {
    "吉松源明": "2024年12月補選",
    "亀崎大介": "2024年12月補選",
    "佐藤かえで": "2025年3月補選",
    "吉岡れい子": "2025年3月補選",
    "となり祥平": "2025年3月補選",
    "うらただいじ": "2025年6月補選",
}

UNCONTESTED_2023_DISTRICTS = {
    "大牟田市",
    "飯塚市・嘉穂郡",
    "田川市",
    "大川市・三潴郡",
    "大野城市",
    "福津市",
    "宮若市・鞍手郡",
    "嘉麻市",
    "朝倉市・朝倉郡",
    "みやま市",
    "那珂川市",
    "糟屋郡",
    "田川郡",
    "築上郡・豊前市",
}


def clean_cell(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    text = text.replace("\u3000", " ")
    return re.sub(r"\s+", " ", text)


def seat_number(text: object) -> int:
    normalized = str(text or "").translate(FULLWIDTH_DIGITS)
    match = re.search(r"\d+", normalized)
    return int(match.group(0)) if match else 0


def is_travel_note(text: str) -> bool:
    normalized = text.translate(FULLWIDTH_DIGITS)
    return bool(re.search(r"\d+回", normalized) and ("万円" in normalized or "不明" in normalized))


def parse_travel_note(text: str) -> dict[str, object] | None:
    if not is_travel_note(text):
        return None
    return {"raw": text}


def parse_person_entry(text: str) -> dict[str, object]:
    parts = text.rsplit(" ", 1)
    if len(parts) == 2:
        name, party = parts
        by_election_label = "" if "×" in party else BY_ELECTION_LABELS.get(name, "")
        return {"name": name, "party": party, "byElection": bool(by_election_label), "byElectionLabel": by_election_label}
    by_election_label = BY_ELECTION_LABELS.get(text, "")
    return {"name": text, "party": "", "byElection": bool(by_election_label), "byElectionLabel": by_election_label}


def load_candidate_photos() -> dict[str, dict[str, str]]:
    if not PHOTO_DATA_PATH.exists():
        return {}
    data = json.loads(PHOTO_DATA_PATH.read_text(encoding="utf-8"))
    photos = data.get("photos", data)
    return photos if isinstance(photos, dict) else {}


def attach_photo(entry: dict[str, object], photos: dict[str, dict[str, str]]) -> None:
    photo = photos.get(str(entry.get("name", "")))
    if not photo:
        return
    image_url = photo.get("imageUrl", "")
    if not image_url:
        return
    entry["photoUrl"] = image_url
    entry["profileUrl"] = photo.get("profileUrl", "")
    entry["photoSource"] = photo.get("source", "選挙ドットコム")


def image_data_uri(path: Path) -> str:
    if not path.exists():
        return ""
    mime = "image/jpeg" if path.suffix.lower() in {".jpg", ".jpeg"} else "image/png"
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:{mime};base64,{encoded}"


def topo_area_name(properties: dict[str, object]) -> str:
    city = properties.get("N03_003") or ""
    ward_or_town = properties.get("N03_004") or ""
    if city and ward_or_town:
        return f"{city}{ward_or_town}"
    return str(ward_or_town or city)


def load_area_names(topo: dict[str, object]) -> list[str]:
    geometries = topo["objects"]["city"]["geometries"]
    return [topo_area_name(g["properties"]) for g in geometries]


def parse_excel(candidate_photos: dict[str, dict[str, str]] | None = None) -> list[dict[str, object]]:
    df = pd.read_excel(EXCEL_PATH, header=None, dtype=object)
    photos = candidate_photos or {}
    records: list[dict[str, object]] = []
    current: dict[str, object] | None = None

    for _, row in df.iloc[3:].iterrows():
        district = clean_cell(row.iloc[1]) if len(row) > 1 else ""
        seats = clean_cell(row.iloc[2]) if len(row) > 2 else ""
        cells = [clean_cell(row.iloc[col]) if len(row) > col else "" for col in range(3, 8)]

        if district:
            current = {
                "name": district,
                "seats": seats,
                "columns": [[] for _ in range(5)],
            }
            records.append(current)

        if current is None:
            continue

        for index, text in enumerate(cells):
            if text:
                current["columns"][index].append(text)

    normalized: list[dict[str, object]] = []
    for record in records:
        slots = []
        seat_count = seat_number(record["seats"])

        for values in record["columns"]:
            if not values:
                continue
            entries = []
            for value in values:
                parsed = parse_travel_note(value)
                if parsed:

                    entries.append({"kind": "travel", "text": value})
                else:
                    person = parse_person_entry(value)
                    entry = {"kind": "person", "text": value, "name": person["name"], "party": person["party"], "byElection": person["byElection"], "byElectionLabel": person["byElectionLabel"]}
                    attach_photo(entry, photos)
                    entries.append(entry)
            slots.append({"entries": entries})

        normalized.append(
            {
                "name": record["name"],
                "seats": record["seats"],
                "seatNumber": seat_count,
                "singleMember": seat_count == 1,
                "uncontested2023": record["name"] in UNCONTESTED_2023_DISTRICTS,
                "slots": slots,
            }
        )

    return normalized

def load_header_note() -> str:
    df = pd.read_excel(EXCEL_PATH, header=None, dtype=object, nrows=1)
    if df.shape[1] <= 3:
        return ""
    return clean_cell(df.iloc[0, 3])


def expand_district_areas(district_name: str, area_names: list[str]) -> list[str]:
    area_set = set(area_names)
    expanded: list[str] = []

    for part in district_name.split("・"):
        part = part.strip()
        if not part:
            continue
        if part in area_set:
            expanded.append(part)
            continue
        if part.endswith("郡"):
            matches = [name for name in area_names if name.startswith(part)]
            expanded.extend(matches)
            continue
        raise ValueError(f"地図境界に対応できない選挙区名です: {district_name} / {part}")

    return expanded


def attach_areas(records: list[dict[str, object]], area_names: list[str]) -> dict[str, str]:
    area_to_district: dict[str, str] = {}
    for record in records:
        areas = expand_district_areas(str(record["name"]), area_names)
        record["areas"] = areas
        for area in areas:
            if area in area_to_district:
                raise ValueError(f"区域が複数の選挙区に割り当てられています: {area}")
            area_to_district[area] = str(record["name"])
    return area_to_district


def json_for_script(value: object) -> str:
    text = json.dumps(value, ensure_ascii=False, separators=(",", ":"))
    return text.replace("</", "<\\/")


def build_html(
    topo: dict[str, object],
    records: list[dict[str, object]],
    area_to_district: dict[str, str],
    header_note: str,
    travel_cost_image_data_uri: str,
) -> str:
    topo_json = json_for_script(topo)
    records_json = json_for_script(records)
    area_map_json = json_for_script(area_to_district)
    header_note_html = html.escape(header_note)
    travel_cost_section = ""
    if travel_cost_image_data_uri:
        travel_cost_section = f"""
    <section class="expense-panel" aria-label="海外視察費総額一覧">
      <div class="expense-head">
        <h2>海外視察費総額一覧</h2>
        <p>海外渡航費の出典: <a href="https://www.nishinippon.co.jp/item/1512796/">西日本新聞meの記事</a></p>
      </div>
      <figure>
        <img src="{html.escape(travel_cost_image_data_uri, quote=True)}" alt="福岡県議会の海外視察費総額一覧" loading="lazy">
        <figcaption>海外渡航費の記載は、西日本新聞me「【渡航した全議員の一覧】福岡県議会の公費海外視察　活動内容と旅費」を元にしています。</figcaption>
      </figure>
    </section>
"""

    template = """<!doctype html>
<html lang="ja">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>何が隠されている？　福岡県政問題　福岡県議一覧2026</title>
  <style>
    :root {
      color-scheme: light;
      --bg: #f6f8f6;
      --surface: #ffffff;
      --surface-2: #eef4f1;
      --text: #1f2a27;
      --muted: #63716c;
      --line: #cdd8d3;
      --strong: #214d46;
      --accent: #b94f36;
      --map-empty: #e4e9e7;
      --map-multi: #d4ece5;
      --map-single: #f2ce7a;
      --map-single-strong: #a45f17;
      --map-uncontested-line: #b94f36;
      --map-uncontested-soft: rgba(185, 79, 54, 0.16);
      --map-low: #d4ece5;
      --map-mid: #8fd0c2;
      --map-high: #efba6b;
      --map-top: #c95645;
      --focus: #1a766b;
      --shadow: 0 18px 42px rgba(31, 42, 39, 0.12);
      font-family: "Yu Gothic", "Hiragino Sans", "Meiryo", system-ui, sans-serif;
    }

    * { box-sizing: border-box; }

    body {
      margin: 0;
      background:
        linear-gradient(180deg, rgba(255,255,255,0.92), rgba(246,248,246,0.96)),
        radial-gradient(circle at top left, rgba(143,208,194,0.24), transparent 34%),
        var(--bg);
      color: var(--text);
    }

    .app {
      width: min(1540px, calc(100% - 32px));
      margin: 0 auto;
      padding: 28px 0 24px;
    }

    header {
      display: flex;
      align-items: end;
      justify-content: space-between;
      gap: 18px;
      margin-bottom: 18px;
    }

    h1 {
      margin: 0;
      font-size: 36px;
      font-weight: 500;
      letter-spacing: 0;
    }

    .lead {
      margin: 8px 0 0;
      color: var(--muted);
      max-width: 760px;
      line-height: 1.7;
    }

    .summary {
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
      justify-content: flex-end;
    }

    .summary span {
      display: inline-flex;
      align-items: baseline;
      gap: 4px;
      padding: 8px 10px;
      background: var(--surface);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: 0 8px 18px rgba(31, 42, 39, 0.06);
      color: var(--muted);
      white-space: nowrap;
    }

    .summary strong {
      color: var(--text);
      font-size: 18px;
    }

    .layout {
      display: grid;
      grid-template-columns: minmax(0, 2.25fr) minmax(340px, 0.65fr);
      gap: 18px;
      align-items: stretch;
    }

    .map-panel,
    .detail {
      background: rgba(255,255,255,0.92);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
    }

    .map-panel {
      padding: 16px;
      min-width: 0;
    }

    .toolbar {
      display: flex;
      flex-wrap: wrap;
      align-items: end;
      justify-content: space-between;
      gap: 12px;
      margin-bottom: 10px;
    }

    .map-legend {
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      gap: 8px 14px;
      margin: 0 0 12px;
      color: var(--muted);
      font-size: 13px;
    }

    .map-legend span {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      white-space: nowrap;
    }

    .legend-swatch {
      width: 18px;
      height: 14px;
      border-radius: 4px;
      border: 1px solid rgba(31, 42, 39, 0.36);
      background: var(--map-multi);
      flex: 0 0 auto;
    }

    .legend-swatch.single {
      background: var(--map-single);
      border-color: var(--map-single-strong);
    }

    .legend-swatch.uncontested {
      background:
        repeating-linear-gradient(135deg, transparent 0 4px, rgba(185, 79, 54, 0.58) 4px 6px),
        var(--map-multi);
      border-color: var(--map-uncontested-line);
    }

    label {
      display: grid;
      gap: 5px;
      color: var(--muted);
      font-size: 13px;
      min-width: min(100%, 280px);
    }

    select,
    button {
      font: inherit;
      border-radius: 8px;
      border: 1px solid var(--line);
      background: var(--surface);
      color: var(--text);
      min-height: 40px;
    }

    select {
      padding: 8px 36px 8px 10px;
      width: 100%;
    }

    button {
      padding: 8px 14px;
      cursor: pointer;
    }

    button:hover,
    select:hover {
      border-color: var(--focus);
    }

    button:focus,
    select:focus {
      outline: 3px solid rgba(26, 118, 107, 0.22);
      outline-offset: 2px;
    }

    .map-wrap {
      position: relative;
      width: 100%;
      overflow: hidden;
      border-radius: 8px;
      background:
        linear-gradient(180deg, rgba(238,244,241,0.72), rgba(255,255,255,0.88));
      border: 1px solid var(--line);
    }

    svg {
      display: block;
      width: 100%;
      height: auto;
      aspect-ratio: 1.18 / 1;
    }

    .area {
      stroke: rgba(31, 42, 39, 0.42);
      stroke-width: 0.72;
      cursor: pointer;
      transition: fill 160ms ease, stroke-width 160ms ease, opacity 160ms ease;
    }

    .area:hover {
      stroke: var(--strong);
      stroke-width: 1.55;
    }

    .area.is-muted {
      opacity: 0.55;
    }

    .area.is-uncontested {
      stroke: var(--map-uncontested-line);
      stroke-width: 1.12;
    }

    .area.is-selected {
      stroke: var(--strong);
      stroke-width: 2.25;
      opacity: 1;
    }

    .uncontested-layer,
    .seat-label-layer {
      pointer-events: none;
    }

    .uncontested-overlay {
      fill: url(#uncontested-hatch);
      stroke: var(--map-uncontested-line);
      stroke-width: 1.35;
      opacity: 0.84;
      pointer-events: none;
    }

    .uncontested-overlay.is-muted {
      opacity: 0.28;
    }

    .uncontested-overlay.is-selected {
      opacity: 0.96;
      stroke-width: 2;
    }

    .seat-label {
      pointer-events: none;
    }

    .seat-label circle {
      fill: rgba(255,255,255,0.94);
      stroke: rgba(31, 42, 39, 0.42);
      stroke-width: 1.1;
    }

    .seat-label.is-single circle {
      fill: var(--map-single);
      stroke: var(--map-single-strong);
      stroke-width: 1.45;
    }

    .seat-label.is-uncontested circle {
      stroke: var(--map-uncontested-line);
      stroke-width: 1.8;
    }

    .seat-label.is-muted {
      opacity: 0.56;
    }

    .seat-label.is-selected {
      opacity: 1;
    }

    .seat-label text {
      font-size: 12px;
      font-weight: 500;
      text-anchor: middle;
      paint-order: stroke;
      stroke: rgba(255,255,255,0.82);
      stroke-width: 2px;
      fill: var(--text);
    }

    .label-layer text {
      font-size: 13px;
      font-weight: 500;
      paint-order: stroke;
      stroke: rgba(255,255,255,0.86);
      stroke-width: 4px;
      fill: var(--text);
      pointer-events: none;
    }

    .tooltip {
      position: absolute;
      z-index: 4;
      display: none;
      max-width: 260px;
      padding: 8px 10px;
      border-radius: 8px;
      background: rgba(31,42,39,0.92);
      color: #fff;
      font-size: 12px;
      line-height: 1.5;
      pointer-events: none;
      box-shadow: 0 10px 24px rgba(31, 42, 39, 0.22);
    }

    .detail {
      padding: 18px;
      position: sticky;
      top: 14px;
      align-self: stretch;
      min-height: 100%;
    }

    .detail h2 {
      margin: 0;
      font-size: 24px;
      line-height: 1.35;
      letter-spacing: 0;
    }

    .detail .sub {
      margin: 6px 0 14px;
      color: var(--muted);
      line-height: 1.65;
    }
    .seat-line {
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      gap: 8px;
      margin: 0 0 12px;
      color: var(--muted);
      font-weight: 500;
    }

    .district-flag {
      display: inline-flex;
      align-items: center;
      padding: 2px 7px;
      border-radius: 999px;
      border: 1px solid rgba(164, 95, 23, 0.36);
      background: rgba(242, 206, 122, 0.34);
      color: var(--strong);
      font-size: 12px;
      font-weight: 500;
      white-space: nowrap;
    }

    .district-flag.uncontested {
      border-color: rgba(185, 79, 54, 0.38);
      background: var(--map-uncontested-soft);
      color: var(--accent);
    }

    .member-list {
      display: grid;
      gap: 8px;
      margin: 0;
      padding: 0;
      list-style: none;
    }

    .member {
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 10px;
      background: var(--surface);
    }

    .person-row {
      display: grid;
      grid-template-columns: 90px 1fr;
      gap: 12px;
      align-items: center;
      min-width: 0;
    }

    .person-row + .person-row,
    .travel + .person-row {
      border-top: 1px solid var(--line);
      padding-top: 8px;
      margin-top: 8px;
    }

    .person-photo-link,
    .person-photo-missing {
      width: 88px;
      height: 88px;
      border-radius: 8px;
    }

    .person-photo-link {
      display: block;
      overflow: hidden;
      border: 1px solid var(--line);
      background: var(--surface-2);
    }

    .person-photo {
      display: block;
      width: 100%;
      height: 100%;
      object-fit: cover;
    }

    .person-photo-missing {
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 4px;
      border: 1px dashed var(--line);
      background: var(--surface-2);
      color: var(--muted);
      font-size: 12px;
      line-height: 1.25;
      text-align: center;
    }

    .person-main {
      display: grid;
      gap: 5px;
      min-width: 0;
    }

    .person-line {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      align-items: center;
      min-width: 0;
    }

    .party {
      display: inline-flex;
      justify-content: center;
      align-items: center;
      padding: 3px 8px;
      border-radius: 999px;
      background: var(--surface-2);
      color: var(--strong);
      font-weight: 500;
      white-space: nowrap;
    }

    .person-name {
      min-width: 0;
      font-weight: 500;
      word-break: keep-all;
      overflow-wrap: anywhere;
    }
    .by-election {
      display: inline-flex;
      align-items: center;
      margin-left: 0;
      padding: 2px 7px;
      border-radius: 999px;
      border: 1px solid rgba(185, 79, 54, 0.34);
      background: rgba(185, 79, 54, 0.10);
      color: var(--accent);
      font-size: 12px;
      font-weight: 500;
      white-space: nowrap;
      vertical-align: middle;
    }

    .travel {
      margin: 8px 0 0 102px;
      color: var(--accent);
      font-weight: 500;
      line-height: 1.55;
    }
    .areas {
      margin-top: 14px;
      color: var(--muted);
      line-height: 1.6;
      font-size: 13px;
    }

    .expense-panel {
      margin: 18px 0;
      padding: 18px;
      background: rgba(255,255,255,0.92);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
    }

    .expense-head {
      display: flex;
      flex-wrap: wrap;
      justify-content: space-between;
      align-items: baseline;
      gap: 8px 14px;
      margin-bottom: 12px;
    }

    .expense-head h2 {
      margin: 0;
      font-size: 22px;
      font-weight: 500;
      letter-spacing: 0;
    }

    .expense-head p {
      margin: 0;
      color: var(--muted);
      font-size: 13px;
    }

    .expense-panel figure {
      margin: 0;
    }

    .expense-panel img {
      display: block;
      width: min(100%, 920px);
      height: auto;
      margin: 0 auto;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--surface);
    }

    .expense-panel figcaption {
      margin-top: 8px;
      color: var(--muted);
      font-size: 12px;
      text-align: center;
    }

    footer {
      margin-top: 18px;
      color: var(--muted);
      font-size: 12px;
      line-height: 1.7;
    }

    a {
      color: var(--strong);
    }

    @media (max-width: 900px) {
      header {
        align-items: start;
        flex-direction: column;
      }
      .summary {
        justify-content: flex-start;
      }
      .layout {
        grid-template-columns: 1fr;
        align-items: start;
      }
      .detail {
        position: static;
        min-height: 0;
        align-self: auto;
      }
    }

    @media (max-width: 560px) {
      .app {
        width: min(100% - 20px, 1540px);
        padding-top: 18px;
      }
      .map-panel,
      .detail,
      .expense-panel {
        padding: 12px;
      }
      .person-row {
        grid-template-columns: 64px 1fr;
        gap: 10px;
      }
      .person-photo-link,
      .person-photo-missing {
        width: 60px;
        height: 60px;
      }
      .person-photo-missing {
        font-size: 10px;
      }
      .travel {
        margin-left: 74px;
      }
      svg {
        aspect-ratio: 0.86 / 1;
      }
    }
  </style>
</head>
<body>
  <div class="app">
    <header>
      <div>
        <h1>何が隠されている？　福岡県政問題　福岡県議一覧2026</h1>
      </div>
      <div class="summary" aria-label="データ概要">
        <span><strong id="district-count">0</strong>選挙区</span>
        <span><strong id="seat-count">0</strong>定数</span>
        <span><strong id="single-member-count">0</strong>一人区</span>
        <span><strong id="uncontested-count">0</strong>2023年無投票</span>
        <span class="note">__HEADER_NOTE__</span>
      </div>
    </header>

__TRAVEL_COST_SECTION__

    <main class="layout">
      <section class="map-panel" aria-label="福岡県内の選挙区地図">
        <div class="toolbar">
          <label for="district-select">
            選挙区
            <select id="district-select">
              <option value="">地図から選択</option>
            </select>
          </label>
          <button type="button" id="reset-button">全体表示</button>
        </div>

        <div class="map-legend" aria-label="地図の凡例">
          <span><i class="legend-swatch single" aria-hidden="true"></i>定数1（一人区）</span>
          <span><i class="legend-swatch multi" aria-hidden="true"></i>定数2以上</span>
          <span><i class="legend-swatch uncontested" aria-hidden="true"></i>2023年無投票</span>
        </div>

        <div class="map-wrap" id="map-wrap">
          <svg id="map" role="img" aria-labelledby="map-title map-desc" viewBox="0 0 920 780">
            <title id="map-title">福岡県の市区町村・行政区別の県議会選挙区地図</title>
            <desc id="map-desc">各区域に定数を表示し、一人区と2023年無投票選挙区を色と斜線で示します。選挙区を選ぶと右側に議員情報が表示されます。</desc>
            <defs>
              <pattern id="uncontested-hatch" patternUnits="userSpaceOnUse" width="8" height="8">
                <path d="M-2,8 L8,-2 M0,10 L10,0" stroke="var(--map-uncontested-line)" stroke-width="1.8" opacity="0.62"></path>
              </pattern>
            </defs>
            <g id="area-layer"></g>
            <g id="uncontested-layer" class="uncontested-layer"></g>
            <g id="seat-label-layer" class="seat-label-layer"></g>
            <g id="label-layer" class="label-layer"></g>
          </svg>
          <div id="tooltip" class="tooltip" role="status"></div>
        </div>
      </section>

      <aside class="detail" id="detail" aria-live="polite"></aside>
    </main>

    <footer>
      議員データ: 2026福岡県議一覧.xlsx / 顔写真: <a href="https://go2senkyo.com/local/senkyo/23137">選挙ドットコム</a> / 境界データ: <a href="https://geoshape.ex.nii.ac.jp/city/choropleth/40_city.html">Geoshape 市区町村TopoJSON 2023-01-01</a>（ODbL） / 2023年無投票: <a href="https://www.pref.fukuoka.lg.jp/contents/r5kengisenkyo.html">福岡県選挙管理委員会「令和5年4月9日執行 福岡県議会議員一般選挙」</a>
    </footer>
  </div>

  <script src="https://cdn.jsdelivr.net/npm/d3@7.9.0/dist/d3.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/topojson-client@3.1.0/dist/topojson-client.min.js"></script>
  <script id="topology-data" type="application/json">__TOPO_JSON__</script>
  <script id="district-data" type="application/json">__RECORDS_JSON__</script>
  <script id="area-map-data" type="application/json">__AREA_MAP_JSON__</script>
  <script>
    const topology = JSON.parse(document.getElementById("topology-data").textContent);
    const districts = JSON.parse(document.getElementById("district-data").textContent);
    const areaToDistrict = JSON.parse(document.getElementById("area-map-data").textContent);

    const byDistrict = new Map(districts.map((district) => [district.name, district]));
    const svg = d3.select("#map");
    const areaLayer = d3.select("#area-layer");
    const labelLayer = d3.select("#label-layer");
    const uncontestedLayer = d3.select("#uncontested-layer");
    const seatLabelLayer = d3.select("#seat-label-layer");
    const tooltip = d3.select("#tooltip");
    const select = document.getElementById("district-select");
    const resetButton = document.getElementById("reset-button");
    const detail = document.getElementById("detail");

    if (!window.d3 || !window.topojson) {
      detail.textContent = "地図ライブラリを読み込めませんでした。インターネット接続を確認してください。";
      throw new Error("D3 or TopoJSON client is unavailable.");
    }

    const featureCollection = topojson.feature(topology, topology.objects.city);
    const features = featureCollection.features.map((feature) => {
      const props = feature.properties || {};
      const city = props.N03_003 || "";
      const ward = props.N03_004 || "";
      const areaName = city && ward ? `${city}${ward}` : (ward || city);
      const districtName = areaToDistrict[areaName] || "";
      return {
        ...feature,
        properties: {
          ...props,
          areaName,
          districtName,
        },
      };
    });

    const featuresByDistrict = new Map();
    features.forEach((feature) => {
      const name = feature.properties.districtName;
      if (!name) return;
      if (!featuresByDistrict.has(name)) {
        featuresByDistrict.set(name, []);
      }
      featuresByDistrict.get(name).push(feature);
    });

    let selectedDistrict = "";

    function seatNumber(text) {
      const normalized = String(text || "").replace(/[０-９]/g, (digit) => "０１２３４５６７８９".indexOf(digit));
      const match = normalized.match(/\\d+/);
      return match ? Number(match[0]) : 0;
    }

    function escapeHtml(value) {
      return String(value || "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
    }

    function recordFor(feature) {
      return byDistrict.get(feature.properties.districtName);
    }

    function districtSeatNumber(record) {
      return Number(record?.seatNumber || seatNumber(record?.seats));
    }

    function isSingleMember(record) {
      return districtSeatNumber(record) === 1;
    }

    function fillFor(feature) {
      const record = recordFor(feature);
      if (!record) return "var(--map-empty)";
      return isSingleMember(record) ? "var(--map-single)" : "var(--map-multi)";
    }

    function mapSize() {
      return {
        width: 980,
        height: window.matchMedia("(max-width: 560px)").matches ? 1120 : 820,
        padding: 10,
      };
    }

    function tooltipLines(feature) {
      const record = recordFor(feature);
      const lines = [`<strong>${escapeHtml(feature.properties.areaName)}</strong>`];
      if (!record) {
        lines.push("表示データなし");
        return lines;
      }
      const seats = districtSeatNumber(record);
      lines.push(escapeHtml(record.name));
      if (seats) {
        lines.push(`定数 ${seats}${seats === 1 ? "（一人区）" : ""}`);
      }
      if (record.uncontested2023) {
        lines.push("2023年無投票");
      }
      return lines;
    }

    function renderMap() {
      const { width, height, padding } = mapSize();
      svg.attr("viewBox", `0 0 ${width} ${height}`);

      const projection = d3.geoMercator().fitExtent([[padding, padding], [width - padding, height - padding]], {
        type: "FeatureCollection",
        features,
      });
      const path = d3.geoPath(projection);

      areaLayer.selectAll("path")
        .data(features, (feature) => feature.properties.N03_007 || feature.properties.areaName)
        .join("path")
        .attr("class", (feature) => areaClass(feature))
        .attr("d", path)
        .attr("fill", fillFor)
        .on("mousemove", (event, feature) => {
          tooltip
            .style("display", "block")
            .style("left", `${event.offsetX + 12}px`)
            .style("top", `${event.offsetY + 12}px`)
            .html(tooltipLines(feature).join("<br>"));
        })
        .on("mouseleave", () => tooltip.style("display", "none"))
        .on("click", (_, feature) => {
          if (feature.properties.districtName) {
            selectDistrict(feature.properties.districtName);
          } else {
            selectedDistrict = "";
            select.value = "";
            updateSelection();
            renderNoData(feature.properties.areaName);
          }
        });

      uncontestedLayer.selectAll("path")
        .data(features.filter((feature) => recordFor(feature)?.uncontested2023), (feature) => feature.properties.N03_007 || feature.properties.areaName)
        .join("path")
        .attr("class", (feature) => overlayClass(feature))
        .attr("d", path);

      drawSeatLabels(path);
      drawSelectedLabels(path);
    }

    function areaClass(feature) {
      const record = recordFor(feature);
      const classes = ["area"];
      if (isSingleMember(record)) classes.push("is-single-member");
      if (record?.uncontested2023) classes.push("is-uncontested");
      if (selectedDistrict) {
        if (feature.properties.districtName === selectedDistrict) {
          classes.push("is-selected");
        } else {
          classes.push("is-muted");
        }
      }
      return classes.join(" ");
    }

    function overlayClass(feature) {
      const classes = ["uncontested-overlay"];
      if (selectedDistrict) {
        if (feature.properties.districtName === selectedDistrict) {
          classes.push("is-selected");
        } else {
          classes.push("is-muted");
        }
      }
      return classes.join(" ");
    }

    function seatLabelClass(datum) {
      const classes = ["seat-label"];
      if (datum.seatCount === 1) classes.push("is-single");
      if (datum.record.uncontested2023) classes.push("is-uncontested");
      if (selectedDistrict) {
        if (datum.record.name === selectedDistrict) {
          classes.push("is-selected");
        } else {
          classes.push("is-muted");
        }
      }
      return classes.join(" ");
    }

    function drawSeatLabels(path) {
      const groups = districts.map((record) => {
        const districtFeatures = featuresByDistrict.get(record.name) || [];
        if (!districtFeatures.length) return null;
        const collection = { type: "FeatureCollection", features: districtFeatures };
        const [x, y] = path.centroid(collection);
        const seatCount = districtSeatNumber(record);
        if (!Number.isFinite(x) || !Number.isFinite(y) || !seatCount) return null;
        return { record, x, y, seatCount };
      }).filter(Boolean);

      const labels = seatLabelLayer.selectAll("g")
        .data(groups, (datum) => datum.record.name)
        .join((enter) => {
          const group = enter.append("g");
          group.append("circle");
          group.append("text").attr("dy", "0.35em");
          return group;
        });

      labels
        .attr("class", seatLabelClass)
        .attr("transform", (datum) => `translate(${datum.x},${datum.y})`);

      labels.select("circle")
        .attr("r", (datum) => datum.seatCount >= 4 ? 15 : 13);

      labels.select("text")
        .text((datum) => datum.seatCount);
    }

    function drawSelectedLabels(path) {
      const selectedFeatures = selectedDistrict
        ? features.filter((feature) => feature.properties.districtName === selectedDistrict)
        : [];

      labelLayer.selectAll("text")
        .data(selectedFeatures, (feature) => feature.properties.areaName)
        .join("text")
        .attr("x", (feature) => path.centroid(feature)[0])
        .attr("y", (feature) => path.centroid(feature)[1])
        .attr("text-anchor", "middle")
        .attr("dy", "0.35em")
        .text((feature) => shortAreaName(feature.properties.areaName));
    }

    function shortAreaName(name) {
      return String(name)
        .replace("北九州市", "")
        .replace("福岡市", "")
        .replace(/^.+郡/, "");
    }

    function updateSelection() {
      areaLayer.selectAll("path").attr("class", (feature) => areaClass(feature));
      uncontestedLayer.selectAll("path").attr("class", (feature) => overlayClass(feature));
      seatLabelLayer.selectAll("g").attr("class", seatLabelClass);
      const { width, height, padding } = mapSize();
      const projection = d3.geoMercator().fitExtent([[padding, padding], [width - padding, height - padding]], {
        type: "FeatureCollection",
        features,
      });
      drawSelectedLabels(d3.geoPath(projection));
    }
    function selectDistrict(name) {
      selectedDistrict = name || "";
      select.value = selectedDistrict;
      updateSelection();
      if (selectedDistrict) {
        renderDetail(byDistrict.get(selectedDistrict));
      } else {
        renderIntro();
      }
    }

    function renderIntro() {
      detail.innerHTML = `
        <h2>選挙区を選択</h2>
        <p class="sub">定数1（一人区）と2023年無投票選挙区が地図上で確認できます。</p>
      `;
    }

    function renderNoData(areaName) {
      detail.innerHTML = `
        <h2>${escapeHtml(areaName)}</h2>
        <p class="sub">この区域はExcelの選挙区データに対応していません。</p>
      `;
    }

    function renderDetail(record) {
      if (!record) {
        renderIntro();
        return;
      }

      const districtFlags = [
        districtSeatNumber(record) === 1 ? `<span class="district-flag">一人区</span>` : "",
        record.uncontested2023 ? `<span class="district-flag uncontested">2023年無投票</span>` : "",
      ].filter(Boolean).join("");

      const members = record.slots.map((slot) => {
        const entries = slot.entries.map((entry) => {
          if (entry.kind === "travel") {
            return `<div class="travel">${escapeHtml(entry.text)}</div>`;
          }
          const party = entry.party ? `<span class="party">${escapeHtml(entry.party)}</span>` : "";
          const byElectionLabel = entry.byElectionLabel || (entry.byElection ? "補選" : "");
          const badge = byElectionLabel ? `<span class="by-election">${escapeHtml(byElectionLabel)}</span>` : "";
          const rawName = entry.name || entry.text;
          const name = escapeHtml(rawName);
          const photoAlt = `${escapeHtml(rawName)}の顔写真`;
          const photo = entry.photoUrl
            ? `<a class="person-photo-link" href="${escapeHtml(entry.profileUrl || entry.photoUrl)}" target="_blank" rel="noopener" aria-label="${name}のプロフィールを開く"><img class="person-photo" src="${escapeHtml(entry.photoUrl)}" alt="${photoAlt}" loading="lazy"></a>`
            : `<div class="person-photo-missing" aria-label="${name}の顔写真は未掲載">写真<br>未掲載</div>`;
          return `<div class="person-row">${photo}<div class="person-main"><div class="person-line">${party}<span class="person-name">${name}</span>${badge}</div></div></div>`;
        }).join("");
        return `<li class="member">${entries}</li>`;
      }).join("");

      detail.innerHTML = `
        <h2>${escapeHtml(record.name)}</h2>
        <p class="sub">${escapeHtml(record.areas.join("、"))}</p>
        <p class="seat-line"><span>定数 ${escapeHtml(record.seats || "-")}</span>${districtFlags}</p>
        <ul class="member-list">${members}</ul>
      `;
    }

    function populateControls() {
      districts.forEach((district) => {
        const option = document.createElement("option");
        option.value = district.name;
        option.textContent = district.name;
        select.appendChild(option);
      });

      select.addEventListener("change", (event) => selectDistrict(event.target.value));
      resetButton.addEventListener("click", () => selectDistrict(""));
    }

    function renderSummary() {
      const seats = districts.reduce((sum, district) => sum + districtSeatNumber(district), 0);
      const singleMembers = districts.filter((district) => districtSeatNumber(district) === 1).length;
      const uncontested = districts.filter((district) => district.uncontested2023).length;
      document.getElementById("district-count").textContent = districts.length.toLocaleString("ja-JP");
      document.getElementById("seat-count").textContent = seats.toLocaleString("ja-JP");
      document.getElementById("single-member-count").textContent = singleMembers.toLocaleString("ja-JP");
      document.getElementById("uncontested-count").textContent = uncontested.toLocaleString("ja-JP");
    }

    populateControls();
    renderSummary();
    renderMap();
    selectDistrict(selectedDistrict);
    window.addEventListener("resize", renderMap);
  </script>
</body>
</html>
"""

    return (
        template.replace("__TOPO_JSON__", topo_json)
        .replace("__RECORDS_JSON__", records_json)
        .replace("__AREA_MAP_JSON__", area_map_json)
        .replace("__HEADER_NOTE__", header_note_html)
        .replace("__TRAVEL_COST_SECTION__", travel_cost_section)
    )


def main() -> None:
    topo = json.loads(TOPO_PATH.read_text(encoding="utf-8"))
    area_names = load_area_names(topo)
    candidate_photos = load_candidate_photos()
    records = parse_excel(candidate_photos)
    header_note = load_header_note()
    travel_cost_image = image_data_uri(TRAVEL_COST_IMAGE_PATH)
    area_to_district = attach_areas(records, area_names)
    html = build_html(topo, records, area_to_district, header_note, travel_cost_image)
    OUTPUT_PATH.write_text(html, encoding="utf-8")

    unmapped_areas = sorted(set(area_names) - set(area_to_district))
    print(f"wrote: {OUTPUT_PATH}")
    print(f"districts: {len(records)}")
    print(f"mapped areas: {len(area_to_district)} / {len(area_names)}")
    if unmapped_areas:
        print("unmapped areas:", ", ".join(unmapped_areas))


if __name__ == "__main__":
    main()
