from __future__ import annotations

import html
import json
import re
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parent
EXCEL_PATH = ROOT / "2026福岡県議一覧.xlsx"
TOPO_PATH = ROOT / "fukuoka_city_20230101.topojson"
OUTPUT_PATH = ROOT / "fukuoka-kenmap.html"

FULLWIDTH_DIGITS = str.maketrans("０１２３４５６７８９", "0123456789")
BY_ELECTION_NAMES = {"佐藤かえで", "吉岡れい子", "うらただいじ", "吉松源明", "亀崎大介", "となり祥平"}


def clean_cell(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    text = text.replace("\u3000", " ")
    return re.sub(r"\s+", " ", text)


def is_travel_note(text: str) -> bool:
    normalized = text.translate(FULLWIDTH_DIGITS)
    return bool(re.search(r"\d+回", normalized) and ("万円" in normalized or "不明" in normalized))


def parse_travel_note(text: str) -> dict[str, object] | None:
    if not is_travel_note(text):
        return None
    return {"raw": text}


def parse_person_entry(text: str) -> dict[str, str]:
    parts = text.rsplit(" ", 1)
    if len(parts) == 2:
        return {"name": parts[0], "party": parts[1], "byElection": parts[0] in BY_ELECTION_NAMES}
    return {"name": text, "party": "", "byElection": text in BY_ELECTION_NAMES}


def topo_area_name(properties: dict[str, object]) -> str:
    city = properties.get("N03_003") or ""
    ward_or_town = properties.get("N03_004") or ""
    if city and ward_or_town:
        return f"{city}{ward_or_town}"
    return str(ward_or_town or city)


def load_area_names(topo: dict[str, object]) -> list[str]:
    geometries = topo["objects"]["city"]["geometries"]
    return [topo_area_name(g["properties"]) for g in geometries]


def parse_excel() -> list[dict[str, object]]:
    df = pd.read_excel(EXCEL_PATH, header=None, dtype=object)
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
                    entries.append({"kind": "person", "text": value, "name": person["name"], "party": person["party"], "byElection": person["byElection"]})
            slots.append({"entries": entries})

        normalized.append(
            {
                "name": record["name"],
                "seats": record["seats"],
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
) -> str:
    topo_json = json_for_script(topo)
    records_json = json_for_script(records)
    area_map_json = json_for_script(area_to_district)
    header_note_html = html.escape(header_note)

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
      width: min(1240px, calc(100% - 32px));
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
      grid-template-columns: minmax(0, 1.45fr) minmax(320px, 0.8fr);
      gap: 18px;
      align-items: start;
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
      opacity: 0.42;
    }

    .area.is-selected {
      stroke: var(--strong);
      stroke-width: 2.25;
      opacity: 1;
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
      margin: 0 0 12px;
      color: var(--muted);
      font-weight: 500;
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
      grid-template-columns: minmax(68px, auto) 1fr;
      gap: 10px;
      align-items: center;
      min-width: 0;
    }

    .person-row + .person-row,
    .travel + .person-row {
      border-top: 1px solid var(--line);
      padding-top: 8px;
      margin-top: 8px;
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
      word-break: keep-all;
      overflow-wrap: anywhere;
    }
    .by-election {
      display: inline-flex;
      align-items: center;
      margin-left: 8px;
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
      margin: 5px 0 0 calc(68px + 10px);
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
      }
      .detail {
        position: static;
      }
    }

    @media (max-width: 560px) {
      .app {
        width: min(100% - 20px, 1240px);
        padding-top: 18px;
      }
      .map-panel,
      .detail {
        padding: 12px;
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
        <span class="note">__HEADER_NOTE__</span>
      </div>
    </header>

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

        <div class="map-wrap" id="map-wrap">
          <svg id="map" role="img" aria-labelledby="map-title map-desc" viewBox="0 0 920 780">
            <title id="map-title">福岡県の市区町村・行政区別の県議会選挙区地図</title>
            <desc id="map-desc">各区域をクリックすると、右側に対応する選挙区の議員情報が表示されます。</desc>
            <g id="area-layer"></g>
            <g id="label-layer" class="label-layer"></g>
          </svg>
          <div id="tooltip" class="tooltip" role="status"></div>
        </div>
      </section>

      <aside class="detail" id="detail" aria-live="polite"></aside>
    </main>

    <footer>
      議員データ: 2026福岡県議一覧.xlsx / 境界データ: <a href="https://geoshape.ex.nii.ac.jp/city/choropleth/40_city.html">Geoshape 市区町村TopoJSON 2023-01-01</a>（CODH作成）
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


    let selectedDistrict = "福岡市東区";

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

    function fillFor(feature) {
      const record = byDistrict.get(feature.properties.districtName);
      return record ? "var(--map-low)" : "var(--map-empty)";
    }

    function renderMap() {
      const width = 920;
      const height = window.matchMedia("(max-width: 560px)").matches ? 1060 : 780;
      svg.attr("viewBox", `0 0 ${width} ${height}`);

      const projection = d3.geoMercator().fitExtent([[18, 18], [width - 18, height - 18]], {
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
          const record = byDistrict.get(feature.properties.districtName);
          const lines = [
            `<strong>${escapeHtml(feature.properties.areaName)}</strong>`,
            record ? escapeHtml(record.name) : "表データなし",
          ];
          tooltip
            .style("display", "block")
            .style("left", `${event.offsetX + 12}px`)
            .style("top", `${event.offsetY + 12}px`)
            .html(lines.join("<br>"));
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

      drawSelectedLabels(path);
    }

    function areaClass(feature) {
      const classes = ["area"];
      if (selectedDistrict) {
        if (feature.properties.districtName === selectedDistrict) {
          classes.push("is-selected");
        } else {
          classes.push("is-muted");
        }
      }
      return classes.join(" ");
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
      const width = 920;
      const height = window.matchMedia("(max-width: 560px)").matches ? 1060 : 780;
      const projection = d3.geoMercator().fitExtent([[18, 18], [width - 18, height - 18]], {
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
        <p class="sub">地図上の区域をクリックすると、該当する選挙区の議員情報がここに表示されます。</p>
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

      const members = record.slots.map((slot) => {
        const entries = slot.entries.map((entry) => {
          if (entry.kind === "travel") {
            return `<div class="travel">${escapeHtml(entry.text)}</div>`;
          }
          const party = entry.party ? `<span class="party">${escapeHtml(entry.party)}</span>` : "";
          const badge = entry.byElection ? `<span class="by-election">2024・2025年補選</span>` : "";
          const name = escapeHtml(entry.name || entry.text);
          return `<div class="person-row">${party}<span class="person-name">${name}${badge}</span></div>`;
        }).join("");
        return `<li class="member">${entries}</li>`;
      }).join("");

      detail.innerHTML = `
        <h2>${escapeHtml(record.name)}</h2>
        <p class="sub">${escapeHtml(record.areas.join("、"))}</p>
        <p class="seat-line">定数 ${escapeHtml(record.seats || "-")}</p>
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
      const seats = districts.reduce((sum, district) => sum + seatNumber(district.seats), 0);
      document.getElementById("district-count").textContent = districts.length.toLocaleString("ja-JP");
      document.getElementById("seat-count").textContent = seats.toLocaleString("ja-JP");
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
    )


def main() -> None:
    topo = json.loads(TOPO_PATH.read_text(encoding="utf-8"))
    area_names = load_area_names(topo)
    records = parse_excel()
    header_note = load_header_note()
    area_to_district = attach_areas(records, area_names)
    html = build_html(topo, records, area_to_district, header_note)
    OUTPUT_PATH.write_text(html, encoding="utf-8")

    unmapped_areas = sorted(set(area_names) - set(area_to_district))
    print(f"wrote: {OUTPUT_PATH}")
    print(f"districts: {len(records)}")
    print(f"mapped areas: {len(area_to_district)} / {len(area_names)}")
    if unmapped_areas:
        print("unmapped areas:", ", ".join(unmapped_areas))


if __name__ == "__main__":
    main()
