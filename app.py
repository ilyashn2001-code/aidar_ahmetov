from __future__ import annotations

from io import BytesIO
from typing import Dict, List, Tuple

from flask import Flask, render_template, request, session, send_file
from openpyxl import Workbook


app = Flask(__name__)
app.secret_key = "dev-secret-key"  # для MVP достаточно


# --------- простая оценка (заглушка на время) ----------
def assess(value: float, warn: float, crit: float, higher_is_worse: bool = True) -> Tuple[str, float]:
    """
    Возвращает: (зона, штраф 0..1)
    """
    if higher_is_worse:
        if value <= warn:
            return "Норма", 0.0
        if value <= crit:
            return "Предупреждение", 0.5
        return "Критично", 1.0
    else:
        if value >= warn:
            return "Норма", 0.0
        if value >= crit:
            return "Предупреждение", 0.5
        return "Критично", 1.0


def compute_index(penalties: Dict[str, float]) -> float:
    # веса для MVP
    weights = {
        "moisture": 0.25,
        "bdv": 0.25,
        "acid": 0.15,
        "tg": 0.15,
        "flash": 0.10,
        "imp": 0.05,
        "wext": 0.05,
    }
    total = 0.0
    wsum = 0.0
    for k, w in weights.items():
        if k in penalties:
            total += penalties[k] * w
            wsum += w
    if wsum == 0:
        return 0.0
    return round((total / wsum) * 100.0, 1)


def status_from(index_score: float, any_critical: bool) -> str:
    if any_critical or index_score >= 60:
        return "КРИТИЧЕСКОЕ"
    if index_score >= 25:
        return "ПОГРАНИЧНОЕ"
    return "НОРМА"


def excel_from_result(result: dict) -> BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "Результат"

    ws["A1"] = "Заключение по пробе трансформаторного масла"
    ws["A3"] = "Трансформатор"
    ws["B3"] = result.get("transformer_id", "")
    ws["A4"] = "Дата пробы"
    ws["B4"] = result.get("sample_date", "")

    ws["A6"] = "Индекс"
    ws["B6"] = f'{result.get("index_score", 0)} / 100'
    ws["A7"] = "Состояние"
    ws["B7"] = result.get("status", "")

    ws["A9"] = "Показатель"
    ws["B9"] = "Значение"
    ws["C9"] = "Оценка"

    row = 10
    for r in result.get("rows", []):
        ws[f"A{row}"] = r["name"]
        ws[f"B{row}"] = r["value"]
        ws[f"C{row}"] = r["zone"]
        row += 1

    row += 2
    ws[f"A{row}"] = "Принадлежит Казанскому Государственному Энергетическому Университету. Создал Ахметов Айдар Русланович."

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


# ------------------- routes -------------------

@app.get("/")
def index():
    return render_template("index.html")


@app.post("/evaluate")
def evaluate():
    def get_float(name: str) -> float:
        raw = (request.form.get(name) or "").replace(",", ".").strip()
        return float(raw)

    transformer_id = (request.form.get("transformer_id") or "").strip()
    sample_date = (request.form.get("sample_date") or "").strip()

    moisture = get_float("moisture_ppm")
    bdv = get_float("bdv_kv")
    acid = get_float("acid_mgkoh_g")
    tg = get_float("tgdelta_pct")
    flash = get_float("flash_c")

    impurities = (request.form.get("impurities") or "нет").strip()
    water_extract = (request.form.get("water_extract") or "нейтральная").strip()

    rows: List[dict] = []
    penalties: Dict[str, float] = {}
    any_critical = False

    # пороги MVP (позже заменим на ГОСТ-профили)
    z, p = assess(moisture, warn=35, crit=55, higher_is_worse=True)
    rows.append({"name": "Влагосодержание, ppm", "value": moisture, "zone": z})
    penalties["moisture"] = p
    any_critical = any_critical or (z == "Критично")

    z, p = assess(bdv, warn=40, crit=30, higher_is_worse=False)
    rows.append({"name": "Пробивное напряжение, кВ", "value": bdv, "zone": z})
    penalties["bdv"] = p
    any_critical = any_critical or (z == "Критично")

    z, p = assess(acid, warn=0.10, crit=0.20, higher_is_worse=True)
    rows.append({"name": "Кислотное число, мг КОН/г", "value": acid, "zone": z})
    penalties["acid"] = p
    any_critical = any_critical or (z == "Критично")

    z, p = assess(tg, warn=0.50, crit=1.00, higher_is_worse=True)
    rows.append({"name": "tg δ при 90°C, %", "value": tg, "zone": z})
    penalties["tg"] = p
    any_critical = any_critical or (z == "Критично")

    z, p = assess(flash, warn=140, crit=135, higher_is_worse=False)
    rows.append({"name": "Температура вспышки, °C", "value": flash, "zone": z})
    penalties["flash"] = p
    any_critical = any_critical or (z == "Критично")

    # простая логика по строкам
    imp_zone = "Норма" if impurities.lower() == "нет" else ("Предупреждение" if impurities.lower() == "следы" else "Критично")
    wext_zone = "Норма" if water_extract.lower() == "нейтральная" else ("Предупреждение" if "слаб" in water_extract.lower() else "Критично")

    rows.append({"name": "Механические примеси", "value": impurities, "zone": imp_zone})
    rows.append({"name": "Реакция водной вытяжки", "value": water_extract, "zone": wext_zone})

    penalties["imp"] = 0.0 if imp_zone == "Норма" else (0.5 if imp_zone == "Предупреждение" else 1.0)
    penalties["wext"] = 0.0 if wext_zone == "Норма" else (0.5 if wext_zone == "Предупреждение" else 1.0)
    any_critical = any_critical or (imp_zone == "Критично") or (wext_zone == "Критично")

    index_score = compute_index(penalties)
    status = status_from(index_score, any_critical)

    result = {
        "transformer_id": transformer_id,
        "sample_date": sample_date,
        "index_score": index_score,
        "status": status,
        "rows": rows,
    }
    session["last_result"] = result

    return render_template(
        "result.html",
        transformer_id=transformer_id,
        sample_date=sample_date,
        index_score=index_score,
        status=status,
        rows=rows,
    )


@app.get("/export/xlsx")
def export_xlsx():
    result = session.get("last_result")
    if not result:
        return "Нет данных для экспорта. Сначала выполните расчет.", 400

    bio = excel_from_result(result)
    return send_file(
        bio,
        as_attachment=True,
        download_name="zaklyuchenie_transformatornoe_maslo.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
