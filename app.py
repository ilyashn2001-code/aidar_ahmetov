from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Tuple

from flask import Flask, render_template, request

app = Flask(__name__)


# -----------------------------
# Нормативные пороги (DEMO)
# Потом заменишь на реальные ГОСТ/РД.
# -----------------------------

@dataclass(frozen=True)
class Thresholds:
    # Для параметров "чем меньше, тем лучше" (влага, кислотное, tgδ)
    warn_max: float
    crit_max: float


@dataclass(frozen=True)
class ThresholdsMin:
    # Для параметров "чем больше, тем лучше" (пробивное напряжение)
    warn_min: float
    crit_min: float


@dataclass(frozen=True)
class ThresholdsMin2:
    # Для температуры вспышки "чем больше, тем лучше"
    warn_min: float
    crit_min: float


# Демонстрационные пороги:
# - Влага: ppm
# - Пробивное напряжение: kV
# - Кислотное число: mgKOH/g
# - tgδ: %
# - Температура вспышки: °C
TH = {
    "moisture_ppm": Thresholds(warn_max=35.0, crit_max=55.0),
    "acid_mgkoh_g": Thresholds(warn_max=0.10, crit_max=0.20),
    "tgdelta_pct": Thresholds(warn_max=0.50, crit_max=1.00),
    "bdv_kv": ThresholdsMin(warn_min=40.0, crit_min=30.0),
    "flash_c": ThresholdsMin2(warn_min=140.0, crit_min=135.0),
}

# Вес параметров для индекса риска 0..100
WEIGHTS = {
    "moisture_ppm": 0.25,
    "bdv_kv": 0.25,
    "acid_mgkoh_g": 0.15,
    "tgdelta_pct": 0.15,
    "flash_c": 0.10,
    "impurities": 0.05,
    "water_extract": 0.05,
}


# -----------------------------
# Логика оценки
# -----------------------------

def zone_max(value: float, t: Thresholds) -> Tuple[str, float]:
    """
    Чем больше значение, тем хуже.
    Возвращает (zone, score), где score: 0.0 (норма), 0.5 (предупр.), 1.0 (критично)
    """
    if value <= t.warn_max:
        return "Норма", 0.0
    if value <= t.crit_max:
        return "Предупреждение", 0.5
    return "Критично", 1.0


def zone_min(value: float, t: ThresholdsMin) -> Tuple[str, float]:
    """
    Чем меньше значение, тем хуже (для пробивного).
    """
    if value >= t.warn_min:
        return "Норма", 0.0
    if value >= t.crit_min:
        return "Предупреждение", 0.5
    return "Критично", 1.0


def zone_min2(value: float, t: ThresholdsMin2) -> Tuple[str, float]:
    """
    Температура вспышки: чем меньше, тем хуже.
    """
    if value >= t.warn_min:
        return "Норма", 0.0
    if value >= t.crit_min:
        return "Предупреждение", 0.5
    return "Критично", 1.0


def zone_impurities(value: str) -> Tuple[str, float]:
    # no / trace / yes
    v = (value or "").strip().lower()
    if v == "нет":
        return "Норма", 0.0
    if v in {"следы", "незначительно"}:
        return "Предупреждение", 0.5
    return "Критично", 1.0


def zone_water_extract(value: str) -> Tuple[str, float]:
    # neutral / slightly_acidic / acidic
    v = (value or "").strip().lower()
    if v == "нейтральная":
        return "Норма", 0.0
    if v in {"слабокислая", "слабокисл."}:
        return "Предупреждение", 0.5
    return "Критично", 1.0


def compute_index(scores: Dict[str, float]) -> float:
    # Взвешенная сумма, нормированная к 0..100
    total = 0.0
    wsum = 0.0
    for k, w in WEIGHTS.items():
        if k in scores:
            total += scores[k] * w
            wsum += w
    if wsum == 0:
        return 0.0
    return round((total / wsum) * 100.0, 1)


def overall_status(index: float, any_critical: bool) -> str:
    # Простая и понятная градация
    if any_critical or index >= 60:
        return "КРИТИЧЕСКОЕ"
    if index >= 25:
        return "ПОГРАНИЧНОЕ"
    return "НОРМА"


def build_recommendations(rows: List[dict]) -> List[str]:
    rec: List[str] = []

    # Правила — простые и объяснимые (потом расширим)
    by_name = {r["name"]: r for r in rows}

    moist = by_name["Влагосодержание (ppm)"]["zone"]
    bdv = by_name["Пробивное напряжение (кВ)"]["zone"]
    acid = by_name["Кислотное число (мг КОН/г)"]["zone"]
    tg = by_name["tg δ при 90°C (%)"]["zone"]
    imp = by_name["Механические примеси"]["zone"]
    wext = by_name["Реакция водной вытяжки"]["zone"]
    flash = by_name["Температура вспышки (°C)"]["zone"]

    if moist == "Критично" or bdv == "Критично":
        rec.append("Срочно выполнить сушку/дегазацию и фильтрацию масла (приоритет: высокий).")
        rec.append("Проверить герметичность трансформатора и состояние дыхательной системы/адсорбента.")
        rec.append("Повторить измерение пробивного напряжения после обработки масла.")

    if imp == "Критично":
        rec.append("Выполнить фильтрацию для удаления механических примесей; оценить возможный источник загрязнения.")

    if acid in {"Предупреждение", "Критично"} or tg in {"Предупреждение", "Критично"} or wext in {"Предупреждение", "Критично"}:
        rec.append("Рассмотреть регенерацию масла или частичную/полную замену при подтверждении устойчивого роста параметров старения.")
        rec.append("Рекомендуется контроль повторной пробы в ближайшие 2–4 недели (или по регламенту предприятия).")

    if flash in {"Предупреждение", "Критично"}:
        rec.append("Проверить пожаробезопасные характеристики масла; при снижении ниже нормы рассмотреть замену.")

    if not rec:
        rec.append("Параметры в пределах нормы. Продолжать эксплуатацию согласно регламенту, выполнять плановый контроль.")

    return rec


# -----------------------------
# Роуты
# -----------------------------

@app.get("/")
def index():
    return render_template("index.html")


@app.post("/evaluate")
def evaluate():
    # Без излишней магии: читаем как float, валидируем
    def f(name: str) -> float:
        raw = (request.form.get(name) or "").replace(",", ".").strip()
        return float(raw)

    transformer_id = (request.form.get("transformer_id") or "").strip()
    sample_date = (request.form.get("sample_date") or "").strip()

    moisture_ppm = f("moisture_ppm")
    bdv_kv = f("bdv_kv")
    acid = f("acid_mgkoh_g")
    tg = f("tgdelta_pct")
    flash = f("flash_c")

    impurities = (request.form.get("impurities") or "нет").strip()
    water_extract = (request.form.get("water_extract") or "нейтральная").strip()

    rows: List[dict] = []
    scores: Dict[str, float] = {}
    any_critical = False

    z, s = zone_max(moisture_ppm, TH["moisture_ppm"])
    rows.append({
        "name": "Влагосодержание (ppm)",
        "value": moisture_ppm,
        "unit": "ppm",
        "zone": z,
        "comment": "Рост влаги снижает электрическую прочность и ускоряет старение изоляции."
    })
    scores["moisture_ppm"] = s
    any_critical = any_critical or (z == "Критично")

    z, s = zone_min(bdv_kv, TH["bdv_kv"])
    rows.append({
        "name": "Пробивное напряжение (кВ)",
        "value": bdv_kv,
        "unit": "кВ",
        "zone": z,
        "comment": "Низкое пробивное напряжение повышает риск пробоя изоляции."
    })
    scores["bdv_kv"] = s
    any_critical = any_critical or (z == "Критично")

    z, s = zone_max(acid, TH["acid_mgkoh_g"])
    rows.append({
        "name": "Кислотное число (мг КОН/г)",
        "value": acid,
        "unit": "мг КОН/г",
        "zone": z,
        "comment": "Рост кислотного числа отражает старение масла и накопление продуктов окисления."
    })
    scores["acid_mgkoh_g"] = s
    any_critical = any_critical or (z == "Критично")

    z, s = zone_max(tg, TH["tgdelta_pct"])
    rows.append({
        "name": "tg δ при 90°C (%)",
        "value": tg,
        "unit": "%",
        "zone": z,
        "comment": "Повышенный tg δ указывает на рост диэлектрических потерь и загрязнение/старение."
    })
    scores["tgdelta_pct"] = s
    any_critical = any_critical or (z == "Критично")

    z, s = zone_min2(flash, TH["flash_c"])
    rows.append({
        "name": "Температура вспышки (°C)",
        "value": flash,
        "unit": "°C",
        "zone": z,
        "comment": "Снижение температуры вспышки ухудшает пожаробезопасность."
    })
    scores["flash_c"] = s
    any_critical = any_critical or (z == "Критично")

    z, s = zone_impurities(impurities)
    rows.append({
        "name": "Механические примеси",
        "value": impurities,
        "unit": "",
        "zone": z,
        "comment": "Примеси снижают электрическую прочность и ускоряют старение масла."
    })
    scores["impurities"] = s
    any_critical = any_critical or (z == "Критично")

    z, s = zone_water_extract(water_extract)
    rows.append({
        "name": "Реакция водной вытяжки",
        "value": water_extract,
        "unit": "",
        "zone": z,
        "comment": "Кислая реакция может указывать на наличие кислотных продуктов старения."
    })
    scores["water_extract"] = s
    any_critical = any_critical or (z == "Критично")

    index_score = compute_index(scores)
    status = overall_status(index_score, any_critical)
    recs = build_recommendations(rows)

    return render_template(
        "result.html",
        transformer_id=transformer_id,
        sample_date=sample_date,
        status=status,
        index_score=index_score,
        rows=rows,
        recommendations=recs,
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
