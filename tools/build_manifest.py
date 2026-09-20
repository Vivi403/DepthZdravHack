"""
Сборка единого training-манифеста из двух источников разметки:

1. routing_labels.csv  -- какой файл к какой анатомической области относится
   (study_uid, relative_path, region), заполняется вручную по contact sheet'ам.
2. разметка.xlsx        -- экспертная оценка качества по (study, область):
   бинарные критерии + агрегированный quality_class, вердикт по каждой области.

На выходе -- manifest.csv с одной строкой на РАЗМЕЧЕННЫЙ файл:
    image_path, study_uid, region, quality_class, violation_type

Строки с пустым region в routing_labels.csv (неразмеченные дубли) пропускаются --
они не входят в обучающую выборку.

Запуск:
    python tools/build_manifest.py \
        data/raw/routing_labels.csv \
        data/raw/разметка.xlsx \
        data/raw/Исследования \
        data/manifest.csv
"""

from __future__ import annotations

import csv
import sys
from pathlib import Path

import openpyxl

# Соответствие между колонками разметка.xlsx и (регион -> критерии).
# Каждый критерий: (буква колонки в листе "Калибровка", человекочитаемый код нарушения)
SPINE_CRITERIA = [
    ("C", "incorrect_positioning"),  # корректная укладка
    ("D", "axis_tilt_over_5deg"),  # ось позвоночника, допустимый наклон 5°
    ("E", "foreign_object_or_artifact"),  # посторонние предметы/артефакты
]
HIP_CRITERIA = [
    ("rotation", "rotation_error"),  # позиционирование/ротация
    ("roi", "roi_incorrect"),  # корректность области интереса
]
# для бедра колонки разные для left/right, задаём отдельно ниже

REGION_TO_ITOG_COL = {
    "spine": "J",
    "hip_right": "K",
    "hip_left": "L",
}

HIP_COLUMN_LETTERS = {
    "hip_right": {"rotation": "F", "roi": "G"},
    "hip_left": {"rotation": "H", "roi": "I"},
}

COL_LETTER_TO_IDX = {chr(ord("A") + i): i for i in range(26)}  # A=0, B=1, ...


def load_quality_labels(xlsx_path: str) -> dict:
    """Возвращает {study_uid: {region: (quality_class, [violation_codes])}}."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb["Калибровка"]

    result = {}
    for row in ws.iter_rows(min_row=3, values_only=True):
        study_uid = row[1]
        if not study_uid:
            continue

        entry = {}

        # позвоночник
        spine_flags = []
        for col_letter, code in SPINE_CRITERIA:
            idx = COL_LETTER_TO_IDX[col_letter]
            if row[idx] == 1:
                spine_flags.append(code)
        spine_itog = row[COL_LETTER_TO_IDX["J"]]
        if spine_itog is not None:
            entry["spine"] = (int(spine_itog), spine_flags)

        # бедра
        for region in ("hip_right", "hip_left"):
            flags = []
            for crit_key, col_letter in HIP_COLUMN_LETTERS[region].items():
                idx = COL_LETTER_TO_IDX[col_letter]
                code = dict(HIP_CRITERIA)[crit_key]
                if row[idx] == 1:
                    flags.append(code)
            itog_col = REGION_TO_ITOG_COL[region]
            itog_val = row[COL_LETTER_TO_IDX[itog_col]]
            if itog_val is not None:
                entry[region] = (int(itog_val), flags)

        result[str(study_uid)] = entry

    return result


def build_manifest(
    routing_csv: str, xlsx_path: str, studies_root: str, out_csv: str
) -> None:
    quality_labels = load_quality_labels(xlsx_path)
    studies_root = Path(studies_root)

    manifest_rows = []
    skipped_no_region = 0
    skipped_no_quality_label = 0

    with open(routing_csv, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            region = row["region"].strip()
            if not region:
                skipped_no_region += 1
                continue

            study_uid = row["study_uid"]
            study_labels = quality_labels.get(study_uid, {})
            if region not in study_labels:
                skipped_no_quality_label += 1
                print(
                    f"[WARN] нет метки качества для study={study_uid} region={region} "
                    f"(проверьте разметка.xlsx)"
                )
                continue

            quality_class, violation_codes = study_labels[region]
            image_path = studies_root / study_uid / row["relative_path"]

            manifest_rows.append(
                {
                    "image_path": str(image_path),
                    "study_uid": study_uid,
                    "region": region,
                    "quality_class": quality_class,
                    "violation_type": ";".join(violation_codes),
                }
            )

    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=[
                "image_path",
                "study_uid",
                "region",
                "quality_class",
                "violation_type",
            ],
        )
        writer.writeheader()
        writer.writerows(manifest_rows)

    print(f"\nСобрано строк манифеста: {len(manifest_rows)}")
    print(f"Пропущено (region не заполнен): {skipped_no_region}")
    print(f"Пропущено (нет метки качества в разметка.xlsx): {skipped_no_quality_label}")

    n_violations = sum(1 for r in manifest_rows if r["quality_class"] == 1)
    print(f"Из них с нарушением качества: {n_violations} / {len(manifest_rows)}")
    print(f"\nМанифест сохранён: {out_csv}")


if __name__ == "__main__":
    if len(sys.argv) != 5:
        print(
            "Использование: python build_manifest.py "
            "<routing_labels.csv> <разметка.xlsx> <studies_root> <out_manifest.csv>"
        )
        sys.exit(1)
    build_manifest(*sys.argv[1:])
