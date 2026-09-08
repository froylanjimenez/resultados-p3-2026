"""
Pipeline principal: cruza el listado oficial de estudiantes con los
resultados crudos de ZipGrade (sesion 1 y 2, los 6 grados), calcula
puntajes por area/sesion usando answer_keys.json, y escribe los JSON
que consume el sitio estatico (data/estudiantes.json, data/promedios.json,
data/meta.json).

Solo escribe a la salida: nombre, matricula, grado, grupo y resultados
calculados. Telefono, direccion y numero de documento del listado NUNCA
se copian a la salida.

Uso:
    python3 build/extract_answer_keys.py   # primero, si answer_keys.json no existe o cambio la fuente
    python3 build/build_data.py
"""
import csv
import io
import json
import re
import unicodedata
import zipfile
from datetime import datetime
from pathlib import Path

import openpyxl

BASE = Path("/home/froylan/Documents/Pruebas periodicas CESUM/Tercer periodo 2026")
ROSTER_XLSX = BASE / "LISTADO DE 6º A 11º.xlsx"
ZIP_S1 = BASE / "Resultados_sesion_1.zip"
ZIP_S2 = BASE / "Resultados_sesion2.zip"

BUILD_DIR = Path(__file__).parent
ANSWER_KEYS_PATH = BUILD_DIR / "answer_keys.json"
OVERRIDES_PATH = BUILD_DIR / "overrides.json"
DATA_DIR = BUILD_DIR.parent / "data"
UNMATCHED_REPORT = BUILD_DIR / "unmatched_report.txt"

GRADE_NAME_MAP = {
    "SEXTO": "6",
    "SEPTIMO": "7",
    "OCTAVO": "8",
    "NOVENO": "9",
    "DECIMO": "10",
    "ONCE": "11",
    "UNDECIMO": "11",
}

# prefijos (en minuscula) de los nombres de archivo CSV de ZipGrade por grado,
# para cada sesion (no dependemos de las marcas de tiempo del nombre completo).
ZIPGRADE_PREFIXES = {
    "session1": {
        "6": "sexto",
        "7": "septimo",
        "8": "octavo",
        "9": "noveno",
        "10": "decimo",
        "11": "once",
    },
    "session2": {
        "6": "sexto",
        "7": "septimo",
        "8": "octavo",
        "9": "noveno",
        "10": "decimo",
        "11": "undecimo",
    },
}


def strip_accents(s):
    return "".join(
        c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn"
    )


def norm(s):
    if s is None:
        return ""
    s = strip_accents(str(s)).upper()
    s = re.sub(r"\s+", " ", s).strip()
    return s


def norm_grade_label(s):
    s = strip_accents(str(s or "")).upper().strip()
    return GRADE_NAME_MAP.get(s)


# ---------------------------------------------------------------------------
# 1. Listado oficial (roster)
# ---------------------------------------------------------------------------
def load_roster():
    """Devuelve:
    roster_by_key: {(grado, matricula): {"nombre":..., "grado":..., "grupo":...}}
    roster_by_norm_name: {grado: {nombre_normalizado: record}}
    grupos_por_grado: {grado: [lista de grupos ordenada]}
    """
    wb = openpyxl.load_workbook(ROSTER_XLSX, data_only=True)
    roster_by_key = {}
    roster_by_norm_name = {g: {} for g in GRADE_NAME_MAP.values()}
    grupos_por_grado = {g: set() for g in set(GRADE_NAME_MAP.values())}

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        grado_label = ws["H3"].value
        grado = norm_grade_label(grado_label)
        if grado is None:
            print(f"AVISO: hoja {sheet_name} sin grado reconocible ({grado_label!r}), se omite")
            continue

        grupo_cell = ws["K1"].value or ""
        m = re.search(r"Grupo:\s*(\S+)", str(grupo_cell))
        if not m:
            print(f"AVISO: hoja {sheet_name} sin 'Grupo:' reconocible, se omite")
            continue
        letra = m.group(1).strip()[-1].upper()
        grupo_id = f"{grado}{letra}"
        grupos_por_grado[grado].add(grupo_id)

        r = 5
        while True:
            apellidos_nombres = ws.cell(row=r, column=2).value
            if apellidos_nombres is None or str(apellidos_nombres).strip() == "":
                break
            matricula = ws.cell(row=r, column=5).value
            matricula = str(matricula).strip() if matricula is not None else None
            nombre_fmt = str(apellidos_nombres).strip().title()

            record = {"nombre": nombre_fmt, "grado": grado, "grupo": grupo_id}
            if matricula:
                roster_by_key[(grado, matricula)] = record
            roster_by_norm_name[grado][norm(apellidos_nombres)] = {
                **record,
                "matricula": matricula,
            }
            r += 1

    grupos_por_grado = {g: sorted(v) for g, v in grupos_por_grado.items()}
    return roster_by_key, roster_by_norm_name, grupos_por_grado


# ---------------------------------------------------------------------------
# 2. Overrides manuales
# ---------------------------------------------------------------------------
def load_overrides():
    raw = json.loads(OVERRIDES_PATH.read_text(encoding="utf-8"))
    by_id = {}
    for entry in raw["by_id"]:
        key = (entry["grado"], entry["session"], str(entry["zipgrade_id"]))
        by_id[key] = entry["matricula"]
    return by_id


# ---------------------------------------------------------------------------
# 3. Resultados crudos de ZipGrade
# ---------------------------------------------------------------------------
def find_zip_member(zf, grade, session):
    prefix = ZIPGRADE_PREFIXES[session][grade]
    for name in zf.namelist():
        base = name.rsplit("/", 1)[-1]
        if base.lower().startswith(prefix.lower()) and base.lower().endswith(".csv"):
            return name
    raise FileNotFoundError(f"No se encontro CSV de ZipGrade para grado {grade} / {session}")


def load_zipgrade_csv(zip_path, grade, session):
    with zipfile.ZipFile(zip_path) as zf:
        member = find_zip_member(zf, grade, session)
        raw = zf.read(member).decode("utf-8-sig")
    return list(csv.DictReader(io.StringIO(raw)))


# ---------------------------------------------------------------------------
# 4. Cruce + calificacion
# ---------------------------------------------------------------------------
def score_student_session(row, areas, n_questions, key):
    """A partir de una fila de ZipGrade (dict) y la lista de areas
    [{area,start,end}], calcula el detalle por area."""
    points = {}
    for q in range(1, n_questions + 1):
        col = f"#{q} Points Earned"
        val = row.get(col)
        try:
            points[q] = float(val) if val not in (None, "") else 0.0
        except ValueError:
            points[q] = 0.0

    areas_out = []
    total_correct = 0.0
    for a in areas:
        qs = range(a["start"], a["end"] + 1)
        correct = sum(points[q] for q in qs)
        total = a["end"] - a["start"] + 1
        total_correct += correct
        preguntas = [
            {"n": q, "correcta": points[q] >= 0.999, "clave": key.get(str(q))}
            for q in qs
        ]
        areas_out.append(
            {
                "area": a["area"],
                "correctas": correct,
                "total": total,
                "porcentaje": round(correct / total * 100, 1) if total else 0.0,
                "preguntas": preguntas,
            }
        )
    pct_total = round(total_correct / n_questions * 100, 1) if n_questions else 0.0
    return areas_out, pct_total


def main():
    answer_keys = json.loads(ANSWER_KEYS_PATH.read_text(encoding="utf-8"))
    overrides_by_id = load_overrides()
    roster_by_key, roster_by_norm_name, grupos_por_grado = load_roster()

    estudiantes = {}  # (grado, matricula) -> record
    unmatched_lines = []
    stats = {}

    for grade in answer_keys.keys():
        for session, zip_path in (("session1", ZIP_S1), ("session2", ZIP_S2)):
            areas = answer_keys[grade][session]["areas"]
            n_q = answer_keys[grade][session]["total"]
            key = answer_keys[grade][session]["key"]
            rows = load_zipgrade_csv(zip_path, grade, session)

            matched = 0
            unmatched = 0
            for row in rows:
                first = (row.get("Student First Name") or "").strip()
                last = (row.get("Student Last Name") or "").strip()
                zid = (row.get("Student ID") or "").strip()

                matricula = overrides_by_id.get((grade, session, zid))
                roster_rec = None
                if matricula:
                    roster_rec = roster_by_key.get((grade, matricula))
                elif first or last:
                    name_key = norm(f"{last} {first}")
                    roster_rec = roster_by_norm_name[grade].get(name_key)
                    if roster_rec:
                        matricula = roster_rec.get("matricula")

                if not roster_rec:
                    unmatched += 1
                    unmatched_lines.append(
                        f"grado={grade} session={session} zid={zid!r} "
                        f"nombre={first!r} {last!r} -> SIN COINCIDENCIA"
                    )
                    continue

                matched += 1
                areas_out, pct_total = score_student_session(row, areas, n_q, key)

                sid = matricula or f"{grade}-{zid}"
                student_key = (grade, sid)
                if student_key not in estudiantes:
                    estudiantes[student_key] = {
                        "id": sid,
                        "nombre": roster_rec["nombre"],
                        "grado": grade,
                        "grupo": roster_rec["grupo"],
                        "sesiones": {},
                    }
                estudiantes[student_key]["sesiones"][session] = {
                    "areas": areas_out,
                    "porcentaje_total": pct_total,
                }

            stats[(grade, session)] = {
                "total_filas": len(rows),
                "matched": matched,
                "unmatched": unmatched,
            }
            print(
                f"grado {grade:>2} {session}: {matched}/{len(rows)} unidos "
                f"({unmatched} sin coincidencia)"
            )

    # overall pct combinando ambas sesiones si existen
    for rec in estudiantes.values():
        s1 = rec["sesiones"].get("session1")
        s2 = rec["sesiones"].get("session2")
        if s1 and s2:
            c1 = sum(a["correctas"] for a in s1["areas"])
            t1 = sum(a["total"] for a in s1["areas"])
            c2 = sum(a["correctas"] for a in s2["areas"])
            t2 = sum(a["total"] for a in s2["areas"])
            rec["porcentaje_general"] = round((c1 + c2) / (t1 + t2) * 100, 1) if (t1 + t2) else 0.0
        elif s1:
            rec["porcentaje_general"] = s1["porcentaje_total"]
        elif s2:
            rec["porcentaje_general"] = s2["porcentaje_total"]
        else:
            rec["porcentaje_general"] = None

    UNMATCHED_REPORT.write_text("\n".join(unmatched_lines) + "\n", encoding="utf-8")
    print(f"\nReporte de no coincidencias: {UNMATCHED_REPORT} ({len(unmatched_lines)} filas)")

    # ---- agregados ----
    by_grado_area = {}  # grado -> session -> area -> [pcts]
    by_grado = {}  # grado -> [pcts generales]
    by_grupo = {}  # grupo -> [pcts generales]
    colegio_pcts = []
    colegio_por_sesion = {"session1": [], "session2": []}

    for rec in estudiantes.values():
        g = rec["grado"]
        if rec["porcentaje_general"] is not None:
            by_grado.setdefault(g, []).append(rec["porcentaje_general"])
            by_grupo.setdefault(rec["grupo"], []).append(rec["porcentaje_general"])
            colegio_pcts.append(rec["porcentaje_general"])
        for session, sdata in rec["sesiones"].items():
            colegio_por_sesion[session].append(sdata["porcentaje_total"])
            area_map = by_grado_area.setdefault(g, {}).setdefault(session, {})
            for a in sdata["areas"]:
                area_map.setdefault(a["area"], []).append(a["porcentaje"])

    def avg(lst):
        return round(sum(lst) / len(lst), 1) if lst else None

    promedios = {
        "colegio": {
            "promedio_general": avg(colegio_pcts),
            "por_sesion": {k: avg(v) for k, v in colegio_por_sesion.items()},
            "n_estudiantes": len(estudiantes),
        },
        "por_grado": {g: avg(v) for g, v in by_grado.items()},
        "por_grupo": {g: avg(v) for g, v in by_grupo.items()},
        "por_area": {
            g: {s: {a: avg(v) for a, v in amap.items()} for s, amap in sessions.items()}
            for g, sessions in by_grado_area.items()
        },
    }

    meta = {
        "generado": datetime.now().isoformat(timespec="seconds"),
        "grados": sorted(answer_keys.keys(), key=int),
        "grupos_por_grado": grupos_por_grado,
        "areas_por_grado": {
            g: {
                "session1": [a["area"] for a in answer_keys[g]["session1"]["areas"]],
                "session2": [a["area"] for a in answer_keys[g]["session2"]["areas"]],
            }
            for g in answer_keys
        },
    }

    DATA_DIR.mkdir(exist_ok=True)
    (DATA_DIR / "estudiantes.json").write_text(
        json.dumps(list(estudiantes.values()), ensure_ascii=False), encoding="utf-8"
    )
    (DATA_DIR / "promedios.json").write_text(
        json.dumps(promedios, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (DATA_DIR / "meta.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"\nEscritos: {DATA_DIR}/estudiantes.json, promedios.json, meta.json")
    print(f"Total estudiantes con al menos 1 sesion: {len(estudiantes)}")


if __name__ == "__main__":
    main()
