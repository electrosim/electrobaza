# -*- coding: utf-8 -*-
"""Миграция каталога: папки тем и id → <номер>-<транслит>."""

import json
import os
import shutil

ROOT = r"D:\Electrobaza"
CAT = os.path.join(ROOT, "materials.json")

# (старый путь темы, новый путь темы)
FOLDERS = [
    ("content/groups/1me-25/mdk-05-01/benchwork-attestation", "content/groups/1me-25/mdk-05-01/1-attestaciya-slesarnoe-delo"),
    ("content/groups/1me-25/mdk-05-01/wire-connections", "content/groups/1me-25/mdk-05-01/2-vidy-soedineniy-zhil"),
    ("content/groups/1me-25/mdk-05-02/switch-socket-connection", "content/groups/1me-25/mdk-05-02/1-vyklyuchateli-i-rozetki"),
    ("content/groups/1me-25/mdk-05-02/lighting-network-check", "content/groups/1me-25/mdk-05-02/2-proverka-osvetitelnoy-seti"),
    ("content/groups/1me-25/mdk-05-02/circuit-breaker-selection", "content/groups/1me-25/mdk-05-02/3-avtomaticheskiy-vyklyuchatel"),
    ("content/groups/1eo-25/mdk-04-02/semester-1/light-sources-luminaires", "content/groups/1eo-25/mdk-04-02/semester-1/1-istochniki-sveta-i-armatura"),
    ("content/groups/1eo-25/mdk-04-02/semester-1/lighting-installation-repair", "content/groups/1eo-25/mdk-04-02/semester-1/2-montazh-i-remont-svetilnikov"),
    ("content/groups/1eo-24/mdk-04-02/semester-3/undervoltage-release-light-sensor", "content/groups/1eo-24/mdk-04-02/semester-3/1-rascepitel-i-datchik-sveta"),
    ("content/groups/1eo-24/mdk-04-02/semester-3/three-phase-meter-connection", "content/groups/1eo-24/mdk-04-02/semester-3/2-trehfaznyy-schetchik"),
    ("content/groups/1eo-24/mdk-04-02/semester-3/electrical-apparatus-faults", "content/groups/1eo-24/mdk-04-02/semester-3/3-neispravnosti-elektroapparatov"),
    ("content/groups/1eo-24/mdk-04-02/semester-3/attestation", "content/groups/1eo-24/mdk-04-02/semester-3/4-attestaciya"),
    ("content/groups/1eo-24/mdk-04-02/semester-3/pue-and-gost", "content/groups/1eo-24/mdk-04-02/semester-3/5-pue-i-gost"),
    ("content/groups/1eo-25/mdk-04-01/soldering-tinning", "content/groups/1eo-25/mdk-04-01/1-payka-i-luzhenie"),
    ("content/groups/1eo-24/op-08/information-resources", "content/groups/1eo-24/op-08/1-informacionnye-resursy"),
    ("content/groups/1eo-24/op-08/information-systems", "content/groups/1eo-24/op-08/2-informacionnye-sistemy"),
    ("content/groups/1eo-24/op-08/engineering-calculations-automation", "content/groups/1eo-24/op-08/3-avtomatizaciya-raschetov"),
    ("content/groups/1eo-24/op-08/math-software-packages", "content/groups/1eo-24/op-08/4-pakety-dlya-raschetov"),
    ("content/groups/1eo-24/op-12/constitutional-system", "content/groups/1eo-24/op-12/1-konstitucionnyy-stroy"),
    ("content/groups/1eo-24/op-12/law-sources", "content/groups/1eo-24/op-12/2-istochniki-prava"),
    ("content/groups/1eo-24/op-12/law-personality", "content/groups/1eo-24/op-12/3-pravo-i-lichnost"),
    ("content/groups/1eo-24/op-12/npa-validity", "content/groups/1eo-24/op-12/4-deystvie-npa"),
    ("content/groups/1eo-24/op-12/attestation", "content/groups/1eo-24/op-12/5-attestaciya-pravo"),
    ("content/groups/1eo-24/op-12/judicial-authority", "content/groups/1eo-24/op-12/6-sudebnaya-vlast"),
    ("content/groups/1a-25-3a-26-11/op-03/electromagnetism", "content/groups/1a-25-3a-26-11/op-03/1-elektromagnetizm"),
    ("content/groups/1a-25-3a-26-11/op-03/electromagnetic-induction", "content/groups/1a-25-3a-26-11/op-03/2-elektromagnitnaya-indukciya"),
    ("content/groups/1a-25-3a-26-11/op-03/alternating-current", "content/groups/1a-25-3a-26-11/op-03/3-peremenny-tok"),
]

# (дисциплина, старый id темы) -> новый id
TOPICS = {
    ("mdk-05-01", "benchwork-attestation"): "1-attestaciya-slesarnoe-delo",
    ("mdk-05-01", "wire-connections"): "2-vidy-soedineniy-zhil",
    ("mdk-05-02", "switch-socket-connection"): "1-vyklyuchateli-i-rozetki",
    ("mdk-05-02", "lighting-network-check"): "2-proverka-osvetitelnoy-seti",
    ("mdk-05-02", "circuit-breaker-selection"): "3-avtomaticheskiy-vyklyuchatel",
    ("mdk-04-02-1eo-25-s1", "light-sources-luminaires"): "1-istochniki-sveta-i-armatura",
    ("mdk-04-02-1eo-25-s1", "lighting-installation-repair"): "2-montazh-i-remont-svetilnikov",
    ("mdk-04-02-1eo-24-s3", "undervoltage-release-light-sensor"): "1-rascepitel-i-datchik-sveta",
    ("mdk-04-02-1eo-24-s3", "three-phase-meter-connection"): "2-trehfaznyy-schetchik",
    ("mdk-04-02-1eo-24-s3", "electrical-apparatus-faults"): "3-neispravnosti-elektroapparatov",
    ("mdk-04-02-1eo-24-s3", "attestation"): "4-attestaciya",
    ("mdk-04-02-1eo-24-s3", "pue-and-gost"): "5-pue-i-gost",
    ("mdk-04-01", "soldering-tinning"): "1-payka-i-luzhenie",
    ("mdk-04-01", "benchwork-attestation"): "2-attestaciya-slesarnoe-delo",
    ("op-08", "information-resources"): "1-informacionnye-resursy",
    ("op-08", "information-systems"): "2-informacionnye-sistemy",
    ("op-08", "engineering-calculations-automation"): "3-avtomatizaciya-raschetov",
    ("op-08", "math-software-packages"): "4-pakety-dlya-raschetov",
    ("op-12", "constitutional-system"): "1-konstitucionnyy-stroy",
    ("op-12", "law-sources"): "2-istochniki-prava",
    ("op-12", "law-personality"): "3-pravo-i-lichnost",
    ("op-12", "npa-validity"): "4-deystvie-npa",
    ("op-12", "law-attestation"): "5-attestaciya-pravo",
    ("op-12", "judicial-authority"): "6-sudebnaya-vlast",
    ("op-03", "electromagnetism"): "1-elektromagnetizm",
    ("op-03", "electromagnetic-induction"): "2-elektromagnitnaya-indukciya",
    ("op-03", "alternating-current"): "3-peremenny-tok",
}

# старый id материала -> новый
MATERIALS = {
    "benchwork-attestation-practice": "1-attestaciya-slesarnoe-delo-attestation",
    "wire-connections-lecture": "2-vidy-soedineniy-zhil-lecture",
    "switch-socket-connection-practice": "1-vyklyuchateli-i-rozetki-practice",
    "lighting-network-check-practice": "2-proverka-osvetitelnoy-seti-practice",
    "circuit-breaker-selection-practice": "3-avtomaticheskiy-vyklyuchatel-practice",
    "light-sources-luminaires-lecture": "1-istochniki-sveta-i-armatura-lecture",
    "lighting-installation-repair-lecture": "2-montazh-i-remont-svetilnikov-lecture",
    "undervoltage-release-light-sensor-lecture": "1-rascepitel-i-datchik-sveta-lecture",
    "undervoltage-release-light-sensor-practice": "1-rascepitel-i-datchik-sveta-practice",
    "three-phase-meter-connection-practice": "2-trehfaznyy-schetchik-practice",
    "electrical-apparatus-faults-lecture": "3-neispravnosti-elektroapparatov-lecture",
    "mdk-04-02-attestation": "4-attestaciya-attestation",
    "pue-and-gost-lecture": "5-pue-i-gost-lecture",
    "soldering-tinning-lecture": "1-payka-i-luzhenie-lecture",
    "information-resources-lecture": "1-informacionnye-resursy-lecture",
    "information-systems-lecture": "2-informacionnye-sistemy-lecture",
    "engineering-calculations-automation-lecture": "3-avtomatizaciya-raschetov-lecture",
    "math-software-packages-lecture": "4-pakety-dlya-raschetov-lecture",
    "constitutional-system-lecture": "1-konstitucionnyy-stroy-lecture",
    "law-sources-lecture": "2-istochniki-prava-lecture",
    "law-personality-lecture": "3-pravo-i-lichnost-lecture",
    "npa-validity-lecture": "4-deystvie-npa-lecture",
    "law-attestation": "5-attestaciya-pravo-attestation",
    "judicial-authority-lecture": "6-sudebnaya-vlast-lecture",
    "electromagnetism-lecture": "1-elektromagnetizm-lecture",
    "electromagnetic-induction-lecture": "2-elektromagnitnaya-indukciya-lecture",
    "alternating-current-lecture": "3-peremenny-tok-lecture",
}


def replace_paths(text):
    for old, new in FOLDERS:
        text = text.replace(old + "/", new + "/")
    return text


def main():
    # 1. Переименовать папки
    for old, new in FOLDERS:
        src = os.path.join(ROOT, old)
        dst = os.path.join(ROOT, new)
        if not os.path.isdir(src):
            print("SKIP (нет папки):", old)
            continue
        if os.path.isdir(dst):
            raise SystemExit(f"Уже существует: {new}")
        try:
            os.rename(src, dst)
        except PermissionError:
            # папку держит проводник/индексатор — переносим содержимое
            os.makedirs(dst, exist_ok=True)
            for entry in os.listdir(src):
                shutil.move(os.path.join(src, entry), os.path.join(dst, entry))
            try:
                os.rmdir(src)
            except OSError:
                pass
        print("renamed:", old, "->", new.split("/")[-1])

    # 2. Каталог
    cat = json.load(open(CAT, encoding="utf-8"))
    for d in cat["disciplines"]:
        for t in d["topics"]:
            new_id = TOPICS.get((d["id"], t["id"]))
            if new_id:
                t["id"] = new_id
            t["materials"] = [MATERIALS.get(m, m) for m in t["materials"]]
    for m in cat["materials"]:
        m["id"] = MATERIALS.get(m["id"], m["id"])
        m["file"] = replace_paths(m["file"])
    with open(CAT, "w", encoding="utf-8", newline="\n") as fh:
        json.dump(cat, fh, ensure_ascii=False, indent=2)
        fh.write("\n")

    # 3. Ссылки внутри HTML
    changed = 0
    for old, new in FOLDERS:
        base = os.path.join(ROOT, new)
        for dirpath, _, files in os.walk(base):
            for name in files:
                if not name.lower().endswith(".html"):
                    continue
                p = os.path.join(dirpath, name)
                text = open(p, encoding="utf-8").read()
                updated = replace_paths(text)
                if updated != text:
                    open(p, "w", encoding="utf-8", newline="\n").write(updated)
                    changed += 1
    print("html updated:", changed)

    # 4. Проверка
    errors = []
    ids = [m["id"] for m in cat["materials"]]
    if len(ids) != len(set(ids)):
        errors.append("дубли id материалов")
    known = set(ids)
    for d in cat["disciplines"]:
        for t in d["topics"]:
            for mid in t["materials"]:
                if mid not in known:
                    errors.append(f"{d['id']}/{t['id']}: нет материала {mid}")
    for m in cat["materials"]:
        if not os.path.exists(os.path.join(ROOT, m["file"])):
            errors.append(f"нет файла: {m['file']}")
    print("errors:", errors if errors else "нет")


if __name__ == "__main__":
    main()
