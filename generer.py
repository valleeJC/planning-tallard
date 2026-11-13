"""Planning voiture de service — Tallard (v3).

Récupère les EDT ADE, garde les cours en présentiel des enseignants habilités,
détecte les changements depuis la dernière synchro et génère docs/index.html
(publiable tel quel via GitHub Pages, dossier /docs).

Tout ce qui dépend de « aujourd'hui » (passé / à venir, séjours et
covoiturages possibles...) est calculé dans le navigateur : la page reste juste même si
elle n'a pas été régénérée depuis plusieurs jours.
"""

from datetime import datetime, timedelta
import json
from pathlib import Path
import re
import time
import zoneinfo

from icalendar import Calendar
import requests

# ==============================================================================
# 1. CONFIGURATION
# ==============================================================================

TZ_PARIS = zoneinfo.ZoneInfo("Europe/Paris")

ADE_BASE = (
    "https://agenda-web-consult.univ-amu.fr/jsp/custom/modules/plannings/"
    "anonymous_cal.jsp?projectId=8&calType=ical"
    "&firstDate=2026-08-17&lastDate=2027-08-15&resources="
)

COHORTES = {
    "LP MIA": "127685,127687,127688",
    "BUT 3 SNRV": "58374,58375",
    "M1 MPAD": "70202",
    "M2 MPAD": "391",
}

ENSEIGNANTS_AUTORISES = [
    "ATTAFI",
    "CORNUEAU",
    "GUEUDRE",
    "VALLEE",
    "SANCHEZ",
    "CHAVES-JACOB",
    "AMADEI",
    "MAZOYER",
    "RAYNAL",
    "MOYSAN",
]

# Le changelog ne garde que les modifications récentes.
CHANGELOG_JOURS = 30

DOSSIER = Path(__file__).resolve().parent
TEMPLATE = DOSSIER / "template.html"
SORTIE_HTML = DOSSIER / "docs" / "index.html"
SORTIE_JSON = DOSSIER / "docs" / "data.json"

RE_VISIO = re.compile(r"visio salle", re.IGNORECASE)
RE_PROFS = {
    nom: re.compile(rf"(?<![\w-]){re.escape(nom)}(?![\w-])", re.IGNORECASE)
    for nom in ENSEIGNANTS_AUTORISES
}


# ==============================================================================
# 2. EXTRACTION ADE
# ==============================================================================


def telecharger(url, essais=6):
    headers = {"User-Agent": "Mozilla/5.0 planning-tallard"}
    for i in range(essais):
        try:
            r = requests.get(url, headers=headers, timeout=(5, 30))
            r.raise_for_status()
            return r.content
        except requests.RequestException:
            print(f"  ⚠️ tentative {i + 1}/{essais} échouée")
            if i == essais - 1:
                raise
            time.sleep(i + 1)


def en_paris(valeur):
    dt = valeur.dt if hasattr(valeur, "dt") else valeur
    if not isinstance(dt, datetime):
        return None
    if dt.tzinfo is None:
        return dt.replace(tzinfo=TZ_PARIS)
    return dt.astimezone(TZ_PARIS)


def extraire_cours():
    """Renvoie ({uid: cours}, nb de flux lus, tous les flux lus ?).

    Un même cours partagé par deux cohortes n'apparaît qu'une fois
    (cohortes fusionnées)."""
    cours = {}
    flux_ok = 0

    for cohorte, ressources in COHORTES.items():
        print(f"📥 {cohorte}...")
        try:
            cal = Calendar.from_ical(telecharger(ADE_BASE + ressources))
        except Exception as e:
            print(f"  ❌ échec : {e}")
            continue
        flux_ok += 1

        for ev in cal.walk("VEVENT"):
            debut = en_paris(ev.get("dtstart"))
            fin = en_paris(ev.get("dtend"))
            if not debut or not fin:
                continue

            titre = str(ev.get("summary", "")).strip()
            description = str(ev.get("description", ""))
            salle = str(ev.get("location", "")).strip()
            texte = f"{titre}\n{description}\n{salle}"

            profs = [nom for nom, rx in RE_PROFS.items() if rx.search(texte)]
            if not profs or RE_VISIO.search(texte):
                continue

            uid = str(ev.get("uid", "")) or f"{debut:%Y%m%d%H%M}_{titre}"
            modif = ev.get("last-modified")
            modif = en_paris(modif) if modif else None

            if uid in cours:
                if cohorte not in cours[uid]["cohortes"]:
                    cours[uid]["cohortes"].append(cohorte)
                continue

            cours[uid] = {
                "uid": uid,
                "date": debut.date().isoformat(),
                "debut": debut.strftime("%H:%M"),
                "fin": fin.strftime("%H:%M"),
                "cohortes": [cohorte],
                "titre": titre,
                "salle": salle,
                "profs": profs,
                "modifie_ade": modif.date().isoformat() if modif else None,
            }

    return cours, flux_ok, flux_ok == len(COHORTES)


# ==============================================================================
# 3. DÉTECTION DES CHANGEMENTS
# ==============================================================================

CHAMPS_SUIVIS = ("date", "debut", "fin", "salle", "profs")


def resume(c):
    return {k: c[k] for k in ("date", "debut", "fin", "titre", "salle", "profs", "cohortes")}


def calculer_changelog(anciens, nouveaux, ancien_changelog, maintenant):
    """Compare avec la synchro précédente. Ne garde que les cours futurs :
    un cours passé qui disparaît du flux n'intéresse personne."""
    aujourdhui = maintenant.date().isoformat()
    horodatage = maintenant.isoformat(timespec="minutes")
    entrees = []

    for uid, n in nouveaux.items():
        a = anciens.get(uid)
        if a is None:
            if n["date"] >= aujourdhui:
                entrees.append({"type": "ajout", "quand": horodatage, "apres": resume(n)})
        elif any(a.get(k) != n.get(k) for k in CHAMPS_SUIVIS):
            if max(a["date"], n["date"]) >= aujourdhui:
                entrees.append({
                    "type": "modif", "quand": horodatage,
                    "avant": resume(a), "apres": resume(n),
                })

    for uid, a in anciens.items():
        if uid not in nouveaux and a["date"] >= aujourdhui:
            entrees.append({"type": "suppression", "quand": horodatage, "avant": resume(a)})

    limite = (maintenant - timedelta(days=CHANGELOG_JOURS)).isoformat(timespec="minutes")
    conserves = [e for e in ancien_changelog if e["quand"] >= limite]
    return entrees + conserves, len(entrees)


# ==============================================================================
# 4. GÉNÉRATION
# ==============================================================================


def charger_precedent():
    try:
        data = json.loads(SORTIE_JSON.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}, [], False
    return {c["uid"]: c for c in data["cours"]}, data.get("changelog", []), True


def generer(cours, complet):
    maintenant = datetime.now(TZ_PARIS)
    anciens, changelog, a_un_precedent = charger_precedent()

    if not a_un_precedent:
        print("🆕 Première synchro : pas de comparaison possible")
    elif not complet:
        # Un flux manquant ferait passer tous ses cours pour « supprimés ».
        print("⚠️ Flux incomplets : changements non calculés cette fois-ci")
    else:
        changelog, nb = calculer_changelog(anciens, cours, changelog, maintenant)
        print(f"🔁 {nb} changement(s) depuis la dernière synchro")

    data = {
        "maj": maintenant.isoformat(timespec="minutes"),
        "profs": ENSEIGNANTS_AUTORISES,
        "cohortes": list(COHORTES),
        "cours": sorted(cours.values(), key=lambda c: (c["date"], c["debut"])),
        "changelog": changelog,
    }

    SORTIE_HTML.parent.mkdir(exist_ok=True)
    SORTIE_JSON.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")

    # "</" échappé pour qu'un titre de cours ne puisse pas fermer la balise <script>.
    json_inline = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
    html = TEMPLATE.read_text(encoding="utf-8").replace("__DATA__", json_inline)
    SORTIE_HTML.write_text(html, encoding="utf-8")
    print(f"🌐 {SORTIE_HTML} ({len(cours)} cours)")


if __name__ == "__main__":
    cours, flux_ok, complet = extraire_cours()
    if flux_ok == 0:
        print("⚠️ Aucune donnée ADE récupérée : la page précédente est conservée.")
        raise SystemExit(1)
    generer(cours, complet)
