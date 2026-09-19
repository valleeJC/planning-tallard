from datetime import datetime, timedelta, date
import os
import re
import time
import zoneinfo  # Module natif pour la gestion automatique des fuseaux horaires (Paris)
from icalendar import Calendar
import pandas as pd
import requests
import urllib3
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Désactivation des avertissements de sécurité SSL si verify=False est utilisé
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ==============================================================================
# 0. GESTION DU FUSEAU HORAIRE (EUROPE/PARIS)
# ==============================================================================

TZ_PARIS = zoneinfo.ZoneInfo("Europe/Paris")


def convertir_en_heure_paris(dt_object):
    """Gère l'attribution du fuseau horaire Europe/Paris sans appliquer

    de décalage UTC intempestif sur les heures locales transmises par ADE.
    """
    if dt_object is None:
        return None

    # Extraction de l'objet datetime brut depuis l'objet icalendar
    if hasattr(dt_object, "dt"):
        dt_object = dt_object.dt

    # S'il s'agit d'un objet date simple (ex: événement journée entière), on le laisse tel quel
    if not isinstance(dt_object, datetime):
        return dt_object

    # Si l'objet est 'naïf' (cas standard ADE), on lui assigne directement le fuseau Paris
    if dt_object.tzinfo is None:
        return dt_object.replace(tzinfo=TZ_PARIS)

    # Si le fichier iCal contenait déjà une timezone explicite (ex: UTC)
    return dt_object.astimezone(TZ_PARIS)


# ==============================================================================
# 1. CONFIGURATION
# ==============================================================================

URLS_ADE = {
    "LP MIA": (
        "https://agenda-web-consult.univ-amu.fr/jsp/custom/modules/plannings/anonymous_cal.jsp?projectId=8&resources=127685,127687,127688&calType=ical&firstDate=2026-08-17&lastDate=2027-08-15"
    ),
    "BUT 3 SNRV": (
        "https://agenda-web-consult.univ-amu.fr/jsp/custom/modules/plannings/anonymous_cal.jsp?projectId=8&resources=58374,58375&calType=ical&firstDate=2026-08-17&lastDate=2027-08-15"
    ),
    "M1 MPAD": (
        "https://agenda-web-consult.univ-amu.fr/jsp/custom/modules/plannings/anonymous_cal.jsp?projectId=8&resources=70202&calType=ical&firstDate=2026-08-17&lastDate=2027-08-15"
    ),
    "M2 MPAD": (
        "https://agenda-web-consult.univ-amu.fr/jsp/custom/modules/plannings/anonymous_cal.jsp?projectId=8&resources=391&calType=ical&firstDate=2026-08-17&lastDate=2027-08-15"
    ),
}

ENSEIGNANTS_AUTORISES = [
    "ATTAFI",
    "CORNUEAU",
    "MOYSAN",
    "GUEUDRE",
    "VALLEE",
    "SANCHEZ",
    "CHAVES-JACOB",
    "AMADEI",
    "MAZOYER",
    "RAYNAL",
]

FICHIER_EXCEL_SORTIE = "Planning_Voitures_Tallard.xlsx"
FICHIER_EXCEL_TEMP = "Planning_Voitures_Tallard_temp.xlsx"
FICHIER_HTML_SORTIE = "Planning_Voitures_Tallard.html"

JOURS_FR = {
    "Monday": "Lundi",
    "Tuesday": "Mardi",
    "Wednesday": "Mercredi",
    "Thursday": "Jeudi",
    "Friday": "Vendredi",
    "Saturday": "Samedi",
    "Sunday": "Dimanche",
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML,"
        " like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
}


# ==============================================================================
# 2. EXTRACTION ADE (AVEC SÉCURITÉ RÉSEAU & CONVERSION HORAIRE)
# ==============================================================================


def telecharger_ical_avec_retry(url, retries=2, backoff_factor=1):
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
            " (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
    }
    for i in range(retries):
        try:
            response = requests.get(url, headers=headers, timeout=(5, 5))
            response.raise_for_status()
            return response.content
        except requests.RequestException as e:
            print(f"⚠️ Tentative {i+1}/{retries} échouée")
            if i == retries - 1:
                raise e
            time.sleep(backoff_factor * (i + 1))


def est_enseignant_autorise(texte_evenement):
    texte_lower = texte_evenement.lower()
    for nom in ENSEIGNANTS_AUTORISES:
        if nom.lower() in texte_lower:
            return True, nom.upper()
    return False, "Vacataire / Inconnu"


def est_en_visio(texte_evenement):
    return bool(re.search(r"visio salle", texte_evenement, re.IGNORECASE))


def extraire_cours():
    tous_les_cours = []
    aujourdhui = date.today()
    il_y_a_7j = aujourdhui - timedelta(days=7)
    dans_7j = aujourdhui + timedelta(days=7)

    for cohorte, url in URLS_ADE.items():
        print(f"📥 Téléchargement du flux ADE pour {cohorte}...")
        try:
            contenu_ical = telecharger_ical_avec_retry(url)
            cal = Calendar.from_ical(contenu_ical)

            for component in cal.walk("VEVENT"):
                summary = str(component.get("summary", ""))
                description = str(component.get("description", ""))
                location = str(component.get("location", ""))

                # Extraction et conversion instantanée des heures UTC vers l'heure locale française
                dtstart_raw = component.get("dtstart")
                dtend_raw = component.get("dtend")

                if not dtstart_raw:
                    continue

                dtstart = convertir_en_heure_paris(dtstart_raw)
                dtend = convertir_en_heure_paris(dtend_raw)

                if not isinstance(dtstart, datetime):
                    continue

                texte_complet = f"{summary} {description} {location}"
                est_autorise, nom_prof = est_enseignant_autorise(texte_complet)
                visio_detectee = est_en_visio(texte_complet)

                if not est_autorise:
                    statut_voiture = "NON (Vacataire / Perso)"
                    code_statut = "NON"
                    motif = "Enseignant non habilité voiture de service"
                elif visio_detectee:
                    statut_voiture = "NON (Visio salle)"
                    code_statut = "NON"
                    motif = "Cours assuré à distance"
                else:
                    statut_voiture = "OUI (Voiture requise)"
                    code_statut = "OUI"
                    motif = "Présentiel - Enseignant habilité"

                dt_date = dtstart.date()
                jour_fr = JOURS_FR.get(
                    dtstart.strftime("%A"), dtstart.strftime("%A")
                )

                # Catégorisation temporelle pour tri & affichage
                est_7j_passes = il_y_a_7j <= dt_date < aujourdhui
                est_7j_futurs = aujourdhui <= dt_date <= dans_7j
                est_vieux_passe = dt_date < il_y_a_7j

                # Ordre de tri :
                # 0 = 7j passés, 1 = 7j futurs, 2 = Futurs lointains (+7j), 3 = Vieux passés (-7j)
                if est_7j_passes:
                    ordre_groupe = 0
                elif est_7j_futurs:
                    ordre_groupe = 1
                elif dt_date > dans_7j:
                    ordre_groupe = 2
                else:
                    ordre_groupe = 3

                tous_les_cours.append({
                    "_dt_start": dtstart,
                    "Date": dtstart.strftime("%d/%m/%Y"),
                    "Jour": jour_fr,
                    "Heure Début": dtstart.strftime("%H:%M"),
                    "Heure Fin": dtend.strftime("%H:%M"),
                    "Horaires": (
                        f"{dtstart.strftime('%H:%M')} -"
                        f" {dtend.strftime('%H:%M')}"
                    ),
                    "Cohorte": cohorte,
                    "Matière / Cours": summary,
                    "Enseignant détecté": nom_prof,
                    "Salle / Équipement": location,
                    "Visio Salle ?": "Oui" if visio_detectee else "Non",
                    "Besoin Voiture Service": statut_voiture,
                    "code_statut": code_statut,
                    "Explication": motif,
                    "est_7j_passes": est_7j_passes,
                    "est_7j_futurs": est_7j_futurs,
                    "est_vieux_passe": est_vieux_passe,
                    "ordre_groupe": ordre_groupe,
                })

        except Exception as e:
            print(f"❌ Échec de la récupération pour {cohorte}: {e}")

    # Tri : Groupe temporalisé puis chronologie dans chaque groupe
    tous_les_cours.sort(key=lambda x: (x["ordre_groupe"], x["_dt_start"]))
    return tous_les_cours


# ==============================================================================
# 3. GENERATION EXCEL
# ==============================================================================


def appliquer_mise_en_forme_excel(chemin_fichier):
    wb = openpyxl.load_workbook(chemin_fichier)

    fill_header = PatternFill(
        start_color="1F4E78", end_color="1F4E78", fill_type="solid"
    )
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

    fill_bleu = PatternFill(
        start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
    )
    font_bleu = Font(name="Calibri", size=10, color="1F4E78", bold=True)

    fill_vert = PatternFill(
        start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"
    )
    font_vert = Font(name="Calibri", size=10, color="375623", bold=True)

    fill_rouge = PatternFill(
        start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"
    )
    font_rouge = Font(name="Calibri", size=10, color="C00000")

    font_normal = Font(name="Calibri", size=10)

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    aujourdhui = date.today()
    il_y_a_7j = aujourdhui - timedelta(days=7)
    dans_7j = aujourdhui + timedelta(days=7)

    for sheetname in wb.sheetnames:
        ws = wb[sheetname]

        for cell in ws[1]:
            cell.fill = fill_header
            cell.font = font_header
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for row in ws.iter_rows(
            min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column
        ):
            date_str = ws.cell(row=row[0].row, column=1).value
            dt_cours = None
            if date_str:
                try:
                    dt_cours = datetime.strptime(
                        str(date_str), "%d/%m/%Y"
                    ).date()
                except ValueError:
                    pass

            for cell in row:
                cell.border = thin_border
                cell.alignment = Alignment(vertical="center")

                if dt_cours:
                    if il_y_a_7j <= dt_cours < aujourdhui:
                        cell.fill = fill_bleu
                        cell.font = font_bleu
                    elif aujourdhui <= dt_cours <= dans_7j:
                        cell.fill = fill_vert
                        cell.font = font_vert
                    elif dt_cours < il_y_a_7j:
                        cell.fill = fill_rouge
                        cell.font = font_rouge
                    else:
                        cell.font = font_normal

        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val = str(cell.value or "")
                if len(val) > max_len:
                    max_len = len(val)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(chemin_fichier)


def generer_excel(liste_cours):
    if not liste_cours:
        return

    df = pd.DataFrame(liste_cours)
    df_clean = df.drop(
        columns=[
            "_dt_start",
            "Horaires",
            "code_statut",
            "est_7j_passes",
            "est_7j_futurs",
            "est_vieux_passe",
            "ordre_groupe",
        ]
    )

    df_voiture_requise = df_clean[
        df_clean["Besoin Voiture Service"] == "OUI (Voiture requise)"
    ]

    with pd.ExcelWriter(FICHIER_EXCEL_TEMP, engine="openpyxl") as writer:
        df_voiture_requise.to_excel(
            writer, sheet_name="Voiture à réserver", index=False
        )
        df_clean.to_excel(
            writer, sheet_name="Planning Global ADE", index=False
        )

    appliquer_mise_en_forme_excel(FICHIER_EXCEL_TEMP)

    for i in range(5):
        try:
            if os.path.exists(FICHIER_EXCEL_SORTIE):
                os.remove(FICHIER_EXCEL_SORTIE)
            os.rename(FICHIER_EXCEL_TEMP, FICHIER_EXCEL_SORTIE)
            print(
                "📊 Fichier Excel généré :"
                f" {os.path.abspath(FICHIER_EXCEL_SORTIE)}"
            )
            break
        except PermissionError:
            time.sleep(3)


# ==============================================================================
# 4. GENERATION HTML
# ==============================================================================


def generer_tableau_html(cours_list, inclure_colonne_voiture=True):
    if not cours_list:
        return (
            "<p class='p-4 text-gray-500'>Aucun cours trouvé ou données"
            " indisponibles.</p>"
        )

    colonne_voiture_th = (
        '<th class="p-3">Besoin Voiture</th>'
        if inclure_colonne_voiture
        else ""
    )

    html = f"""
    <div class="overflow-x-auto">
        <table class="w-full text-left border-collapse text-sm">
            <thead>
                <tr class="bg-slate-800 text-white uppercase text-xs">
                    <th class="p-3">Date</th>
                    <th class="p-3">Jour</th>
                    <th class="p-3">Horaire</th>
                    <th class="p-3">Cohorte</th>
                    <th class="p-3">Matière / Cours</th>
                    <th class="p-3">Enseignant</th>
                    <th class="p-3">Salle</th>
                    {colonne_voiture_th}
                </tr>
            </thead>
            <tbody class="divide-y divide-gray-200">
    """
    for c in cours_list:
        if c["est_7j_passes"]:
            row_class = "bg-blue-50 text-blue-900 font-medium"
            badge_class = "bg-blue-200 text-blue-800"
        elif c["est_7j_futurs"]:
            row_class = "bg-emerald-50 text-emerald-900 font-medium"
            badge_class = "bg-emerald-200 text-emerald-800"
        elif c["est_vieux_passe"]:
            row_class = "bg-red-50 text-red-700 opacity-75"
            badge_class = "bg-red-200 text-red-800"
        else:
            row_class = "bg-white text-gray-800 hover:bg-gray-50"
            badge_class = "bg-gray-100 text-gray-700"

        colonne_voiture_td = (
            f'<td class="p-3"><span class="px-2.5 py-1 rounded-full text-xs'
            f' font-bold {badge_class}">{c["Besoin Voiture Service"]}</span></td>'
            if inclure_colonne_voiture
            else ""
        )

        html += f"""
        <tr class="{row_class} transition-colors">
            <td class="p-3 font-semibold">{c['Date']}</td>
            <td class="p-3">{c['Jour']}</td>
            <td class="p-3 whitespace-nowrap">{c['Horaires']}</td>
            <td class="p-3"><span class="px-2 py-1 rounded bg-slate-200 text-slate-800 text-xs font-bold">{c['Cohorte']}</span></td>
            <td class="p-3 font-medium">{c['Matière / Cours']}</td>
            <td class="p-3">{c['Enseignant détecté']}</td>
            <td class="p-3">{c['Salle / Équipement']}</td>
            {colonne_voiture_td}
        </tr>
        """
    html += "</tbody></table></div>"
    return html


def generer_html(cours):
    voitures_requises = [c for c in cours if c["code_statut"] == "OUI"]
    date_maj = datetime.now(TZ_PARIS).strftime("%d/%m/%Y à %H:%M")

    html_content = f"""<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Planning Voiture - Tallard</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-slate-100 font-sans min-h-screen p-4 md:p-8">
    <div class="max-w-7xl mx-auto bg-white rounded-xl shadow-md overflow-hidden">

        <div class="bg-slate-900 text-white p-6 flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
            <div>
                <h1 class="text-2xl font-bold">🚗 Planning Voiture & Cours — Tallard</h1>
                <p class="text-slate-400 text-sm mt-1">Dernière actualisation ADE : <span class="text-slate-200 font-semibold">{date_maj}</span></p>
            </div>
            <input type="text" id="searchInput" onkeyup="filtrerTableau()" placeholder="🔍 Rechercher (nom, cours, date)..." 
                   class="px-4 py-2 rounded-lg text-gray-900 text-sm w-full md:w-72 focus:outline-none focus:ring-2 focus:ring-blue-500">
        </div>

        <!-- LÉGENDE DE COULEURS -->
        <div class="bg-slate-50 px-6 py-3 border-b border-gray-200 flex flex-wrap gap-4 text-xs font-medium">
            <span class="flex items-center gap-1.5"><span class="w-3 h-3 rounded-full bg-blue-500"></span> 7 derniers jours (Passés récents)</span>
            <span class="flex items-center gap-1.5"><span class="w-3 h-3 rounded-full bg-emerald-500"></span> 7 prochains jours (À venir)</span>
            <span class="flex items-center gap-1.5"><span class="w-3 h-3 rounded-full bg-gray-300"></span> Futurs (+7 jours)</span>
            <span class="flex items-center gap-1.5"><span class="w-3 h-3 rounded-full bg-red-400"></span> Anciens cours (-7 jours)</span>
        </div>

        <!-- ONGLETS -->
        <div class="border-b border-gray-200 bg-white">
            <nav class="flex -mb-px px-6 gap-6">
                <button onclick="changerOnglet('tab-voitures')" id="btn-tab-voitures" class="tab-btn py-4 px-1 border-b-2 font-bold text-sm text-blue-600 border-blue-600">
                    🚘 Voiture à réserver ({len(voitures_requises)})
                </button>
                <button onclick="changerOnglet('tab-global')" id="btn-tab-global" class="tab-btn py-4 px-1 border-b-2 font-medium text-sm text-gray-500 border-transparent hover:text-gray-700">
                    📅 Planning Global ADE ({len(cours)})
                </button>
            </nav>
        </div>

        <div id="tab-voitures" class="tab-content">{generer_tableau_html(voitures_requises, inclure_colonne_voiture=False)}</div>
        <div id="tab-global" class="tab-content hidden">{generer_tableau_html(cours, inclure_colonne_voiture=True)}</div>

    </div>

    <script>
        function changerOnglet(tabId) {{
            document.querySelectorAll('.tab-content').forEach(el => el.classList.add('hidden'));
            document.querySelectorAll('.tab-btn').forEach(el => {{
                el.classList.remove('text-blue-600', 'border-blue-600', 'font-bold');
                el.classList.add('text-gray-500', 'border-transparent', 'font-medium');
            }});
            document.getElementById(tabId).classList.remove('hidden');
            const btn = document.getElementById('btn-' + tabId);
            btn.classList.add('text-blue-600', 'border-blue-600', 'font-bold');
            btn.classList.remove('text-gray-500', 'border-transparent', 'font-medium');
        }}

        function filtrerTableau() {{
            const input = document.getElementById('searchInput').value.toLowerCase();
            const activeTab = document.querySelector('.tab-content:not(.hidden)');
            const rows = activeTab.querySelectorAll('tbody tr');
            rows.forEach(row => {{
                const text = row.innerText.toLowerCase();
                row.style.display = text.includes(input) ? '' : 'none';
            }});
        }}
    </script>
</body>
</html>
"""

    with open(FICHIER_HTML_SORTIE, "w", encoding="utf-8") as f:
        f.write(html_content)

    print(f"🌐 Fichier HTML généré : {os.path.abspath(FICHIER_HTML_SORTIE)}")


# ==============================================================================
# 5. EXECUTION
# ==============================================================================

if __name__ == "__main__":
    cours = extraire_cours()

    # Génération systématique du fichier HTML (évite l'échec de la pipeline en cas d'erreur de téléchargement)
    generer_html(cours)

    if cours:
        generer_excel(cours)
        print("\n✅ Fichiers mis à jour avec la bonne heure locale française.")
