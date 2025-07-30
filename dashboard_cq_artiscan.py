
import getpass
from flask import Flask, jsonify, render_template_string, request, make_response
import pyodbc
from collections import defaultdict
from datetime import datetime, timedelta, date
import re
import pandas as pd
from apscheduler.schedulers.background import BackgroundScheduler
import requests
import urllib3
import io
import sqlite3

# Importer la config générale (machines, regex, etc.)
from config import SQL_CONFIG, MACHINES, EXCLUSION_REGEX, CQM_REGEX, CQH_REGEX, CQS_REGEX, MODULES_PAR_TYPE, WEBHOOK_URL, COMMENT_DB

# Demande des identifiants SQL à l'exécution
print("=== Authentification SQL ===")
username = input("Entrez l'identifiant SQL : ")
password = getpass.getpass("Entrez le mot de passe SQL : ")

# Connexion SQL
conn_str = f"DRIVER={{SQL Server}};SERVER={SQL_CONFIG['server']};DATABASE={SQL_CONFIG['database']};UID={username};PWD={password}"

app = Flask(__name__)

# Base de données commentaires légère sous forme de fichier pour justifier si il y a un non-conformité de périodicité sur un contrôle
def init_commentaires_db():
    conn = sqlite3.connect(COMMENT_DB)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS commentaires (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            machine TEXT,
            semaine TEXT,
            commentaire TEXT,
            auteur TEXT,
            date_commentaire DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.commit()
    conn.close()

init_commentaires_db()

def get_commentaires():
    conn = sqlite3.connect("commentaires_cq.db")
    c = conn.cursor()
    c.execute("SELECT machine, semaine, commentaire, auteur, date_commentaire FROM commentaires")
    data = c.fetchall()
    conn.close()
    # Indexation rapide par (machine, semaine)
    return {(row[0], row[1]): {"commentaire": row[2], "auteur": row[3], "date": row[4]} for row in data}


def get_taux_conformite(annee=2025):
    # 1. Génère les périodes de référence comme dans /cq_dashboard
    today = date.today()
    # Générer les semaines complètes pour 2024 et 2025 (lundi-vendredi)
    semaines = []
    current = date(2023, 12, 25)  # pour être sûr d'inclure S1 de 2024
    current -= timedelta(days=current.weekday())
    while current.year < 2026:
        y, w, _ = current.isocalendar()
        if y >= 2024:
            semaines.append({
                "Semaine": f"S{w}",
                "DateDebut": current,
                "DateFin": current + timedelta(days=4)
            })
        current += timedelta(weeks=1)
    df_semaines = pd.DataFrame(semaines)


    # Mois CQM
    mois = []
    for y in (2024, 2025):
        for m in range(1, 13):
            mois.append({
                "Mois": f"{datetime(y, m, 1).strftime('%B').capitalize()} {y}",
                "DateDebut": datetime(y, m, 1).date(),
                "DateFin": (datetime(y, m + 1, 1) - timedelta(days=1)).date() if m < 12 else date(y, 12, 31)
            })
    df_mois = pd.DataFrame(mois)


    # Semestres CQS
    semestres = [
        {"Semestre": "S1 2024", "DateDebut": date(2024, 1, 1), "DateFin": date(2024, 6, 30), "Year": 2024},
        {"Semestre": "S2 2024", "DateDebut": date(2024, 7, 1), "DateFin": date(2024, 12, 31), "Year": 2024},
        {"Semestre": "S1 2025", "DateDebut": date(2025, 1, 1), "DateFin": date(2025, 6, 30), "Year": 2025},
        {"Semestre": "S2 2025", "DateDebut": date(2025, 7, 1), "DateFin": date(2025, 12, 31), "Year": 2025},
    ]
    df_semestres = pd.DataFrame(semestres)

   

    # 2. Récupère les résultats CQ
    try:
        conn = pyodbc.connect(conn_str, timeout=5)
        rows = pd.read_sql("""
            SELECT cs.Id_Object, cs.Id_UserModule, cs.Name, cs.StudyDate
            FROM CONTROLE_STUDY cs
            JOIN RESULT r ON cs.Id_ControleStudy = r.Id_ControleStudy
            WHERE cs.StudyDate IS NOT NULL
        """, conn)
        conn.close()
    except Exception as e:
        print(f"Erreur SQL : {e}")
        return [], {}, {}, {}

    rows['StudyDate'] = pd.to_datetime(rows['StudyDate'])

    # Modules par type
    MODULES_PAR_TYPE = {
        "CQH": {66, 64, 63, 62, 61, 60},
        "CQM": {97, 98, 92, 90, 105, 104, 103},
        "CQS": {96, 95, 94, 93, 91, 106},
    }

    # 3. Construis les tableaux comme dans /cq_dashboard
    # Pour CQH
    real_cqh = []
    for _, row in rows.iterrows():
        module_id = row["Id_UserModule"]
        id_obj = row["Id_Object"]
        name = row["Name"]
        study_date = row["StudyDate"]

        # --- Correction: logique TOMO 1 et 2 comme dans dashboard car tout nos protocoles TOMO partagent le meme module sinon pas besoin de regex---
        if id_obj in [99, 121]:  # TOMO1 ou TOMO2
            if not CQH_REGEX.search(str(name)):
                continue
        else:
            if module_id not in MODULES_PAR_TYPE["CQH"]:
                continue

        if id_obj not in MACHINES:
            continue
        if pd.isna(study_date):
            continue
        study_date = pd.to_datetime(study_date).date()
        real_cqh.append({"Id_Object": id_obj, "Date": study_date})

    df_cqh = pd.DataFrame(real_cqh)
    machines_cqh = sorted(df_cqh['Id_Object'].unique())
    cqh_data = []
    for _, sem in df_semaines.iterrows():
        row = {"Semaine": sem["Semaine"], "Year": sem["DateDebut"].year}
        for id_obj in machines_cqh:
            machine_name = MACHINES.get(id_obj, ("❓",))[0]
            done = df_cqh[
                (df_cqh["Id_Object"] == id_obj) &
                (df_cqh["Date"] >= sem["DateDebut"]) &
                (df_cqh["Date"] <= sem["DateFin"])
            ]
            if not done.empty:
                row[machine_name] = "✅"
            elif sem["DateFin"] < today:
                row[machine_name] = "❌"
            else:
                row[machine_name] = "⏳"
        cqh_data.append(row)
    df_cqh_final = pd.DataFrame(cqh_data)

    # Pour CQM
    real_cqm = []
    for _, row in rows.iterrows():
        if is_valid_cq_name(row["Name"], "CQM", module_id=row["Id_UserModule"]):
            result_date = pd.to_datetime(row["StudyDate"]).date()
            id_obj = row["Id_Object"]
            machine = MACHINES.get(id_obj, (None,))[0]
            if not machine: continue
            match = df_mois[
                (df_mois["DateDebut"] <= result_date) &
                (df_mois["DateFin"] >= result_date)
            ]
            if match.empty: continue
            real_cqm.append({
                "Machine": machine,
                "Date": result_date
            })
    df_cqm = pd.DataFrame(real_cqm)
    machines_cqm = sorted(df_cqm['Machine'].unique())
    cqm_data = []
    for _, mois in df_mois.iterrows():
        row = {"Mois": mois["Mois"]}
        for m in machines_cqm:
            done = df_cqm[
                (df_cqm["Machine"] == m) &
                (df_cqm["Date"] >= mois["DateDebut"]) &
                (df_cqm["Date"] <= mois["DateFin"])
            ]
            if not done.empty:
                row[m] = "✅"
            elif mois["DateFin"] < today:
                row[m] = "❌"
            else:
                row[m] = "⏳"
        cqm_data.append(row)
    df_cqm_final = pd.DataFrame(cqm_data)

    # Pour CQS
    real_cqs = []
    for _, row in rows.iterrows():
        if is_valid_cq_name(row["Name"], "CQS", module_id=row["Id_UserModule"]):
            result_date = pd.to_datetime(row["StudyDate"]).date()
            id_obj = row["Id_Object"]
            machine = MACHINES.get(id_obj, (None,))[0]
            if machine:
                real_cqs.append({
                    "Machine": machine,
                    "Date": result_date
                })
    df_cqs = pd.DataFrame(real_cqs)
    machines_cqs = sorted(df_cqs['Machine'].unique())
    cqs_data = []
    for _, sem in df_semestres.iterrows():
        row = {"Semestre": sem["Semestre"], "Year": sem["Year"]}
        for m in machines_cqs:
            done = df_cqs[
                (df_cqs["Machine"] == m) &
                (df_cqs["Date"] >= sem["DateDebut"]) &
                (df_cqs["Date"] <= sem["DateFin"])
            ]
            if not done.empty:
                row[m] = "✅"
            elif sem["DateFin"] < today:
                row[m] = "❌"
            else:
                row[m] = "⏳"
        cqs_data.append(row)

    df_cqs_final = pd.DataFrame(cqs_data)

    if 'Year' in df_cqh_final:
        df_cqh_final = df_cqh_final[df_cqh_final['Year'] == annee]
    if 'Year' in df_cqm_final:
        df_cqm_final = df_cqm_final[df_cqm_final['Year'] == annee]
    if 'Year' in df_cqs_final:
        df_cqs_final = df_cqs_final[df_cqs_final['Year'] == annee] 

    # 4. Calcul des taux de conformité pour chaque machine
        # Liste brute de toutes les colonnes (hors 1re colonne)
    cols_cqh = [c for c in df_cqh_final.columns if c not in ("Semaine", "Year")]
    cols_cqm = [c for c in df_cqm_final.columns if c not in ("Mois", "Year")]
    cols_cqs = [c for c in df_cqs_final.columns if c not in ("Semestre", "Year")]

    machines = sorted(set(cols_cqh) | set(cols_cqm) | set(cols_cqs))

    taux_cqh, taux_cqm, taux_cqs = {}, {}, {}

    for m in machines:
        # Trouve la colonne correspondante, insensible à la casse et aux espaces
        col_cqh = [c for c in df_cqh_final.columns if c.replace(" ", "").upper() == m.replace(" ", "").upper()]
        if col_cqh:
            col = col_cqh[0]
            conforme = (df_cqh_final[col] == "✅").sum()
            a_juger = (df_cqh_final[col].isin(["✅", "❌"])).sum()
            taux_cqh[m] = round((conforme / a_juger) * 100, 1) if a_juger > 0 else 100.0
        else:
            taux_cqh[m] = 0.0

        col_cqm = [c for c in df_cqm_final.columns if c.replace(" ", "").upper() == m.replace(" ", "").upper()]
        if col_cqm:
            col = col_cqm[0]
            conforme = (df_cqm_final[col] == "✅").sum()
            a_juger = (df_cqm_final[col].isin(["✅", "❌"])).sum()
            taux_cqm[m] = round((conforme / a_juger) * 100, 1) if a_juger > 0 else 100.0
        else:
            taux_cqm[m] = 0.0

        col_cqs = [c for c in df_cqs_final.columns if c.replace(" ", "").upper() == m.replace(" ", "").upper()]
        if col_cqs:
            col = col_cqs[0]
            conforme = (df_cqs_final[col] == "✅").sum()
            a_juger = (df_cqs_final[col].isin(["✅", "❌"])).sum()
            taux_cqs[m] = round((conforme / a_juger) * 100, 1) if a_juger > 0 else 100.0
        else:
            taux_cqs[m] = 0.0

    


    return machines, taux_cqh, taux_cqm, taux_cqs




def is_valid_cq_name(name, typ, module_id=None):
    name_clean = name.lower().replace(" ", "").replace("_", "")

    # 🔴 Exclusion forte (à conserver)
    EXCLUSION_REGEX = re.compile(r"(?i)(test|à ?supp|a ?supp|à ?supprimer|a ?supprimer|essai|asupprimer|àsupprimer)")
    if EXCLUSION_REGEX.search(name):
        return False

    # ✅ CQH : seulement exclusion, plus de test sur le nom
    if typ == "CQH":
        return module_id in MODULES_PAR_TYPE["CQH"]

    # ✅ CQM : on garde regex + fallback sur module
    if typ == "CQM":
        CQM_REGEX = re.compile(r"(?i)(cqm|controlequalitemensuel(le)?|contrôlequalitémensuel(le)?)")
        return bool(CQM_REGEX.search(name)) or (module_id in MODULES_PAR_TYPE["CQM"])

    # ✅ CQS (même logique que CQM)
    if typ == "CQS":
        CQS_REGEX = re.compile(r"(?i)(cqs|controlequalitesemestriel(le)?|contrôlequalitésemestriel(le)?)")
        return bool(CQS_REGEX.search(name)) or (module_id in MODULES_PAR_TYPE["CQS"])

    return False



@app.route("/")
def index():
    annee = int(request.args.get("annee", 2025))
    machines, taux_cqh, taux_cqm, taux_cqs = get_taux_conformite(annee)
    today = date.today()
    week_number = today.isocalendar()[1]
    total_weeks = date(today.year, 12, 28).isocalendar()[1]  # 28 déc = dernière semaine ISO
    progress_percent = round((week_number / total_weeks) * 100, 1)

    def moyenne(taux):
        vals = [v for v in taux.values() if isinstance(v, (int, float)) and v is not None]
        return round(sum(vals) / len(vals), 1) if vals else 0

    moyenne_cqh = moyenne(taux_cqh)
    moyenne_cqm = moyenne(taux_cqm)
    moyenne_cqs = moyenne(taux_cqs)
    return render_template_string("""
<!DOCTYPE html>
<html lang="fr">
<head>
  <meta charset="UTF-8">
  <title>📅 Planning CQ - Artiscan</title>
  <link href="https://cdn.jsdelivr.net/npm/fullcalendar@6.1.8/main.min.css" rel="stylesheet">
  <script src="https://cdn.jsdelivr.net/npm/fullcalendar@6.1.8/index.global.min.js"></script>
  <style>
    body {
      font-family: 'Segoe UI', Arial, sans-serif;
      background: #f6f8fb;
      margin: 0;
      padding: 0;
    }
    header {
      background: #fff;
      box-shadow: 0 2px 14px rgba(32,40,64,0.07);
      padding: 25px 0 14px 0;
      border-bottom: 1px solid #eee;
      margin-bottom: 0;
    }
    .header-bar {
      max-width: 1400px;
      margin: auto;
      display: flex;
      align-items: center;
      justify-content: space-between;
    }
    .header-title {
      font-size: 2rem;
      font-weight: 600;
      color: #253858;
      letter-spacing: -1px;
      display: flex;
      align-items: center;
      gap: 16px;
    }
    .header-logo {
      height: 120px;
    }
    #filterBar {
      max-width: 1400px;
      margin: 38px auto 0 auto;
      text-align: center;
      background: #fff;
      border-radius: 16px;
      padding: 18px 0 6px 0;
      box-shadow: 0 2px 14px #0001;
    }
    #searchInput {
      padding: 12px 18px;
      border-radius: 22px;
      border: 1px solid #d3d3e0;
      width: 320px;
      box-shadow: 0 2px 8px #0001;
      font-size: 1.1rem;
      margin-bottom: 10px;
    }
    label {
      margin: 0 8px;
      font-weight: 500;
    }
    #calendar {
      max-width: 1600px;
      margin: 35px auto 0 auto;
      background: #fff;
      border-radius: 26px;
      padding: 36px 24px;
      box-shadow: 0 8px 40px rgba(32,40,64,0.12);
    }
    #stats {
      max-width: 1000px;
      margin: 40px auto;
      background: #fff;
      padding: 28px;
      border-radius: 16px;
      box-shadow: 0 0 10px rgba(0,0,0,0.09);
    }
    table {
      width: 100%;
      border-collapse: collapse;
      margin-top: 10px;
      font-size: 1rem;
    }
    th, td {
      padding: 9px;
      text-align: center;
      border: 1px solid #ececec;
    }
    th {
      background-color: #f0f0f0;
    }
    .btn-primary {
      background: #0051ba;
      border: none;
      border-radius: 18px;
      font-weight: 600;
      font-size: 1.1rem;
      padding: 10px 32px;
      color: #fff;
      transition: background .2s;
      box-shadow: 0 2px 8px #0051ba20;
      margin-top: 14px;
    }
    .btn-primary:hover {
      background: #0063cc;
      color: #fff;
    }
  </style>
</head>
<body>
  <!-- Barre d'en-tête moderne -->
  <header>
    <div class="header-bar">
      <div class="header-title">
        📊 Suivi de la réalisation périodique des Contrôles Qualité à l'Institut Gustave Roussy
      </div>
      <img src="/static/logo_gustave_roussy_rvb.jpg" alt="Gustave Roussy" class="header-logo">
    </div>
  </header>

  <div id="filterBar">
    <input type="text" id="searchInput" placeholder="🔍 Rechercher un mot-clé...">
    <br><br>
    <label><input type="checkbox" class="type-filter" value="CQH" checked> CQH</label>
    <label><input type="checkbox" class="type-filter" value="CQM" checked> CQM</label>
    <label><input type="checkbox" class="type-filter" value="CQS" checked> CQS</label>
    <label><input type="checkbox" class="type-filter" value="CQQ"> CQQ</label>
    <label><input type="checkbox" class="type-filter" value="TOMO" checked> TOMO</label>
  </div>
  <div style="text-align:center; margin: 30px;">
    <a href="/cq_dashboard" class="btn btn-primary">📊 Suivi Global CQ</a>
  </div>
  <div id="calendar"></div>



  <div id="stats">
    <h2 style="margin-bottom:22px;">📈 Taux de conformité par machine</h2>
    <div style="margin-bottom:15px;">
    <form method="get" style="display:inline;">
        <label>Sélectionner l'année :
        <select name="annee" onchange="this.form.submit()">
            <option value="2024" {% if annee == 2024 %}selected{% endif %}>2024</option>
            <option value="2025" {% if annee == 2025 %}selected{% endif %}>2025</option>
        </select>
        </label>
    </form>
    </div>
    <table>
      <thead>
        <tr>
          <th>Machine</th>
          <th>CQH (%)</th>
          <th>CQM (%)</th>
          <th>CQS (%)</th>
        </tr>
      </thead>
      <tbody>
        {% for m in machines %}
        <tr>
          <td>{{ m }}</td>
          <td>{{ taux_cqh[m] }}</td>
          <td>{{ taux_cqm[m] }}</td>
          <td>{{ taux_cqs[m] }}</td>
        </tr>
        {% endfor %}
      </tbody>
        <tfoot>
            <tr style="background:#f3f3f3; font-weight:600;">
             <td>Moyenne conformité</td>
             <td>{{ moyenne_cqh }}</td>
             <td>{{ moyenne_cqm }}</td>
             <td>{{ moyenne_cqs }}</td>
             </tr>
        </tfoot>                            
    </table>
    <div style="max-width: 600px; margin: 30px auto 0;">
      <div style="display:flex; justify-content:space-between; margin-bottom: 6px;">
        <span>Progression annuelle :</span>
        <span>{{ week_number }}/{{ total_weeks }} semaines ({{ progress_percent }}%)</span>
      </div>
      <div style="background: #eee; border-radius: 6px; overflow: hidden;">
        <div style="height: 22px; background: linear-gradient(90deg, #73d13d, #4096ff); width: {{ progress_percent }}%; transition: width 0.8s;"></div>
      </div>
    </div>
  </div>
                                  
  <script>
    let allEvents = [];
    let calendar;

    async function loadCalendar() {
      const res = await fetch("/cq");
      allEvents = await res.json();

      calendar = new FullCalendar.Calendar(document.getElementById("calendar"), {
        initialView: 'dayGridMonth',
        locale: 'fr',
        firstDay: 1,
        hiddenDays: [0, 6],
        headerToolbar: {
          left: 'prev,next today',
          center: 'title',
          right: 'dayGridMonth,workWeek,listWeek'
        },
        views: {
          workWeek: {
            type: 'timeGridWeek',
            buttonText: 'Semaine (L-V)'
          }
        },
        events: allEvents,
        eventClick: function(info) {
          const e = info.event;
          alert(`🗂 ${e.title}\n📆 ${e.start.toLocaleString()}\n🖥️ Machine : ${e.extendedProps.machine}`);
        }
      });

      calendar.render();
    }

    function applyFilters() {
      const search = document.getElementById("searchInput").value.toLowerCase();
      const activeTypes = Array.from(document.querySelectorAll(".type-filter:checked")).map(cb => cb.value);

      const filtered = allEvents.filter(ev => {
        const matchesType = activeTypes.some(type => ev.title.startsWith(type));
        const matchesSearch = ev.title.toLowerCase().includes(search);
        return matchesType && matchesSearch;
      });

      calendar.removeAllEvents();
      calendar.addEventSource(filtered);
    }

    document.addEventListener("DOMContentLoaded", async function () {
      await loadCalendar();
      applyFilters();

      document.getElementById("searchInput").addEventListener("input", applyFilters);
      document.querySelectorAll(".type-filter").forEach(cb => cb.addEventListener("change", applyFilters));
    });
  </script>
</body>
</html>
""", taux_cqh=taux_cqh, taux_cqm=taux_cqm, taux_cqs=taux_cqs, machines=machines, progress_percent=progress_percent, week_number=week_number, total_weeks=total_weeks, moyenne_cqh=moyenne_cqh, moyenne_cqm=moyenne_cqm, moyenne_cqs=moyenne_cqs, annee=annee)



@app.route('/cq')
def get_cq():
    try:
        conn = pyodbc.connect(conn_str, timeout=5)
        cursor = conn.cursor()

        query = """
        SELECT cs.Id_ControleStudy, cs.Id_Object, cs.Name, cs.Id_UserModule, cs.StudyDate
        FROM CONTROLE_STUDY cs
        JOIN RESULT r ON cs.Id_ControleStudy = r.Id_ControleStudy
        ORDER BY cs.Id_ControleStudy DESC
        """
        cursor.execute(query)
        rows = cursor.fetchall()
        #rows = pd.DataFrame(rows, columns=["Id_ControleStudy", "Id_Object", "Name", "Id_UserModule", "ResultDate"])

        conn.close()

        CQ_TYPE = {
            'CQH': {66, 64, 63, 62, 61, 60},
            'CQM': {97, 98, 92, 90, 105, 104, 103},
            'CQS': {96, 95, 94, 93, 91, 106},
            'CQQ': {25, 27, 28, 29, 30, 32, 36, 37, 38},
            'TOMO': {24, 26}
        }

        events = []
        for row in rows:
            id_control, id_object, name, id_user_module, study_date = row

            if not study_date:
                continue  
            if isinstance(study_date, str):
                study_date = pd.to_datetime(study_date)
            if hasattr(study_date, "date"):
                study_date = study_date

            cq_prefix = ""
            for label, ids in CQ_TYPE.items():
                if id_user_module in ids:
                    cq_prefix = f"{label} - "
                    break

            if id_object in MACHINES:
                machine_name, color = MACHINES[id_object]
                events.append({
                    'id': id_control,
                    'title': f"{name} ({machine_name})",
                    'start': study_date.strftime('%Y-%m-%dT%H:%M:%S'),
                    'color': color,
                    'machine': machine_name
                })


        return jsonify(events)

    except Exception as e:
        print(f"❌ Erreur dans /cq : {e}")
        return jsonify([])

import io
from flask import make_response

@app.route('/export_cqh_csv')
def export_cqh_csv():
    try:
        # --- Génère ou récupère df_cqh_final comme dans ton dashboard ---
        today = date.today()
        semaines = []
        # Générer les semaines complètes pour 2024 et 2025 (lundi-vendredi)
        semaines = []
        current = date(2023, 12, 25)  # pour être sûr d'inclure S1 de 2024
        current -= timedelta(days=current.weekday())
        while current.year < 2026:
            y, w, _ = current.isocalendar()
            if y >= 2024:
                semaines.append({
                    "Semaine": f"S{w}",
                    "DateDebut": current,
                    "DateFin": current + timedelta(days=4)
                })
            current += timedelta(weeks=1)
        df_semaines = pd.DataFrame(semaines)


        conn = pyodbc.connect(conn_str, timeout=5)
        rows = pd.read_sql("""
            SELECT cs.Id_Object, cs.Id_UserModule, cs.Name, cs.StudyDate
            FROM CONTROLE_STUDY cs
            JOIN RESULT r ON cs.Id_ControleStudy = r.Id_ControleStudy
            WHERE cs.StudyDate IS NOT NULL
        """, conn)
        conn.close()
        rows['StudyDate'] = pd.to_datetime(rows['StudyDate'])

        MODULES_PAR_TYPE = {
            "CQH": {66, 64, 63, 62, 61, 60},
        }

        real_cqh = []
        for _, row in rows.iterrows():
            module_id = row["Id_UserModule"]
            id_obj = row["Id_Object"]
            study_date = row["StudyDate"]
            if module_id not in MODULES_PAR_TYPE["CQH"]:
                continue
            if id_obj not in MACHINES:
                continue
            if pd.isna(study_date):
                continue
            study_date = pd.to_datetime(study_date).date()
            real_cqh.append({"Id_Object": id_obj, "Date": study_date})
        df_cqh = pd.DataFrame(real_cqh)

        cqh_data = []
        for _, sem in df_semaines.iterrows():
            row = {"Semaine": sem["Semaine"], "Year": sem["DateDebut"].year}
            for id_obj in MACHINES.keys():
                machine_name = MACHINES.get(id_obj, ("❓",))[0]
                done = df_cqh[
                    (df_cqh["Id_Object"] == id_obj) &
                    (df_cqh["Date"] >= sem["DateDebut"]) &
                    (df_cqh["Date"] <= sem["DateFin"])
                ]
                if not done.empty:
                    row[machine_name] = "✅"
                elif sem["DateFin"] < today:
                    row[machine_name] = "❌"
                else:
                    row[machine_name] = "⏳"
            cqh_data.append(row)
        df_cqh_final = pd.DataFrame(cqh_data)

        # -------- Export UTF-8 avec BOM pour Excel --------
        output = io.StringIO()
        df_cqh_final.to_csv(output, index=False, sep=';')
        csv_content = output.getvalue()
        output.close()
        csv_content = '\ufeff' + csv_content  # Ajoute BOM UTF-8

        response = make_response(csv_content)
        response.headers["Content-Disposition"] = "attachment; filename=cqh_dashboard.csv"
        response.headers["Content-Type"] = "text/csv; charset=utf-8"
        return response

    except Exception as e:
        return f"Erreur export : {e}", 500




@app.route("/cq_dashboard")
def cq_dashboard():
    # Récupération des deux DataFrames (réutilisation des deux logiques précédentes)
    # --- Semaines pour CQH ---
    # Générer les semaines complètes pour 2024 et 2025 (lundi-vendredi)
    semaines = []
    current = date(2023, 12, 25)  # pour être sûr d'inclure S1 de 2024
    current -= timedelta(days=current.weekday())
    while current.year < 2026:
        y, w, _ = current.isocalendar()
        if y >= 2024:
            semaines.append({
                "Semaine": f"S{w}",
                "DateDebut": current,
                "DateFin": current + timedelta(days=4)
            })
        current += timedelta(weeks=1)
    df_semaines = pd.DataFrame(semaines)


    # --- Mois pour CQM ---
    mois = []
    for y in (2024, 2025):
        for m in range(1, 13):
            mois.append({
                "Mois": f"{datetime(y, m, 1).strftime('%B').capitalize()} {y}",
                "Year": y,  # AJOUTE CETTE LIGNE
                "DateDebut": datetime(y, m, 1).date(),
                "DateFin": (datetime(y, m + 1, 1) - timedelta(days=1)).date() if m < 12 else date(y, 12, 31)
            })

    df_mois = pd.DataFrame(mois)



    # --- Mois pour CQS ---
    semestres = [
        {"Semestre": "S1 2024", "DateDebut": date(2024, 1, 1), "DateFin": date(2024, 6, 30), "Year": 2024},
        {"Semestre": "S2 2024", "DateDebut": date(2024, 7, 1), "DateFin": date(2024, 12, 31), "Year": 2024},
        {"Semestre": "S1 2025", "DateDebut": date(2025, 1, 1), "DateFin": date(2025, 6, 30), "Year": 2025},
        {"Semestre": "S2 2025", "DateDebut": date(2025, 7, 1), "DateFin": date(2025, 12, 31), "Year": 2025},
    ]

    df_semestres = pd.DataFrame(semestres)


    # --- Récupération des CQ réalisés ---
    try:
        conn = pyodbc.connect(conn_str, timeout=5)
        rows = pd.read_sql("""
            SELECT cs.Id_Object, cs.Id_UserModule, cs.Name, cs.StudyDate
            FROM CONTROLE_STUDY cs
            JOIN RESULT r ON cs.Id_ControleStudy = r.Id_ControleStudy
            WHERE cs.StudyDate IS NOT NULL
        """, conn)
        conn.close()
    except Exception as e:
        return f"Erreur SQL : {e}"

    rows['StudyDate'] = pd.to_datetime(rows['StudyDate'])


  
    # --- CQH ---


    MODULES_PAR_TYPE = {
        "CQH": {66, 64, 63, 62, 61, 60},
        "CQM": {97, 98, 92, 90, 105, 104, 103},
        "CQS": {96, 95, 94, 93, 91, 106},
    }

    global UNKNOWN_MACHINES_CQH
    UNKNOWN_MACHINES_CQH = []

    real_cqh = []
    cqh_outside_weeks = []

    for _, row in rows.iterrows():
        module_id = row["Id_UserModule"]
        id_obj = row["Id_Object"]
        name = row["Name"]
        study_date = row["StudyDate"]

        # 🎯 Ne retenir que les modules CQH sauf pour TOMO 1 et TOMO 2
        if id_obj in [99, 121]:  # TOMO1 ou TOMO2
            if not CQH_REGEX.search(str(name)):
                continue
        else:
            # Les autres => classique via module
            if module_id not in MODULES_PAR_TYPE["CQH"]:
                continue

        # 🔒 Vérifie que la machine est connue
        if id_obj not in MACHINES:
            UNKNOWN_MACHINES_CQH.append(row)
            continue

        # 🗓 Vérifie que la date est valable
        if pd.isna(study_date):
            continue

        study_date = pd.to_datetime(study_date).date()

        # 🧪 Vérifie que la date est dans une semaine déclarée
        in_range = df_semaines[
            (df_semaines["DateDebut"] <= study_date) &
            (df_semaines["DateFin"] >= study_date)
        ]
        if in_range.empty:
            cqh_outside_weeks.append((name, id_obj, study_date))

        real_cqh.append({
            "Id_Object": id_obj,
            "Date": study_date
        })


    if module_id in MODULES_PAR_TYPE["CQH"]:
      if not is_valid_cq_name(name, "CQH", module_id):
          print(f"❌ Rejeté par nom : {name}")
      elif id_obj not in MACHINES:
          print(f"❌ Machine inconnue : {id_obj} pour {name}")
      elif in_range.empty:
          print(f"❌ Hors plage semaine : {name} — {study_date}")

    # ✅ Logs de debug
    #print("🔎 Machines non reconnues :", len(UNKNOWN_MACHINES_CQH))
    #print("🔎 CQH hors des plages de semaines :", len(cqh_outside_weeks))

    if cqh_outside_weeks:
        #print("⚠️ CQH hors semaine (à partir de 2025) :")
        filtered = [item for item in cqh_outside_weeks if item[2] >= datetime(2025, 1, 1).date()]
        for item in filtered:
            print(f"❌ {item[0]} — Machine ID {item[1]} — Date {item[2]}")
        print(f"➡️ Total après 2025 : {len(filtered)} / {len(cqh_outside_weeks)} hors semaine.")


    # ✅ Construction du tableau CQH final
    df_cqh = pd.DataFrame(real_cqh)
    machines_cqh = sorted(MACHINES.keys())


    cqh_data = []
    today = date.today()

    for _, sem in df_semaines.iterrows():
        row = {"Semaine": sem["Semaine"], "Year": sem["DateDebut"].year}
        for id_obj in machines_cqh:
            machine_name = MACHINES.get(id_obj, ("❓",))[0]
            done = df_cqh[
                (df_cqh["Id_Object"] == id_obj) &
                (df_cqh["Date"] >= sem["DateDebut"]) &
                (df_cqh["Date"] <= sem["DateFin"])
            ]
            if not done.empty:
                row[machine_name] = "✅"
            elif sem["DateFin"] < today:
                row[machine_name] = "❌"
            else:
                row[machine_name] = "⏳"
        cqh_data.append(row)

    df_cqh_final = pd.DataFrame(cqh_data)

    # ✅ Construction du tableau CQM final
    real_cqm = []
    cqm_outside_months = []
    unknown_machine_cqm = []

    for _, row in rows.iterrows():
        module_id = row["Id_UserModule"]
        id_obj = row["Id_Object"]
        name = row["Name"]
        study_date = row["StudyDate"]

        # 🎯 TOMO1 et TOMO2 => se baser sur le nom
        if id_obj in [99, 121]:  # TOMO1 ou TOMO2
            if not CQM_REGEX.search(str(name)):
                continue
        else:
            if module_id not in MODULES_PAR_TYPE["CQM"]:
                continue

        if pd.isna(study_date):
            continue

        result_date = pd.to_datetime(study_date).date()
        machine = MACHINES.get(id_obj, (None,))[0]

        if not machine:
            unknown_machine_cqm.append((name, id_obj, result_date))
            continue

        match = df_mois[
            (df_mois["DateDebut"] <= result_date) & (df_mois["DateFin"] >= result_date)
        ]
        if match.empty:
            cqm_outside_months.append((name, id_obj, result_date))
            continue

        real_cqm.append({
            "Machine": machine,
            "Date": result_date
        })
    # Logs utiles
    #print("🔍 Machines non reconnues (CQM) :", len(unknown_machine_cqm))
    #print("🔍 CQM hors des plages de mois :", len(cqm_outside_months))
    if cqm_outside_months:
        #print("⚠️ CQM hors mois (exemples) :")
        for item in cqm_outside_months[:10]:
            print(item)


    df_cqm = pd.DataFrame(real_cqm)
    machines_cqm = sorted(df_cqm['Machine'].unique())
    cqm_data = []

    for _, mois in df_mois.iterrows():
        row = {"Mois": mois["Mois"], "Year": mois["Year"]}
        for m in machines_cqm:
            done = df_cqm[
                (df_cqm["Machine"] == m) &
                (df_cqm["Date"] >= mois["DateDebut"]) &
                (df_cqm["Date"] <= mois["DateFin"])
            ]
            if not done.empty:
                row[m] = "✅"
            elif mois["DateFin"] < today:
                row[m] = "❌"
            else:
                row[m] = "⏳"
        cqm_data.append(row)

    df_cqm_final = pd.DataFrame(cqm_data)

    # ✅ Construction du tableau CQS final
    real_cqs = []
    unknown_machine_cqs = []

    for _, row in rows.iterrows():
        module_id = row["Id_UserModule"]
        id_obj = row["Id_Object"]
        name = row["Name"]
        study_date = row["StudyDate"]

        # 🎯 TOMO1 et TOMO2 => filtrage par nom (CQS)
        if id_obj in [99, 121]:  # TOMO1 ou TOMO2
            if not CQS_REGEX.search(str(name)):
                continue
        else:
            if module_id not in MODULES_PAR_TYPE["CQS"]:
                continue

        if pd.isna(study_date):
            continue

        result_date = pd.to_datetime(study_date).date()
        machine = MACHINES.get(id_obj, (None,))[0]

        if not machine:
            unknown_machine_cqs.append((name, id_obj, result_date))
            continue

        real_cqs.append({
            "Machine": machine,
            "Date": result_date
        })

    df_cqs = pd.DataFrame(real_cqs)
    machines_cqs = sorted(df_cqs['Machine'].unique())
    cqs_data = []

    for _, sem in df_semestres.iterrows():
        row = {"Semestre": sem["Semestre"], "Year": sem["Year"]}
        for m in machines_cqs:
            done = df_cqs[
                (df_cqs["Machine"] == m) &
                (df_cqs["Date"] >= sem["DateDebut"]) &
                (df_cqs["Date"] <= sem["DateFin"])
            ]
            if not done.empty:
                row[m] = "✅"
            elif sem["DateFin"] < today:
                row[m] = "❌"
            else:
                row[m] = "⏳"
        cqs_data.append(row)

    df_cqs_final = pd.DataFrame(cqs_data)
    commentaires = get_commentaires()


    years = sorted(df_cqh_final["Year"].unique())

    #a personaliser a souhait avec logo du centre, préférence de police, disposition etc..
    return render_template_string("""
    <!DOCTYPE html>
    <html>
    <head>
      <title>Dashboard CQ</title>
      <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@4.6.2/dist/css/bootstrap.min.css">
      <style>
        td.success { background-color: #d4edda; }
        td.danger  { background-color: #f8d7da; }
        td { text-align: center; }
      </style>
    </head>
    <body>
      <div class="container mt-4">
        <h2>📋 Dashboard CQ 2025</h2>
        <ul class="nav nav-tabs" id="cqTab" role="tablist">
          <li class="nav-item">
            <a class="nav-link active" id="cqh-tab" data-toggle="tab" href="#cqh" role="tab">CQH Hebdomadaire</a>
          </li>
          <li class="nav-item">
            <a class="nav-link" id="cqm-tab" data-toggle="tab" href="#cqm" role="tab">CQM Mensuel</a>
          </li>
          <li class="nav-item">
            <a class="nav-link" id="cqs-tab" data-toggle="tab" href="#cqs" role="tab">CQS Semestriel</a>
          </li>
        </ul>

                                  
        <div style="margin-bottom: 16px;">
        <label for="yearSelectGlobal"><b>Année :</b></label>
        <select id="yearSelectGlobal" class="form-control" style="width:auto; display:inline-block;">
            <option value="2024">2024</option>
            <option value="2025" selected>2025</option>
        </select>
        </div>


                                
        <div class="tab-content mt-3">
            <div class="tab-pane fade show active" id="cqh" role="tabpanel">
                <h5>CQH (par semaine)</h5>
                <a href="/export_cqh_csv" class="btn btn-success btn-sm mb-2">
                    ⬇ Télécharger le tableau CQH (Excel/CSV)
                </a>

                <table class="table table-bordered table-sm">
                    <thead>
                        <tr>
                            <th>Semaine</th>
                            <th>Year</th>
                            {% for col in df_cqh.columns if col not in ['Semaine', 'Year'] %}
                                <th>{{ col }}</th>
                            {% endfor %}
                        </tr>
                    </thead>
                    <tbody id="table-cqh-body">
                        {% for _, row in df_cqh.iterrows() %}
                        <tr data-year="{{ row['Year'] }}">
                            <td>{{ row["Semaine"] }}</td>
                            <td>{{ row["Year"] }}</td>
                            {% for machine in df_cqh.columns if machine not in ['Semaine', 'Year'] %}
                                {% set val = row[machine] %}
                                <td class="{{ 'success' if val == '✅' else 'danger' if val == '❌' else '' }}">
                                    {{ val }}
                                    {% if val == "❌" %}
                                        {% set c = commentaires.get((machine, row['Semaine'])) %}
                                        {% if c %}
                                            <!-- Si commentaire existe, badge coloré + tooltip -->
                                            <span 
                                                class="badge badge-info"
                                                data-toggle="tooltip"
                                                data-placement="top"
                                                style="cursor:pointer;"
                                                title="{{ c['commentaire'] }} ({{ c['auteur'] }})"
                                                onclick="ouvrirCommentaire('{{ machine }}', '{{ row['Semaine'] }}')">
                                                💬
                                            </span>
                                        {% else %}
                                            <!-- Sinon, bouton discret pour ajouter -->
                                            <button class="btn btn-link btn-sm p-0"
                                                title="Ajouter un commentaire"
                                                onclick="ouvrirCommentaire('{{ machine }}', '{{ row['Semaine'] }}')">
                                                💬
                                            </button>
                                        {% endif %}
                                    {% endif %}
                                </td>
                            {% endfor %}
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>




            <div class="tab-pane fade" id="cqm" role="tabpanel">
            <h5>CQM (par mois)</h5>
            <table class="table table-bordered table-sm">
                <thead>
                <tr>
                    <th>Mois</th>
                    <th>Year</th>
                    {% for col in df_cqm.columns if col not in ['Mois', 'Year'] %}
                    <th>{{ col }}</th>
                    {% endfor %}
                </tr>
                </thead>
                <tbody>
                {% for _, row in df_cqm.iterrows() %}
                <tr data-year="{{ row['Year'] }}">
                    <td>{{ row["Mois"] }}</td>
                    <td>{{ row["Year"] }}</td>
                    {% for machine in df_cqm.columns if machine not in ['Mois', 'Year'] %}
                        {% set val = row[machine] %}
                        <td class="{{ 'success' if val == '✅' else 'danger' if val == '❌' else '' }}">
                            {{ val }}
                            {% if val == "❌" %}
                                {% set c = commentaires.get((machine, row['Mois'])) %}
                                {% if c %}
                                    <span 
                                        class="badge badge-info"
                                        data-toggle="tooltip"
                                        data-placement="top"
                                        style="cursor:pointer;"
                                        title="{{ c['commentaire'] }} ({{ c['auteur'] }})"
                                        onclick="ouvrirCommentaire('{{ machine }}', '{{ row['Mois'] }}')">
                                        💬
                                    </span>
                                {% else %}
                                    <button class="btn btn-link btn-sm p-0"
                                        title="Ajouter un commentaire"
                                        onclick="ouvrirCommentaire('{{ machine }}', '{{ row['Mois'] }}')">
                                        💬
                                    </button>
                                {% endif %}
                            {% endif %}
                        </td>
                    {% endfor %}
                </tr>
                {% endfor %}
                </tbody>

            </table>
            </div>


        <div class="tab-pane fade" id="cqs" role="tabpanel">
        <h5>CQS (par semestre)</h5>
        <table class="table table-bordered table-sm">
            <thead>
            <tr>
                <th>Semestre</th>
                <th>Year</th>
                {% for col in df_cqs.columns if col not in ['Semestre', 'Year'] %}
                <th>{{ col }}</th>
                {% endfor %}
            </tr>
            </thead>
            <tbody>
            {% for _, row in df_cqs.iterrows() %}
            <tr data-year="{{ row['Year'] }}">
                <td>{{ row["Semestre"] }}</td>
                <td>{{ row["Year"] }}</td>
                {% for machine in df_cqs.columns if machine not in ['Semestre', 'Year'] %}
                    {% set val = row[machine] %}
                    <td class="{{ 'success' if val == '✅' else 'danger' if val == '❌' else '' }}">
                        {{ val }}
                        {% if val == "❌" %}
                            {% set c = commentaires.get((machine, row['Semestre'])) %}
                            {% if c %}
                                <span 
                                    class="badge badge-info"
                                    data-toggle="tooltip"
                                    data-placement="top"
                                    style="cursor:pointer;"
                                    title="{{ c['commentaire'] }} ({{ c['auteur'] }})"
                                    onclick="ouvrirCommentaire('{{ machine }}', '{{ row['Semestre'] }}')">
                                    💬
                                </span>
                            {% else %}
                                <button class="btn btn-link btn-sm p-0"
                                    title="Ajouter un commentaire"
                                    onclick="ouvrirCommentaire('{{ machine }}', '{{ row['Semestre'] }}')">
                                    💬
                                </button>
                            {% endif %}
                        {% endif %}
                    </td>
                {% endfor %}
            </tr>
            {% endfor %}
            </tbody>

        </table>
        </div>


        <a href="/" class="btn btn-secondary mt-3">⬅ Retour</a>
      </div>

      <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
      <script src="https://cdn.jsdelivr.net/npm/bootstrap@4.6.2/dist/js/bootstrap.bundle.min.js"></script>
                                  
            <!-- Modale pour ajouter un commentaire -->
        <div class="modal fade" id="modalCommentaire" tabindex="-1">
        <div class="modal-dialog">
            <form class="modal-content" onsubmit="submitCommentaire(); return false;">
            <div class="modal-header">
                <h5 class="modal-title">Ajouter un commentaire</h5>
                <button type="button" class="close" data-dismiss="modal">&times;</button>
            </div>
            <div class="modal-body">
                <input type="hidden" id="modal_machine">
                <input type="hidden" id="modal_semaine">
                <div class="form-group">
                <label>Commentaire :</label>
                <textarea class="form-control" id="modal_commentaire" required></textarea>
                </div>
                <div class="form-group">
                <label>Votre nom :</label>
                <input type="text" class="form-control" id="modal_auteur" required>
                </div>
            </div>
            <div class="modal-footer">
                <button class="btn btn-primary" type="submit">Enregistrer</button>
            </div>
            </form>
        </div>
        </div>
        <script>
        function ouvrirCommentaire(machine, semaine) {
        $('#modal_machine').val(machine);
        $('#modal_semaine').val(semaine);
        $('#modal_commentaire').val('');
        $('#modal_auteur').val('');
        $('#modalCommentaire').modal('show');
        }

        function submitCommentaire() {
        fetch('/ajoute_commentaire', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({
            machine: $('#modal_machine').val(),
            semaine: $('#modal_semaine').val(),
            commentaire: $('#modal_commentaire').val(),
            auteur: $('#modal_auteur').val()
            })
        }).then(response => {
            $('#modalCommentaire').modal('hide');
            location.reload();
        });
        }
        </script>
        <script>
        $(function () {
        $('[data-toggle="tooltip"]').tooltip()
        })
        </script>

        <script>
        function filtreParAnnee(annee) {
        // Toutes les tables qui ont des lignes data-year
        document.querySelectorAll('tbody tr[data-year]').forEach(function(tr) {
            tr.style.display = (tr.getAttribute('data-year') == annee) ? '' : 'none';
        });
        }

        document.addEventListener('DOMContentLoaded', function() {
        var sel = document.getElementById('yearSelectGlobal');
        if(sel) {
            filtreParAnnee(sel.value); // Initial
            sel.addEventListener('change', function() {
            filtreParAnnee(this.value);
            });
        }
        });
        </script>

    </body>
    </html>
    """, df_cqh=df_cqh_final, df_cqm=df_cqm_final, df_cqs=df_cqs_final, commentaires=commentaires, years=years)

@app.route('/ajoute_commentaire', methods=['POST'])
def ajoute_commentaire():
    data = request.json
    machine = data['machine']
    semaine = data['semaine']
    commentaire = data['commentaire']
    auteur = data.get('auteur', 'inconnu')
    conn = sqlite3.connect("commentaires_cq.db")
    c = conn.cursor()
    c.execute("INSERT INTO commentaires (machine, semaine, commentaire, auteur) VALUES (?, ?, ?, ?)",
              (machine, semaine, commentaire, auteur))
    conn.commit()
    conn.close()
    return jsonify({"status": "ok"})


@app.route("/audit_machines")
def audit_machines():
    if not UNKNOWN_MACHINES_CQH:
        return "<p>Aucune machine inconnue détectée.</p>"

    df = pd.DataFrame(UNKNOWN_MACHINES_CQH)
    html = df[["Name", "Id_Object", "StudyDate"]].drop_duplicates().to_html(index=False)
    return f"<h2>Machines non reconnues</h2>{html}"

if __name__ == "__main__":

    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    WEBHOOK_URL = "https://exemple.ezgeghguzeguze4654654ef6ezgf"

    def send_teams_alert_cqh(machines_en_retard, week_num):
        if not machines_en_retard:
            return  # rien à envoyer
        txt = (
            f"🚨 **Alerte CQH hebdo**\n\n"
            f"Les CQH suivants n'ont pas été réalisés pour la semaine {week_num} :\n"
            + "\n".join(f"• {m}" for m in machines_en_retard) +
            "\nMerci de vérifier avant la clôture de la semaine !"
        )
        payload = {"text": txt}
        try:
            r = requests.post(WEBHOOK_URL, json=payload, verify=False)
            print(f"[CQH] Notification Teams envoyée. Status: {r.status_code}")
        except Exception as e:
            print(f"[CQH] Erreur d'envoi Teams : {e}")

    def verif_cqh_et_alerte():
        print("🔔 [TEST] Exécution de la tâche automatique CQH")
        try:
            # Utilise la même logique que /cq_dashboard pour être cohérent
            # On récupère le tableau CQH semaine/machine déjà généré dans get_taux_conformite()
            today = date.today()
            machines, taux_cqh, _, _ = get_taux_conformite()
            
            # 1. On veut le DataFrame semaine-machine ("df_cqh_final" dans /cq_dashboard)
            # On régénère le DataFrame comme dans la fonction
            # Générer les semaines complètes pour 2024 et 2025 (lundi-vendredi)
            semaines = []
            current = date(2023, 12, 25)  # pour être sûr d'inclure S1 de 2024
            current -= timedelta(days=current.weekday())
            while current.year < 2026:
                y, w, _ = current.isocalendar()
                if y >= 2024:
                    semaines.append({
                        "Semaine": f"S{w}",
                        "DateDebut": current,
                        "DateFin": current + timedelta(days=4)
                    })
                current += timedelta(weeks=1)
            df_semaines = pd.DataFrame(semaines)


            # Requête SQL simplifiée
            conn = pyodbc.connect(conn_str, timeout=5)
            rows = pd.read_sql("""
                SELECT cs.Id_Object, cs.Id_UserModule, cs.Name, cs.StudyDate
                FROM CONTROLE_STUDY cs
                JOIN RESULT r ON cs.Id_ControleStudy = r.Id_ControleStudy
                WHERE cs.StudyDate IS NOT NULL
            """, conn)
            conn.close()
            rows['StudyDate'] = pd.to_datetime(rows['StudyDate'])

            MODULES_PAR_TYPE = {
                "CQH": {66, 64, 63, 62, 61, 60},
            }

            real_cqh = []
            for _, row in rows.iterrows():
                module_id = row["Id_UserModule"]
                id_obj = row["Id_Object"]
                name = row["Name"]
                study_date = row["StudyDate"]
                if id_obj in [99, 121]:  # TOMO1 ou TOMO2
                    if not CQH_REGEX.search(str(name)):
                        continue
                else:
                    if module_id not in MODULES_PAR_TYPE["CQH"]:
                        continue
                if id_obj not in MACHINES: continue
                if pd.isna(study_date): continue
                study_date = pd.to_datetime(study_date).date()
                real_cqh.append({"Id_Object": id_obj, "Date": study_date})

            df_cqh = pd.DataFrame(real_cqh)

            machines_cqh = sorted(df_cqh['Id_Object'].unique())
            cqh_data = []
            for _, sem in df_semaines.iterrows():
                row = {"Semaine": sem["Semaine"], "Year": sem["DateDebut"].year}
                for id_obj in MACHINES.keys():
                    machine_name = MACHINES.get(id_obj, ("❓",))[0]
                    done = df_cqh[
                        (df_cqh["Id_Object"] == id_obj) &
                        (df_cqh["Date"] >= sem["DateDebut"]) &
                        (df_cqh["Date"] <= sem["DateFin"])
                    ]
                    if not done.empty:
                        row[machine_name] = "✅"
                    elif sem["DateFin"] < today:
                        row[machine_name] = "❌"
                    else:
                        row[machine_name] = "⏳"
                cqh_data.append(row)
            df_cqh_final = pd.DataFrame(cqh_data)

            # 2. Repère la semaine en cours
            week_num = today.isocalendar()[1]
            week_label = f"S{week_num}"
            # On prend la ligne de la semaine en cours
            current_week_row = df_cqh_final[df_cqh_final["Semaine"] == week_label]
            if current_week_row.empty:
                #print("⚠️ Semaine en cours introuvable dans le tableau CQH.")
                return

            row = current_week_row.iloc[0]
            print("Colonnes DataFrame CQH :", list(df_cqh_final.columns))
            print("MACHINES dans la boucle d’alerte :", [m[0] for m in MACHINES.values()])
            # Liste des machines non faites (pas '✅') cette semaine
            machines_en_retard = [m for m in MACHINES.values() if row.get(m[0]) != "✅"]
            machines_names = [m[0] for m in machines_en_retard]

            if machines_names:
                #print(f"[CQH] Machines sans CQH cette semaine ({week_num}) :", machines_names)
                send_teams_alert_cqh(machines_names, week_num)
            else:
                print(f"[CQH] Tous les CQH sont faits pour la semaine {week_num}")
        except Exception as e:
            print(f"[CQH] Erreur lors du contrôle hebdo : {e}")

    scheduler = BackgroundScheduler()
    scheduler.add_job(verif_cqh_et_alerte, 'cron', day_of_week='wed', hour=16, minute=0)
    scheduler.add_job(verif_cqh_et_alerte, 'cron', day_of_week='fri', hour=16, minute=0)
    scheduler.start()


    app.run(host="0.0.0.0", port=5000)

