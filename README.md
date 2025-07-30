# dashboard-cq
 Documentation - Dashboard CQ (Contrôle Qualité)

Ce document décrit le fonctionnement global du script `dashboard_cq_artiscan.py` et les prérequis pour son déploiement.
Le script a été élaboré pour l'Institution Gustave Roussy et dois être personalisé pour fonctionner dans un autre centre.

---

## 1. Objectif du Dashboard

Le dashboard a pour but de suivre la réalisation des Contrôles Qualité (CQ) périodiques (hebdomadaires, mensuels, semestriels) pour les centre de radiothérapie utilisant Artiscan.
Il récupère les données depuis la base SQL Artiscan  et affiche l'état de conformité des CQ dans une interface web interactive.

---

## 2. Fonctionnement global

### 2.1. Récupération des données
- Le script se connecte à la base SQL via **pyodbc**.
- Les requêtes SQL interrogent les tables `CONTROLE_STUDY` et `RESULT`.
- Les résultats sont analysés pour déterminer si les CQH, CQM et CQS ont été effectués dans les périodes prévues.

### 2.2. Génération des tableaux
- **CQH (hebdomadaire)** : suivi par semaine ISO (lundi-vendredi).
- **CQM (mensuel)** : suivi par mois.
- **CQS (semestriel)** : suivi par semestre (S1/S2).

Chaque machine est marquée comme :
- ✅ si le CQ a été effectué,
- ❌ si la période est passée sans CQ,
- ⏳ si la période est en cours/non-passées.

### 2.3. Interface web (Flask)
- Le script lance un serveur Flask sur le port 5000 par défaut.
- L’URL `/` affiche une page avec :
    - Le calendrier des CQ (FullCalendar).
    - Les taux de conformité par machine.
    - Un tableau de suivi détaillé (via `/cq_dashboard`).

### 2.4. Commentaires
- Une base SQLite (`commentaires_cq.db`) stocke des notes/commentaires pour expliquer un retard ou une absence de CQ.

### 2.5. Alertes Teams
- Une notification peut être envoyée via un **webhook Teams** si des CQH ne sont pas réalisés (configurable dans `config.py`).

---

## 3. Pré-requis techniques

### 3.1. Environnement Python
- **Python 3.8+** recommandé.
- Modules nécessaires : 
  ```bash
  pip install flask pyodbc pandas apscheduler requests
  ```

### 3.2. Connexion SQL
- Accès réseau au serveur SQL.
- **ODBC Driver for SQL Server** installé (par ex. "ODBC Driver 17 for SQL Server") -- pas forcément nécessaire.
- Identifiant et mot de passe SQL valides (saisis au lancement du script).

### 3.3. Autres
- Navigateur web (Chrome, Edge ou Firefox) pour visualiser le dashboard.
- Accès au port 5000 (configurable dans `dashboard_cq_artiscan.py`).

---

## 4. Lancement du script

1. Configurer `config.py` (server, database, machines, regex, etc.).
2. Exécuter :
   ```bash
   python dashboard_cq_artiscan.py
   ```
3. Entrer l’identifiant SQL et le mot de passe.
4. Accéder au dashboard via : 
   ```
   http://<serveur>:5000/
   ```

---

## 5. Personnalisation pour un autre centre

Pour adapter le script à un autre centre :
- Modifier **SQL_CONFIG** (server, database) dans `config.py`.
- Mettre à jour la liste **MACHINES** (ID machine et noms).
- Adapter les **regex** si les noms des tests diffèrent (CQM, CQH, CQS).
- Ajuster **MODULES_PAR_TYPE** avec les IDs des modules spécifiques au centre.
- Adapter l'ID des modules / protocoles et noms des machines dans le config.py ET le code principal.
- Créer un dossier à la racine du script nommé "static" et y insérer dedans le logo du centre en remplacant également le nom du fichier logo dans le code principal.
---

## 6. Sécurité

- Les identifiants SQL ne sont pas stockés : ils sont saisis au lancement.
- Pour protéger l'accès web, on peut ajouter une authentification HTTP (optionnelle).

Pour toute demande d'information, ne pas hésiter à envoyer un email à marcandre.boivin@gustaveroussy.fr

