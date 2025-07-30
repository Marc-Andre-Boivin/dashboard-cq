
# config.py
# Paramètres de connexion SQL (sans identifiants)
SQL_CONFIG = {
    "server": "31.10.XX.XXX,XXXX", #renseigner l'IP du serveur (retrouvable dans les fichiers de parametrages d'export image vers ArtiscanDB sur un linac )
    "database": "DBArtiscan"
}

# Dictionnaire des machines (ID: (Nom, Couleur)) a retrouver dans la partie inventaire sur artiscan, a adapter au centre
MACHINES = {
    145: ("Versa HD 3", "#1976D2"),
    182: ("Versa HD 2", "#FF9800"),
    99:  ("TOMO1", "#388E3C"),
    121: ("TOMO2", "#8E24AA"),
    177: ("Versa HD 4", "#FBC02D"),
    162: ("Versa HD 5", "#D32F2F"),
    159: ("Versa HD 1", "#0288D1"),
    25:  ("NOVALIS", "#F57C00"),
}

# Regex, a adapter a votre nomenclature des CQ
EXCLUSION_REGEX = r"(?i)\\b(test|à supp|a supp|à supprimer|a supprimer|essai)\\b"
CQM_REGEX = r"(?i)cqm|controle ?qualite ?mensuel|contrôle ?qualité ?mensuel|controlequalitemensuel|contrôlequalitemensuel"
CQH_REGEX = r"(?i)cqh|controle ?qualite ?hebdo|contrôle ?qualité ?hebdo|controlequalitehebdomadaire|contrôlequalitéhebdomadaire"
CQS_REGEX = r"(?i)cqs|controle ?qualite ?semestriel|contrôle ?qualité ?semestriel|controlequalitesemestriel|contrôlequalitesemestriel"

# Modules par type, a adapter et retrouver dans Protocole -> Gerer sur Artiscan
MODULES_PAR_TYPE = {
    "CQH": {66, 64, 63, 62, 61, 60},
    "CQM": {97, 98, 92, 90, 105, 104, 103},
    "CQS": {96, 95, 94, 93, 91, 106},
}

# Base de données pour les commentaires, peut etre laisser comme ca par defaut
COMMENT_DB = "commentaires_cq.db"

# Webhook Teams, seulement si utilisation de la fonction notification
WEBHOOK_URL = "https://prod-07.francecentral.logic.azure.com/..."
