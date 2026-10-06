# -*- coding: utf-8 -*-
"""Cœur métier de la cartographie des biens de l'État.

Ce module ne dépend d'aucun framework d'interface : il contient les règles de
gestion, le nettoyage des données, l'audit, les indicateurs et la construction
des figures Plotly. Il est partagé par les deux interfaces livrées :

- `app7.py`    — interface Streamlit ;
- `app_dash.py` — interface Dash.

Toute règle métier se modifie donc à un seul endroit. Les caches (Streamlit
`st.cache_data`, Flask-Caching ou `functools.lru_cache` côté Dash) sont
appliqués par chaque interface, jamais ici.

Règles de gestion principales :
- valeur du patrimoine = valeur nette comptable, la valeur d'acquisition
  restant affichée à côté ;
- Administration centrale / Administration DÉCONCENTRÉE (jamais « décentralisée ») ;
- À réformer ≠ Réformé ;
- immobilisation si prix unitaire >= 500 000 FCFA (seuil sur le prix UNITAIRE) ;
- état du bien non renseigné imputé à « Bon état », imputation tracée.
"""

import difflib
import io
import json
import re
import unicodedata
import warnings

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# Couleurs ré-exportées telles quelles vers les deux interfaces : elles ne
# sont pas utilisées ici, mais `patrimoine_core` est le point d'entrée unique
# de la charte pour app7.py et app_dash.py.
from visual_theme import (
    COULEURS_ADMIN,
    COULEURS_CLASSIF,
    COULEURS_ETAT,
    COULEURS_MINISTERE,
    COULEURS_PRIORITE,
    COULEURS_RISQUE,
    THEME,
)

warnings.filterwarnings(
    "ignore",
    message="Data Validation extension is not supported and will be removed",
    category=UserWarning,
    module="openpyxl",
)

# =============================================================================
# 1. PARAMÈTRES MÉTIER ET VISUELS
# =============================================================================

SEUIL_IMMOBILISATION_FCFA = 500_000

# Règle métier v6 : un bien dont l'état n'a pas été renseigné lors du
# recensement est réputé en « Bon état ». La modalité « Non renseigné »
# disparaît donc des états, mais l'imputation reste tracée (colonne
# « État imputé ») pour que la qualité de la collecte reste mesurable.
ETAT_PAR_DEFAUT = "Bon état"

ORDRE_ETATS = ["Neuf", "Bon état", "À réformer", "Réformé"]

MAP_ETAT = {
    "NEUF": "Neuf",
    "BON": "Bon état",
    "BON ETAT": "Bon état",
    "BONNE ETAT": "Bon état",
    "ASSEZ BON": "Bon état",
    "ASSEZ BON ETAT": "Bon état",
    "A REFORMER": "À réformer",
    "A REFORME": "À réformer",
    "REFORME": "Réformé",
    "REFORMER": "Réformé",
}

SCORE_SANTE = {"Neuf": 100, "Bon état": 70, "À réformer": 20, "Réformé": 0}

# -----------------------------------------------------------------------------
# Référentiel des ministères et institutions constitutionnelles.
# La base pilote ne contient que le MINTP et le MINAC ; le référentiel permet
# d'ouvrir la liste déroulante à l'ensemble du périmètre de l'État sans
# retoucher le code lorsque de nouvelles administrations seront recensées.
# Chaque entrée : sigle -> libellé complet.
# -----------------------------------------------------------------------------
REFERENTIEL_MINISTERES = {
    "PRC": "Présidence de la République",
    "SPM": "Services du Premier Ministre",
    "AN": "Assemblée nationale",
    "SENAT": "Sénat",
    "CS": "Cour suprême",
    "CC": "Conseil constitutionnel",
    "CONSUPE": "Contrôle supérieur de l'État",
    "CES": "Conseil économique et social",
    "ELECAM": "Elections Cameroon",
    "CNDHL": "Commission des droits de l'homme et des libertés",
    "MINADER": "Ministère de l'Agriculture et du Développement rural",
    "MINAC": "Ministère des Arts et de la Culture",
    "MINAS": "Ministère des Affaires sociales",
    "MINATD": "Ministère de l'Administration territoriale",
    "MINCOM": "Ministère de la Communication",
    "MINCOMMERCE": "Ministère du Commerce",
    "MINDCAF": "Ministère des Domaines, du Cadastre et des Affaires foncières",
    "MINDDEVEL": "Ministère de la Décentralisation et du Développement local",
    "MINDEF": "Ministère de la Défense",
    "MINEDUB": "Ministère de l'Éducation de base",
    "MINEE": "Ministère de l'Eau et de l'Énergie",
    "MINEFOP": "Ministère de l'Emploi et de la Formation professionnelle",
    "MINEPAT": "Ministère de l'Économie, de la Planification et de l'Aménagement du territoire",
    "MINEPDED": "Ministère de l'Environnement, de la Protection de la nature et du Développement durable",
    "MINEPIA": "Ministère de l'Élevage, des Pêches et des Industries animales",
    "MINESEC": "Ministère des Enseignements secondaires",
    "MINESUP": "Ministère de l'Enseignement supérieur",
    "MINFI": "Ministère des Finances",
    "MINFOF": "Ministère des Forêts et de la Faune",
    "MINFOPRA": "Ministère de la Fonction publique et de la Réforme administrative",
    "MINHDU": "Ministère de l'Habitat et du Développement urbain",
    "MINJEC": "Ministère de la Jeunesse et de l'Éducation civique",
    "MINJUSTICE": "Ministère de la Justice",
    "MINMAP": "Ministère des Marchés publics",
    "MINMIDT": "Ministère des Mines, de l'Industrie et du Développement technologique",
    "MINPMEESA": "Ministère des PME, de l'Économie sociale et de l'Artisanat",
    "MINPOSTEL": "Ministère des Postes et Télécommunications",
    "MINPROFF": "Ministère de la Promotion de la femme et de la Famille",
    "MINREX": "Ministère des Relations extérieures",
    "MINRESI": "Ministère de la Recherche scientifique et de l'Innovation",
    "MINSANTE": "Ministère de la Santé publique",
    "MINSEP": "Ministère des Sports et de l'Éducation physique",
    "MINT": "Ministère des Transports",
    "MINTOUL": "Ministère du Tourisme et des Loisirs",
    "MINTP": "Ministère des Travaux publics",
    "MINTSS": "Ministère du Travail et de la Sécurité sociale",
    "DGSN": "Délégation générale à la Sûreté nationale",
}

# Harmonisation du champ « NATURE DU MATERIEL » (doublons d'espaces et
# variantes de saisie constatés dans le classeur source).
MAP_NATURE_MATERIEL = {
    "MOBILIER DE BUREAU": "Mobilier de bureau",
    "MOBILIER DE BUREAU DESK": "Mobilier de bureau",
    "MATERIEL INFORMATIQUE": "Matériel informatique",
    "MATERIEL BUREAUTIQUE": "Matériel bureautique",
    "MATERIEL ELECTROMENAGER": "Matériel électroménager",
    "MATERIEL DE TRANSPORT": "Matériel de transport",
    "MATERIEL DE TELECOMMUNICATION ET AUDIO VISUEL": "Matériel de télécommunication et audiovisuel",
    "MATERIEL DE TELECOMMUNICATION ET AUDIOVISUEL": "Matériel de télécommunication et audiovisuel",
    "MATERIEL DE DECORATION": "Matériel de décoration",
    "MATERIEL ELECTRIQUE": "Matériel électrique",
    "MATERIEL D HEBERGEMENT": "Matériel d'hébergement",
    "MATERIEL MUSICAL": "Matériel musical",
    "MATERIEL D ENTRETIEN": "Matériel d'entretien",
    "MATERIEL DE SECURITE": "Matériel de sécurité",
}

MAP_ADMINISTRATION = {
    "ADMINISTRATION GENERALE": "Centrale",
    "ADMINISTRATION CENTRALE": "Centrale",
    "CENTRALE": "Centrale",
    "ADMINISTRATION DECONCENTREE": "Déconcentrée",
    "DECONCENTREE": "Déconcentrée",
    "ADMINISTRATION DECONCENTRE": "Déconcentrée",
    "DECONCENTRE": "Déconcentrée",
    # Si la base contient cette saisie par erreur, on la rattache aux services déconcentrés.
    "ADMINISTRATION DECENTRALISEE": "Déconcentrée",
    "DECENTRALISEE": "Déconcentrée",
}

# Uniformisation du mode d'acquisition sur toutes les feuilles chargées.
MAP_MODE_ACQUISITION = {
    "ACQUISITION": "Acquisition", "ACHAT": "Acquisition", "ACHATS": "Acquisition",
    "ACQUIS": "Acquisition", "ACQUISE": "Acquisition",
    "BILAN D OUVERTURE": "Bilan d'ouverture", "BILAN OUVERTURE": "Bilan d'ouverture",
    "BILAN DOUVERTURE": "Bilan d'ouverture",
    "DON OU LEG": "Don ou legs", "DON OU LEGS": "Don ou legs",
    "DONS OU LEGS": "Don ou legs", "DON": "Don ou legs",
    "LEGS": "Don ou legs", "LEG": "Don ou legs",
}
ORDRE_MODES_ACQUISITION = ["Acquisition", "Bilan d'ouverture", "Don ou legs",
                           "Non renseigné"]

MAP_REGION = {
    "EXTREME NORD": "Extrême-Nord", "EXTREMENORD": "Extrême-Nord",
    "NORD": "Nord",
    "ADAMOUA": "Adamaoua", "ADAMAOUA": "Adamaoua",
    "CENTRE": "Centre", "EST": "Est", "LITTORAL": "Littoral", "OUEST": "Ouest",
    "NORD OUEST": "Nord-Ouest", "NORDOUEST": "Nord-Ouest",
    "SUD OUEST": "Sud-Ouest", "SUDOUEST": "Sud-Ouest",
    "SUD": "Sud",
}

COORD_REGIONS = {
    "Adamaoua": (7.3167, 13.5833),
    "Centre": (3.8480, 11.5021),
    "Est": (4.5766, 13.6843),
    "Extrême-Nord": (10.5908, 14.3161),
    "Littoral": (4.0511, 9.7679),
    "Nord": (9.3000, 13.3833),
    "Nord-Ouest": (5.9597, 10.1461),
    "Ouest": (5.4737, 10.4179),
    "Sud": (2.9000, 11.1500),
    "Sud-Ouest": (4.1597, 9.2442),
}

ORDRE_REGIONS = [
    "Adamaoua", "Centre", "Est", "Extrême-Nord", "Littoral",
    "Nord", "Nord-Ouest", "Ouest", "Sud", "Sud-Ouest",
]

TOUT_LE_CAMEROUN = "Tout le Cameroun"
GLOBAL_MINISTERES = "Tous les ministères (global)"

# Fonds administratifs : geoBoundaries CMR (licence CC-BY 3.0), stockés
# localement dans assets/ pour que l'application fonctionne sans réseau.
# ADM1 = 10 régions, ADM2 = 58 départements, ADM3 = 360 arrondissements.
FICHIER_GEOJSON_REGIONS = "regions_cameroun.geojson"
FICHIER_GEOJSON_DEPARTEMENTS = "departements_cameroun.geojson"
FICHIER_GEOJSON_ARRONDISSEMENTS = "arrondissements_cameroun.geojson"

NOMS_REGIONS_GEOJSON = {
    "Adamaoua": "Adamaoua",
    "Centre": "Centre",
    "East": "Est",
    "Far North": "Extrême-Nord",
    "Littoral": "Littoral",
    "North": "Nord",
    "North-West": "Nord-Ouest",
    "West": "Ouest",
    "South": "Sud",
    "South-West": "Sud-Ouest",
}

NIVEAU_REGION = "Régions"
NIVEAU_DEPARTEMENT = "Départements"
NIVEAU_ARRONDISSEMENT = "Arrondissements"

# Variantes orthographiques entre le classeur et le fond administratif, que la
# normalisation seule ne rapproche pas (lettre finale, apostrophe).
CORRECTIONS_DEPARTEMENTS = {
    "BAMBOUTOUS": "Bamboutos",
    "KADEY": "Kadei",
    "NYONG ET SO O": "Nyong-et-So",
}

# Chaque niveau : fichier du fond, colonne de la base qui le renseigne, et
# libellé au singulier pour les textes d'interface.
NIVEAUX_ADMIN = {
    # « traduction » renomme les entités du fond (l'ADM1 de geoBoundaries est
    # en anglais) ; « corrections » rapproche les libellés de la base du fond.
    NIVEAU_REGION: {"fichier": FICHIER_GEOJSON_REGIONS, "colonne": "Région",
                    "singulier": "région", "feminin": True, "corrections": {},
                    "traduction": NOMS_REGIONS_GEOJSON},
    NIVEAU_DEPARTEMENT: {"fichier": FICHIER_GEOJSON_DEPARTEMENTS,
                         "colonne": "Département", "singulier": "département",
                         "feminin": False, "corrections": CORRECTIONS_DEPARTEMENTS,
                         "traduction": {}},
    NIVEAU_ARRONDISSEMENT: {"fichier": FICHIER_GEOJSON_ARRONDISSEMENTS,
                            "colonne": "Arrondissement", "singulier": "arrondissement",
                            "feminin": False, "corrections": {}, "traduction": {}},
}


# Couleur propre à chaque région pour la vue nationale « une région =
# une couleur ». Teintes qualitatives distinctes, contraste >= 3:1 sur fond
# clair, et suffisamment éloignées deux à deux pour rester lisibles côte à
# côte (les régions voisines n'ont jamais des teintes proches).
COULEURS_REGIONS = {
    "Adamaoua": "#1F6F54",
    "Centre": "#8A6508",
    "Est": "#2C6E91",
    "Extrême-Nord": "#8C3B72",
    "Littoral": "#0F7C8A",
    "Nord": "#B04A1A",
    "Nord-Ouest": "#5B4B9E",
    "Ouest": "#B8860B",
    "Sud": "#3C7A24",
    "Sud-Ouest": "#A03B3B",
    "Non renseigné": "#6B7280",
}

# Échelle séquentielle pour la coloration « par valeur » (clair -> foncé :
# la valeur la plus forte est la plus sombre, lecture immédiate sans légende).
ECHELLE_CHOROPLETHE = [
    [0.00, "#F4EBCB"], [0.25, "#E0CE8F"], [0.50, "#C2A248"],
    [0.75, "#8A6508"], [1.00, "#4E3803"],
]

COORDS_VILLES = {
    "YAOUNDE": (3.8480, 11.5021), "YAOUNDE 1ER": (3.8480, 11.5021),
    "MBALMAYO": (3.5167, 11.5000), "BAFIA": (4.7500, 11.2333),
    "MONATELE": (4.2667, 11.2000), "MFOU": (3.7167, 11.6333),
    "AKONOLINGA": (3.7667, 12.2500), "ESEKA": (3.6500, 10.7667),
    "NTUI": (4.4500, 11.6333), "NGOUMOU": (3.5500, 11.4167),
    "NANGA EBOKO": (4.6833, 12.3667), "OBALA": (4.1667, 11.5333),
    "SOA": (3.9833, 11.6000),
    "DOUALA": (4.0511, 9.7679), "EDEA": (3.8000, 10.1333),
    "NKONGSAMBA": (4.9500, 9.9333), "YABASSI": (4.4500, 10.1667),
    "LOUM": (4.7167, 9.7333), "MANJO": (4.8417, 9.8222),
    "BAFOUSSAM": (5.4737, 10.4179), "BANGANGTE": (5.1500, 10.5167),
    "DSCHANG": (5.4500, 10.0667), "MBOUDA": (5.6333, 10.2500),
    "FOUMBAN": (5.7167, 10.9000), "FOUMBOT": (5.5167, 10.6500),
    "BAHAM": (5.3333, 10.5667), "BAFANG": (5.1500, 9.9833),
    "BANDJOUN": (5.3667, 10.4167),
    "BAYANGAM": (5.3167, 10.5333), "BAYAMGAM": (5.3167, 10.5333),
    "BAMENDA": (5.9597, 10.1461), "WUM": (6.3833, 10.0667),
    "NKAMBE": (6.6333, 10.6667), "MBENGWI": (6.0167, 10.0000),
    "KUMBO": (6.2000, 10.6667), "NDOP": (5.9667, 10.4333),
    "BUEA": (4.1597, 9.2442), "KUMBA": (4.6333, 9.4500),
    "LIMBE": (4.0167, 9.2000), "TIKO": (4.0833, 9.3667),
    "MAMFE": (5.7500, 9.3167), "MUNDEMBA": (4.9500, 8.8667),
    "BANGEM": (5.0500, 9.7500), "MENJI": (5.4667, 9.9167),
    "EBOLOWA": (2.9000, 11.1500), "AMBAM": (2.3833, 11.2833),
    "SANGMELIMA": (2.9333, 11.9833), "KRIBI": (2.9333, 9.9167),
    "DJOUM": (2.6667, 12.6667),
    "BERTOUA": (4.5766, 13.6843), "BATOURI": (4.4333, 14.3667),
    "ABONG MBANG": (3.9833, 13.1667), "ABONG BANG": (3.9833, 13.1667),
    "YOKADOUMA": (3.5167, 15.0500), "BETARE OYA": (5.5833, 14.0833),
    "NGUELEMENDOUKA": (4.0500, 13.3833),
    "NGAOUNDERE": (7.3167, 13.5833), "TIBATI": (6.4667, 12.6333),
    "MEIGANGA": (6.5167, 14.3000), "BANYO": (6.7500, 11.8167),
    "TIGNERE": (7.3667, 12.6500),
    "GAROUA": (9.3000, 13.3833), "GUIDER": (9.9333, 13.9500),
    "POLI": (8.4833, 13.2333), "TCHOLLIRE": (8.4000, 14.1667),
    "MAROUA": (10.5908, 14.3161), "KOUSSERI": (12.0833, 15.0333),
    "MORA": (11.0500, 14.1333), "KAELE": (10.1000, 14.4500),
    "YAGOUA": (10.3428, 15.2406), "MOKOLO": (10.7333, 13.8000),
}

MAP_LIBELLE = {
    "ORDINATEUR": "Ordinateur", "ORDINATEUR COMPLET": "Ordinateur",
    "MICRO ORDINATEUR": "Ordinateur", "ORDINATEUR DE BUREAU": "Ordinateur",
    "PC COMPLET": "Ordinateur", "COMPUTER": "Ordinateur",
    "LAPTOP": "Ordinateur portable", "ORDINATEUR PORTABLE": "Ordinateur portable",
    "UNITE CENTRALE": "Unité centrale", "ECRAN": "Écran", "MONITEUR": "Écran",
    "IMPRIMANTE": "Imprimante", "IMPRIMANTE HP": "Imprimante", "PRINTER": "Imprimante",
    "PHOTOCOPIEUR": "Photocopieur", "PHOTOCOPIEUSE": "Photocopieur", "COPIEUR": "Photocopieur",
    "SCANNER": "Scanner", "SCANER": "Scanner", "SCANNEUR": "Scanner",
    "ONDULEUR": "Onduleur", "REGULATEUR": "Onduleur", "REGULATEUR DE TENSION": "Onduleur",
    "SERVEUR": "Serveur", "SWITCH": "Équipement réseau", "ROUTEUR": "Équipement réseau",
    "MODEM": "Équipement réseau", "VIDEOPROJECTEUR": "Vidéoprojecteur", "VIDEO PROJECTEUR": "Vidéoprojecteur",
    "BUREAU": "Bureau", "BUREAU METALLIQUE": "Bureau", "BUREAU DIRECTEUR": "Bureau",
    "BUREAU EN BOIS": "Bureau", "TABLE": "Table", "TABLE DE REUNION": "Table",
    "CHAISE": "Chaise", "CHAIR": "Chaise", "FAUTEUIL": "Fauteuil",
    "FAUTEUIL DIRECTEUR": "Fauteuil", "CANAPE": "Salon / canapé", "SALON": "Salon / canapé",
    "ARMOIRE": "Armoire", "ARMOIRE METALLIQUE": "Armoire", "ETAGERE": "Étagère",
    "BIBLIOTHEQUE": "Bibliothèque", "COFFRE FORT": "Coffre-fort", "CLASSEUR": "Classeur",
    "MEUBLE": "Meuble", "CLIMATISEUR": "Climatiseur", "CLIMATISEUR SPLIT": "Climatiseur",
    "SPLIT": "Climatiseur", "REFRIGERATEUR": "Réfrigérateur", "REFRIGIRATEUR": "Réfrigérateur",
    "FRIGO": "Réfrigérateur", "FRIGIDAIRE": "Réfrigérateur", "FRIDGE": "Réfrigérateur",
    "TELEVISEUR": "Téléviseur", "TELEVISION": "Téléviseur", "TV": "Téléviseur",
    "TELEPHONE": "Téléphone", "POSTE TELEPHONIQUE": "Téléphone", "INTERPHONE": "Téléphone",
    "VENTILATEUR": "Ventilateur", "GROUPE ELECTROGENE": "Groupe électrogène", "GENERATOR": "Groupe électrogène",
    "VEHICULE": "Véhicule", "VEHICULE DE SERVICE": "Véhicule", "PICK UP": "Véhicule",
    "VOITURE": "Véhicule", "FORD": "Véhicule", "TOYOTA": "Véhicule",
    "MOTO": "Motocyclette", "MOTOCYCLETTE": "Motocyclette", "CAMION": "Camion",
    "RIDEAU": "Rideau / store", "RIDEAUX": "Rideau / store", "STORES": "Rideau / store",
    "PANIER": "Poubelle / panier", "POUBELLE": "Poubelle / panier", "CENDRIER": "Poubelle / panier",
    "EFFIGIE": "Effigie", "CARTE": "Carte murale", "CARTE DU CAMEROUN": "Carte murale",
    "DICTIONNAIRE": "Documentation", "RALLONGE": "Rallonge électrique", "AGRAFEUSE": "Petit matériel de bureau",
    "PERFORATEUR": "Petit matériel de bureau", "RELIEUSE": "Petit matériel de bureau",
    "CALCULATRICE": "Petit matériel de bureau", "DECAMETRE": "Matériel topographique", "GPS": "Matériel topographique",
    "THEODOLITE": "Matériel topographique", "TEODOLITRE": "Matériel topographique",
    "FAX": "Télécopieur", "FAXEUR": "Télécopieur", "APPAREIL PHOTO": "Appareil photo",
    "EXTINCTEUR": "Extincteur", "LIT": "Lit", "TABOURET": "Tabouret",
    "PENDULE": "Horloge", "CLOCK": "Horloge", "CUISINIERE": "Cuisinière", "FILTRE A EAU": "Filtre à eau",
    "MOUSE": "Souris", "SOURIS": "Souris", "CLAVIER": "Clavier",
    "TRACEUR": "Traceur", "DECODEUR": "Décodeur", "BAFFLE": "Enceinte / sonorisation",
}

REGLES_PREFIXE = [
    ("ORDINATEUR PORTABLE", "Ordinateur portable"), ("MICRO ORDINATEUR", "Ordinateur"),
    ("ORDINATEUR", "Ordinateur"), ("ORDI", "Ordinateur"), ("ODINATEUR", "Ordinateur"),
    ("MIRO ORDINATEUR", "Ordinateur"), ("MICRO ODINATEUR", "Ordinateur"),
    ("STATION DE TRAVAIL", "Ordinateur"), ("IMPRIMANTE", "Imprimante"), ("IMPRMANTE", "Imprimante"),
    ("PHOTOCOP", "Photocopieur"), ("PHTOCOP", "Photocopieur"), ("COPIEUR", "Photocopieur"),
    ("CLIMATIS", "Climatiseur"), ("SPLIT", "Climatiseur"), ("VEHICULE", "Véhicule"),
    ("VEHCULE", "Véhicule"), ("VEHUCULE", "Véhicule"), ("VEHICLE", "Véhicule"),
    ("FORD", "Véhicule"), ("TOYOTA", "Véhicule"), ("BUREAU", "Bureau"),
    ("BUREA", "Bureau"), ("FAUTEUIL", "Fauteuil"), ("FAUTEUL", "Fauteuil"),
    ("ARMOIR", "Armoire"), ("TABLE", "Table"), ("CHAISE", "Chaise"),
    ("CANAPE", "Salon / canapé"), ("SALON", "Salon / canapé"), ("TELEVIS", "Téléviseur"),
    ("TELEVISION", "Téléviseur"), ("T V", "Téléviseur"), ("REFRIG", "Réfrigérateur"),
    ("FRIG", "Réfrigérateur"), ("SCANNER", "Scanner"), ("SCANER", "Scanner"),
    ("ONDULEUR", "Onduleur"), ("REGULATEUR", "Onduleur"), ("PARASUR", "Onduleur"),
    ("MOTO", "Motocyclette"), ("GROUPE", "Groupe électrogène"), ("GENERATOR", "Groupe électrogène"),
    ("TELEPHONE", "Téléphone"), ("INTERPHONE", "Téléphone"), ("CLASSEUR", "Classeur"),
    ("MEUBLE", "Meuble"), ("ECRAN", "Écran"), ("UNITE CENTRALE", "Unité centrale"),
    ("RIDEAU", "Rideau / store"), ("STORE", "Rideau / store"), ("PANIER", "Poubelle / panier"),
    ("POUBELLE", "Poubelle / panier"), ("PENDULE", "Horloge"), ("EXTIN", "Extincteur"),
]

MAP_CATEGORIE = {
    "Ordinateur": "Matériel informatique", "Ordinateur portable": "Matériel informatique",
    "Unité centrale": "Matériel informatique", "Écran": "Matériel informatique",
    "Imprimante": "Matériel informatique", "Photocopieur": "Matériel informatique",
    "Scanner": "Matériel informatique", "Onduleur": "Matériel informatique",
    "Serveur": "Matériel informatique", "Équipement réseau": "Matériel informatique",
    "Vidéoprojecteur": "Matériel informatique", "Télécopieur": "Matériel informatique",
    "Souris": "Matériel informatique", "Clavier": "Matériel informatique",
    "Traceur": "Matériel informatique",
    "Équipement de visioconférence": "Matériel informatique",
    "Bureau": "Mobilier de bureau", "Table": "Mobilier de bureau", "Chaise": "Mobilier de bureau",
    "Fauteuil": "Mobilier de bureau", "Salon / canapé": "Mobilier de bureau",
    "Armoire": "Mobilier de bureau", "Étagère": "Mobilier de bureau", "Bibliothèque": "Mobilier de bureau",
    "Coffre-fort": "Mobilier de bureau", "Classeur": "Mobilier de bureau", "Meuble": "Mobilier de bureau",
    "Rideau / store": "Mobilier de bureau", "Lit": "Mobilier de bureau", "Tabouret": "Mobilier de bureau",
    "Climatiseur": "Équipement technique", "Réfrigérateur": "Équipement technique", "Téléviseur": "Équipement technique",
    "Téléphone": "Équipement technique", "Ventilateur": "Équipement technique", "Groupe électrogène": "Équipement technique",
    "Extincteur": "Équipement technique", "Horloge": "Équipement technique", "Cuisinière": "Équipement technique",
    "Filtre à eau": "Équipement technique", "Appareil photo": "Équipement technique", "Matériel topographique": "Équipement technique",
    "Citerne / cuve": "Équipement technique", "Palan": "Équipement technique", "Tronçonneuse": "Équipement technique",
    "Compresseur": "Équipement technique", "Moquette": "Équipement technique", "Motopompe": "Équipement technique",
    "Décodeur": "Équipement technique", "Enceinte / sonorisation": "Équipement technique",
    "Broyeur": "Équipement technique", "Rallonge électrique": "Équipement technique",
    "Véhicule": "Matériel roulant", "Motocyclette": "Matériel roulant", "Camion": "Matériel roulant",
    "Petit matériel de bureau": "Petit matériel de bureau", "Poubelle / panier": "Petit matériel de bureau",
    "Carte murale": "Documentation / signalétique", "Documentation": "Documentation / signalétique",
    "Effigie": "Documentation / signalétique",
}
CATEGORIE_DEFAUT = "Autres biens"

RENOMMAGE_NORM = {
    "IDENTIFIANT UNIQUE": "Identifiant", "IDENTIFGIANT UNIQUE": "Identifiant",
    "MINISTERES": "Ministère", "MINISTERE": "Ministère",
    "STRUCTURES": "Structure", "STRUCTURE": "Structure",
    "LOCALISATION GEPSOFT FICHE DETENTEUR": "Localisation",
    "LOCALISATION GEPSOFT": "Localisation", "FICHE DETENTEUR": "Localisation",
    "TYPE D ADMINISTRATION": "_type_admin_source",
    "DESIGNATION DES MATIERES DENREES ET OBJETS": "_libelle_source",
    "DESCRIPTION": "Description",
    "REGION": "Région", "CHEF LIEU": "Chef-lieu",
    "DEPARTEMENT": "Département", "DEPARTEMENTS": "Département",
    "VILLES": "Ville", "VILLE": "Ville", "MAPS": "_maps_source",
    "DATE": "Date d'acquisition", "MODE D ACQUISITION": "Mode d'acquisition",
    "DATE D AFFECTATION": "Date d'affectation",
    "QTE": "Quantité", "QUANTITE": "Quantité",
    "PRIX UNITAIRE": "Prix unitaire",
    "MONTANT TOTAL": "Valeur d'acquisition",
    "AMORTISSEME NT DEPRECIATION": "Amortissement",
    "AMORTISSEMENT DEPRECIATION": "Amortissement",
    "AMORTISSEMENT": "Amortissement",
    "VALEUR ACTUELLE": "Valeur nette comptable",
    "ETAT DU BIEN": "_etat_source",
    "NATURE": "_nature_source",
    "IMMO STOCK": "_nature_source2",
    "NATURE DU BIEN": "Nature du bien",
    "NATURE DU MATERIEL": "_nature_materiel_source",
    "CODE MINISTERE OU INSTITUTION CONSTITUTIONNEL": "Code ministère",
    "CODE ARRONDISSEMENT": "Code arrondissement",
    "ARRONDISSEMENT": "Arrondissement",
}

FEUILLES_REGIONALES = [
    "CENTRAL", "ADAMAOUA", "EST", "EXTREME_NORD", "EXTREME NORD", "CENTRE",
    "LITTORAL", "NORD", "NORD_OUEST", "NORD OUEST", "OUEST", "SUD", "SUD_OUEST", "SUD OUEST",
]

# =============================================================================
# 2. UTILITAIRES
# =============================================================================


def _normaliser(x) -> str:
    if x is None or pd.isna(x):
        return ""
    s = str(x)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"[\s_/\\\-\.\,\;\:'’]+", " ", s).strip().upper()
    return s


def _norm_col(x) -> str:
    s = unicodedata.normalize("NFKD", str(x)).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"[\s_/\\\-'’\.]+", " ", s).strip().upper()
    return s


def fmt_fcfa(v) -> str:
    if v is None or pd.isna(v):
        return "n.d."
    v = float(v)
    if abs(v) >= 1_000_000_000:
        return f"{v/1_000_000_000:,.2f} Md FCFA".replace(",", " ").replace(".", ",")
    if abs(v) >= 1_000_000:
        return f"{v/1_000_000:,.1f} M FCFA".replace(",", " ").replace(".", ",")
    if abs(v) >= 1_000:
        return f"{v/1_000:,.0f} k FCFA".replace(",", " ")
    return f"{v:,.0f} FCFA".replace(",", " ")


def fmt_fcfa_complet(v) -> str:
    """Montant en toutes unités (« 900 000 FCFA »), pour les infobulles : au
    survol, le lecteur attend le chiffre exact et non son abréviation."""
    if v is None or pd.isna(v):
        return "n.d."
    return f"{float(v):,.0f} FCFA".replace(",", " ")


def fmt_compact(v) -> str:
    """Format court FRANÇAIS pour les étiquettes de barres.
    Remplace le %{text:.2s} de Plotly qui affichait « 1.6G » (notation SI
    anglo-saxonne, incompréhensible dans un rendu FCFA francophone)."""
    if v is None or pd.isna(v):
        return ""
    v = float(v)
    if abs(v) >= 1_000_000_000:
        return f"{v/1_000_000_000:.2f} Md".replace(".", ",")
    if abs(v) >= 1_000_000:
        return f"{v/1_000_000:.1f} M".replace(".", ",")
    if abs(v) >= 1_000:
        return f"{v/1_000:.0f} k"
    return f"{v:.0f}"


def fmt_int(v) -> str:
    if v is None or pd.isna(v):
        return "0"
    return f"{float(v):,.0f}".replace(",", " ")


def safe_numeric(s: pd.Series) -> pd.Series:
    """Conversion numérique robuste.
    Règle décimale corrigée : la virgule n'est décimale QUE s'il y a une seule
    virgule suivie de 1-2 chiffres en fin de chaîne (« 2500,5 ») ; sinon elle
    est un séparateur de milliers (« 1,500,000 ») et est supprimée. Les points
    multiples (« 1.500.000 ») sont traités comme séparateurs de milliers.
    (L'ancienne heuristique transformait « 1,500,000 » en NaN.)"""
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce")
    txt = s.astype(str).str.replace("\u00a0", " ", regex=False).str.strip()
    txt = txt.str.replace(r"[^0-9,\.\-]", "", regex=True)
    multi_pts = txt.str.count(r"\.") >= 2
    txt = txt.mask(multi_pts, txt.str.replace(".", "", regex=False))
    virg_dec = txt.str.match(r"^-?[^,]*,\d{1,2}$", na=False)
    txt = np.where(
        virg_dec,
        txt.str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
        txt.str.replace(",", "", regex=False),
    )
    return pd.to_numeric(pd.Series(txt, index=s.index), errors="coerce")


def convert_dates(s: pd.Series) -> pd.Series:
    """Conversion de dates robuste (texte + numéros de série Excel).
    Fenêtre de plausibilité 1900-2100 : sans elle, une valeur aberrante dans
    la colonne Date (constatée sur les feuilles CENTRAL/régionales) produit
    un débordement OutOfBoundsDatetime qui fait planter le chargement."""
    def _borner(d):
        return d.where((d >= pd.Timestamp("1900-01-01"))
                       & (d <= pd.Timestamp("2100-12-31")))
    date_txt = _borner(pd.to_datetime(s, errors="coerce", dayfirst=True))
    num = pd.to_numeric(s, errors="coerce")
    # Numéros de série Excel plausibles uniquement (1 ≈ 1900, 73 415 ≈ 2100).
    num = num.where((num >= 1) & (num <= 73_415))
    date_num = _borner(pd.to_datetime(num, origin="1899-12-30", unit="D",
                                      errors="coerce"))
    try:
        date_txt = date_txt.astype("datetime64[ns]")
        date_num = date_num.astype("datetime64[ns]")
    except (TypeError, ValueError):
        pass
    return date_txt.fillna(date_num)


def harmoniser_libelle(libelle):
    """Uniformise les libellés d'objets sur toutes les feuilles chargées.

    Aucune ligne n'est supprimée : le libellé brut est conservé dans
    « Libellé source » et les variantes sont regroupées dans « Libellé
    standardisé ». Retourne (libellé standardisé, libellé reconnu ?).
    Ordre des règles : les préfixes les plus spécifiques d'abord (COFFRET
    avant COFFRE, sinon les coffrets électriques deviennent des
    coffres-forts — règle morte constatée dans une version antérieure)."""
    norme = _normaliser(libelle)
    if not norme or norme in {"NAN", "NONE", "NON RENSEIGNE", "NON RENSEIGNEE"}:
        return "Non renseigné", False

    compact = norme.replace(" ", "")

    # Ordinateurs et assimilés.
    if (
        norme.startswith(("ORDINATEUR", "MICRO ORDINATEUR", "MICRO ORDI", "PC COMPLET", "STATION DE TRAVAIL"))
        or compact.startswith(("ORDINATEUR", "ODINATEUR", "ORIDNATEUR", "MORDINATEUR", "ORDINATCN", "MICROODINATEUR", "MIROORDINATEUR", "MICOORDINATEUR"))
        or norme in {"COMPUTER", "UNITE CENTRALE"}
    ):
        return "Ordinateur", True
    if norme.startswith(("ORDINATEUR PORTABLE", "ORDI PORTABLE")) or norme == "LAPTOP":
        return "Ordinateur portable", True

    # Photocopieur / copieur.
    if (
        norme.startswith(("PHOTOCOP", "PHOTOCO", "PHTOCOP", "PHOTOTOCOP", "COPIEUR", "COPIER"))
        or compact.startswith(("PHOTOCOPIEUR", "PHOTOCOPIEMACHINE", "PHOTOCOPIER", "SCANERCOPIEUR", "IMPRIMANTECOPIEUR", "IMPRIMANTECOPIER"))
        or norme in {"PHOTOCOPIER", "PHOTOCOPIE MACHINE", "COPY MACHINE", "SCANER COPIEUR"}
    ):
        return "Photocopieur", True

    # Imprimante pure.
    if norme.startswith(("IMPRIMANTE", "IMPRMANTE", "PRINTER")):
        return "Imprimante", True

    # Matériel roulant.
    if (
        norme.startswith(("VEHICULE", "VEHCULE", "VEHUCULE", "VEHICLE", "PICK UP", "PICK"))
        or norme.startswith(("TOYOTA", "FORD"))
    ):
        return "Véhicule", True
    if norme.startswith(("MOTO POMPE", "MOTOPOMPE")):
        return "Motopompe", True
    if norme.startswith(("MOTO", "MOTOCYCLETTE")):
        return "Motocyclette", True

    # Mobilier.
    if norme.startswith(("BUREAU", "BUREA", "RETOUR")):
        return "Bureau", True
    if norme.startswith(("FAUTEUIL", "FAUTEUL", "FAUTEIL")):
        return "Fauteuil", True
    if norme.startswith(("CHAISE", "CHAISES", "CHAIR")):
        return "Chaise", True
    if norme.startswith(("ARMOIRE", "ARMOIR")):
        return "Armoire", True
    if norme.startswith("CLASSEUR"):
        return "Classeur", True
    if norme.startswith(("SALON", "CANAPE")):
        return "Salon / canapé", True
    if norme.startswith(("TABLE", "GUERIDON", "GUERRIDON", "GERIDON")) or compact.startswith("TABLE"):
        return "Table", True
    if norme.startswith("ETAGERE"):
        return "Étagère", True
    # COFFRET (coffret électrique -> Onduleur) AVANT COFFRE (coffre-fort).
    if norme.startswith("COFFRET"):
        return "Onduleur", True
    if norme.startswith(("COFFRE FORT", "COFFRE")):
        return "Coffre-fort", True
    if norme.startswith("MEUBLE"):
        return "Meuble", True
    if norme.startswith("LIT"):
        return "Lit", True
    if norme.startswith("TABOURET"):
        return "Tabouret", True

    # Équipements techniques et autres familles récurrentes.
    if norme.startswith(("CLIMATISEUR", "CLIMATISATEUR", "CLIMATISSEUR", "SPLIT")):
        return "Climatiseur", True
    if norme.startswith(("TELEVISEUR", "TELEVISION", "POSTE TELEVISEUR", "TV", "T V")):
        return "Téléviseur", True
    if (norme.startswith(("REFRIGERATEUR", "REFRIGIRATEUR", "REFRIGER", "REFREG", "FRIGO", "FRIGIDAIRE", "FRIDGE", "MINI FRIG", "PORTE REFRIG"))
            or compact.startswith(("MINIFRIG", "PORTEREFRIG"))):
        return "Réfrigérateur", True
    if norme.startswith(("VENTILATEUR", "VENTILATOR", "VENTILO")):
        return "Ventilateur", True
    if norme.startswith(("GROUPE ELECTROGENE", "GROUPE", "GENERATOR")):
        return "Groupe électrogène", True
    if norme.startswith(("SERVEUR", "CERVEUR")):
        return "Serveur", True
    if norme.startswith(("SWITCH", "ROUTEUR", "MODEM")):
        return "Équipement réseau", True
    if norme.startswith(("ONDULEUR", "REGULATEUR", "PARASUR", "PARA SUR")):
        return "Onduleur", True
    if norme.startswith(("ECRAN", "MONITEUR")):
        return "Écran", True
    if norme.startswith("CLAVIER"):
        return "Clavier", True
    if norme.startswith(("MOUSE", "SOURIS")):
        return "Souris", True
    if norme.startswith("TRACEUR"):
        return "Traceur", True
    if norme.startswith(("SCANNER", "SCANER")):
        return "Scanner", True
    if norme.startswith(("FAX", "FAXEUR")):
        return "Télécopieur", True
    if norme.startswith("VIDEOPROJECT") or compact.startswith("VIDEOPROJECT"):
        return "Vidéoprojecteur", True
    if norme.startswith(("EQUIPEMENT VIDEO", "VISIOCONFERENCE", "VIDEO CONFERENCE")):
        return "Équipement de visioconférence", True
    if norme.startswith(("GPS", "TEODOL", "THEODOL", "DECAMETRE")):
        return "Matériel topographique", True
    if norme.startswith(("CITERNE", "CUVE")):
        return "Citerne / cuve", True
    if norme.startswith(("PALLAN", "PALAN")):
        return "Palan", True
    if norme.startswith("TRONCONNEUSE"):
        return "Tronçonneuse", True
    if norme.startswith(("COMPRESSOR", "COMPRESSEUR")):
        return "Compresseur", True
    if norme.startswith("BROYEUR"):
        return "Broyeur", True
    if norme.startswith("EXTINCT") or norme.startswith("EXTINTEUR"):
        return "Extincteur", True
    if norme.startswith("APPAREIL PHOTO"):
        return "Appareil photo", True
    if norme.startswith(("RIDEAU", "STORE", "PORTE RIDEAU")):
        return "Rideau / store", True
    if norme.startswith("MOQUETTE"):
        return "Moquette", True
    if norme.startswith("DECODEUR"):
        return "Décodeur", True
    if norme.startswith("BAFFLE"):
        return "Enceinte / sonorisation", True
    if norme.startswith(("PANIER", "POUBELLE", "CENDRIER", "POT ")) or norme == "POT A FLEURS":
        return "Poubelle / panier", True
    if (norme.startswith(("AGRAFEUSE", "PERFORATEUR", "RELIEUSE", "CALCULATRICE",
                          "ADDITIONNEUSE", "SOUS MAIN", "PORTE CACHET", "POETE CACHET",
                          "CACHET", "MACHINE A PERFORER", "MACHINE SPIRAL",
                          "PORTE UNITE"))
            or compact.startswith(("SOUSMAIN", "PORTECACHET"))):
        return "Petit matériel de bureau", True
    if norme.startswith(("PENDULE", "CLOCK", "HORLOGE")):
        return "Horloge", True
    if norme.startswith("DICTIONNAIRE"):
        return "Documentation", True
    if norme.startswith("RALLONGE"):
        return "Rallonge électrique", True

    # Dictionnaire exact + règles par préfixe : conserve les entrées déjà codées.
    if norme in MAP_LIBELLE:
        return MAP_LIBELLE[norme], True
    for prefixe, standard in REGLES_PREFIXE:
        if norme.startswith(prefixe):
            return standard, True
    for cle, standard in MAP_LIBELLE.items():
        if compact == cle.replace(" ", ""):
            return standard, True
    for prefixe, standard in REGLES_PREFIXE:
        if compact.startswith(prefixe.replace(" ", "")):
            return standard, True

    return norme.title(), False


def harmoniser_mode_acquisition(mode):
    """Uniformise le mode d'acquisition. La valeur brute est conservée dans
    « Mode d'acquisition source » pour la traçabilité."""
    norme = _normaliser(mode)
    if not norme or norme in {"NAN", "NONE", "NON RENSEIGNE", "NON RENSEIGNEE"}:
        return "Non renseigné", False
    if norme in MAP_MODE_ACQUISITION:
        return MAP_MODE_ACQUISITION[norme], True
    if "BILAN" in norme and "OUVERTURE" in norme:
        return "Bilan d'ouverture", True
    if "DON" in norme or "LEG" in norme:
        return "Don ou legs", True
    if "ACQUIS" in norme or "ACHAT" in norme:
        return "Acquisition", True
    return norme.title(), False


def harmoniser_nature_materiel(nature):
    """Uniformise la « nature du matériel » (nomenclature comptable de la
    fiche de recensement). C'est l'axe d'analyse réclamé pour répondre à des
    questions du type « combien de matériel roulant détient telle
    administration ? » sans passer par le libellé de chaque objet."""
    norme = _normaliser(nature)
    if not norme or norme in {"NAN", "NONE", "NON RENSEIGNE", "NON RENSEIGNEE"}:
        return "Non renseignée"
    norme = re.sub(r"\s+", " ", norme)
    if norme in MAP_NATURE_MATERIEL:
        return MAP_NATURE_MATERIEL[norme]
    for cle, standard in MAP_NATURE_MATERIEL.items():
        if norme.startswith(cle):
            return standard
    return norme.capitalize()


def normaliser_ville(v) -> str:
    s = _normaliser(v)
    if not s or s in {"NAN", "NONE"}:
        return ""
    s = re.sub(r"\s+(1ER|1ERE|IER|2E|2EME|3E|3EME|4E|4EME|5E|5EME|I|II|III|IV|V|VI|VII)$", "", s).strip()
    return s


def coords_ville(v):
    s = normaliser_ville(v)
    if not s:
        return None
    if s in COORDS_VILLES:
        return COORDS_VILLES[s]
    for nom, coord in COORDS_VILLES.items():
        if s.startswith(nom) or nom.startswith(s):
            return coord
    return None


_RE_COORD = re.compile(r"(-?\d{1,2}\.\d+)\s*[,;/\s]\s*(-?\d{1,3}\.\d+)")


def parser_maps(val):
    if val is None or pd.isna(val):
        return None
    m = _RE_COORD.search(str(val))
    if not m:
        return None
    a, b = float(m.group(1)), float(m.group(2))
    # Emprise approximative du Cameroun.
    if 1.0 <= a <= 13.6 and 8.0 <= b <= 16.5:
        return a, b
    if 1.0 <= b <= 13.6 and 8.0 <= a <= 16.5:
        return b, a
    return None


def indice_sante(d: pd.DataFrame):
    m = d["Score santé"].notna()
    poids = d.loc[m, "Quantité"].fillna(1)
    if poids.sum() <= 0:
        return None
    return float((d.loc[m, "Score santé"] * poids).sum() / poids.sum())


def indice_sante_par(d: pd.DataFrame, dim: str) -> pd.Series:
    t = d[d["Score santé"].notna()].copy()
    if t.empty or dim not in t.columns:
        return pd.Series(dtype=float)
    t["_num"] = t["Score santé"] * t["Quantité"].fillna(1)
    g = t.groupby(dim, dropna=False)
    return (g["_num"].sum() / g["Quantité"].sum()).rename("Indice de santé")


def label_sante(v):
    if v is None or pd.isna(v):
        return "Non calculé"
    if v >= 80:
        return "Patrimoine sain"
    if v >= 60:
        return "Vigilance"
    if v >= 40:
        return "Critique"
    return "Intervention urgente"


def couleur_sante(v):
    if v is None or pd.isna(v):
        return THEME["texte2"]
    if v >= 80:
        return "#159947"
    if v >= 60:
        return "#B8860B"
    if v >= 40:
        return "#E8710A"
    return "#B42318"


def texte_sante(v):
    """Formatage SÛR de l'indice (l'ancienne version plantait avec
    TypeError quand l'indice était incalculable, p. ex. filtre
    « État = Non renseigné »)."""
    if v is None or pd.isna(v):
        return "non calculable (aucun état renseigné dans le périmètre filtré)"
    return f"{v:.1f}/100 ({label_sante(v)})".replace(".", ",")


def to_excel_bytes(frames: dict) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for sheet, data in frames.items():
            data.to_excel(writer, sheet_name=str(sheet)[:31], index=False)
    return out.getvalue()

# =============================================================================
# 3. CHARGEMENT ET NETTOYAGE
# =============================================================================


def open_excel(source):
    if isinstance(source, (bytes, bytearray)):
        return pd.ExcelFile(io.BytesIO(source))
    return pd.ExcelFile(source)


def detecter_strategies(sheet_names):
    """Stratégies de chargement, la plus couvrante en premier.

    La feuille CAMEROUN ne contient qu'un seul ministère ; l'ensemble
    CENTRAL + régions couvre toutes les administrations recensées. Comme la v6
    est construite autour du filtre ministériel, c'est cette dernière qui est
    proposée par défaut — sinon la liste des ministères n'aurait qu'une entrée.
    """
    normes = {_norm_col(s): s for s in sheet_names}
    strategies = []
    reg = []
    for r in FEUILLES_REGIONALES:
        nr = _norm_col(r)
        if nr in normes:
            reg.append(normes[nr])
    if reg:
        strategies.append(("CENTRAL + régions — tous ministères", list(dict.fromkeys(reg))))
    if "CAMEROUN" in normes:
        strategies.append(("CAMEROUN — base nationale consolidée", [normes["CAMEROUN"]]))
    immo_stock = []
    for n in ["IMMO", "IMMO MINTP", "STOCK MINTP TROUVE", "STOCK MINTP"]:
        if n in normes:
            immo_stock.append(normes[n])
    if immo_stock:
        strategies.append(("IMMO + STOCK — feuilles sources", list(dict.fromkeys(immo_stock))))
    if not strategies:
        strategies.append((f"Première feuille — {sheet_names[0]}", [sheet_names[0]]))
    return strategies


def charger_feuille(xls, nom):
    d = pd.read_excel(xls, sheet_name=nom)
    n_vides = int(d.isna().all(axis=1).sum())
    d = d.dropna(how="all").copy()
    d.attrs["n_lignes_vides_ecartees"] = n_vides
    d.columns = [str(c) for c in d.columns]
    ren = {}
    vus = set()
    for c in d.columns:
        cible = RENOMMAGE_NORM.get(_norm_col(c))
        if cible and cible not in vus:
            ren[c] = cible
            vus.add(cible)
    d = d.rename(columns=ren)
    d["Feuille source"] = nom
    return d


def charger_et_nettoyer(source, feuilles: tuple):
    xls = open_excel(source)
    frames = [charger_feuille(xls, f) for f in feuilles if f in xls.sheet_names]
    if not frames:
        frames = [charger_feuille(xls, xls.sheet_names[0])]
    df = pd.concat(frames, ignore_index=True)

    rapport = {"feuilles_chargees": list(feuilles), "n_lignes_brutes": len(df)}

    rapport["n_lignes_vides_ecartees"] = sum(f.attrs.get("n_lignes_vides_ecartees", 0) for f in frames)

    avant = len(df)
    df = df.drop_duplicates().reset_index(drop=True)
    rapport["n_doublons_stricts_supprimes"] = avant - len(df)

    attendues = [
        "Identifiant", "Ministère", "Structure", "Localisation", "_type_admin_source",
        "_libelle_source", "Description", "Région", "Chef-lieu", "Département", "Ville",
        "_maps_source", "Date d'acquisition", "Mode d'acquisition", "Date d'affectation",
        "Quantité", "Prix unitaire", "Valeur d'acquisition", "Amortissement",
        "Valeur nette comptable", "_etat_source", "_nature_source", "_nature_source2",
        "_nature_materiel_source", "Nature du bien", "Code ministère", "Arrondissement",
    ]
    rapport["colonnes_manquantes"] = [c for c in attendues if c not in df.columns]
    for c in attendues:
        if c not in df.columns:
            df[c] = pd.NA

    # Nettoyage numérique.
    for c in ["Quantité", "Prix unitaire", "Valeur d'acquisition", "Amortissement", "Valeur nette comptable"]:
        df[c] = safe_numeric(df[c])

    df["_qte_imputee"] = df["Quantité"].isna() | (df["Quantité"] <= 0)
    rapport["n_qte_manquante"] = int(df["_qte_imputee"].sum())
    rapport["n_pu_manquant"] = int((df["Prix unitaire"].isna() | (df["Prix unitaire"] <= 0)).sum())
    df["Quantité"] = df["Quantité"].fillna(1).clip(lower=1)

    # Si la VNC manque mais VA et amortissement existent, on la reconstitue,
    # et on COMPTE ces imputations (sinon l'indicateur de qualité serait
    # circulaire : une VNC reconstruite est cohérente par construction).
    vnc_manquante = df["Valeur nette comptable"].isna() & df["Valeur d'acquisition"].notna() & df["Amortissement"].notna()
    df["VNC imputée"] = vnc_manquante
    rapport["n_vnc_imputees"] = int(vnc_manquante.sum())
    vnc_reconst = df["Valeur d'acquisition"] - df["Amortissement"]
    df["Valeur nette comptable"] = df["Valeur nette comptable"].fillna(vnc_reconst)

    ecart = (df["Valeur d'acquisition"] - df["Amortissement"].fillna(0) - df["Valeur nette comptable"]).abs()
    df["_incoherence_vnc"] = ecart > 1
    rapport["n_incoherences_comptables"] = int(df["_incoherence_vnc"].sum())

    # Libellé.
    res = [harmoniser_libelle(x) for x in df["_libelle_source"]]
    df["Libellé standardisé"] = [r[0] for r in res]
    df["_libelle_resolu"] = [r[1] for r in res]
    rapport["taux_harmonisation"] = float(df["_libelle_resolu"].mean()) if len(df) else 0.0
    rapport["libelles_non_resolus"] = (
        df.loc[~df["_libelle_resolu"], "_libelle_source"]
        .dropna().astype(str).str.strip().value_counts()
    )
    df["Libellé source"] = df["_libelle_source"].astype("string").str.strip()
    df["Catégorie"] = df["Libellé standardisé"].map(MAP_CATEGORIE).fillna(CATEGORIE_DEFAUT)

    # Mode d'acquisition : valeur brute conservée, valeur standardisée utilisée.
    df["Mode d'acquisition source"] = df["Mode d'acquisition"].astype("string").str.strip()
    mode_res = [harmoniser_mode_acquisition(x) for x in df["Mode d'acquisition source"]]
    df["Mode d'acquisition"] = [r[0] for r in mode_res]
    df["_mode_acquisition_resolu"] = [r[1] for r in mode_res]
    rapport["taux_harmonisation_modes_acquisition"] = float(df["_mode_acquisition_resolu"].mean()) if len(df) else 0.0
    rapport["modes_acquisition_non_resolus"] = (
        df.loc[~df["_mode_acquisition_resolu"], "Mode d'acquisition source"]
        .dropna().astype(str).str.strip().value_counts()
    )

    # Administration.
    df["Administration"] = df["_type_admin_source"].map(_normaliser).map(MAP_ADMINISTRATION).fillna("Non renseigné")
    rapport["n_admin_non_renseignee"] = int((df["Administration"] == "Non renseigné").sum())

    # État du bien.
    # Règle v6 : l'état manquant est imputé à « Bon état » — c'est la
    # convention retenue par le métier pour que 100 % du parc soit exploitable
    # dans les indicateurs. L'imputation est conservée dans « État imputé » :
    # sans cette trace, l'indice de santé deviendrait un artefact de la règle
    # et la qualité du recensement ne serait plus mesurable.
    etat_brut = df["_etat_source"].map(_normaliser).map(MAP_ETAT)
    df["État imputé"] = etat_brut.isna()
    df["État du bien"] = etat_brut.fillna(ETAT_PAR_DEFAUT)
    rapport["n_etat_non_renseigne"] = int(df["État imputé"].sum())
    rapport["etat_par_defaut"] = ETAT_PAR_DEFAUT
    df["Score santé"] = df["État du bien"].map(SCORE_SANTE)

    # Nature du matériel (nomenclature de la fiche de recensement).
    source_nature_mat = df["_nature_materiel_source"]
    if "Nature du bien" in df.columns:
        source_nature_mat = source_nature_mat.fillna(df["Nature du bien"])
    df["Nature du matériel"] = [harmoniser_nature_materiel(x) for x in source_nature_mat]
    rapport["n_nature_materiel_non_renseignee"] = int((df["Nature du matériel"] == "Non renseignée").sum())

    # Textes utiles.
    df["Région"] = df["Région"].map(_normaliser).map(MAP_REGION).fillna("Non renseigné")
    for c in ["Département", "Ville", "Structure", "Localisation", "Ministère",
              "Description", "Identifiant", "Arrondissement"]:
        df[c] = df[c].astype("string").str.strip()
        df.loc[df[c].isin(["", "nan", "NaN", "None", "<NA>"]), c] = pd.NA

    # Ministère : sigle normalisé + libellé complet issu du référentiel.
    # Les lignes sans ministère saisi ne sont pas écartées : elles sont
    # regroupées sous « Non renseigné » et restent filtrables, sans quoi leur
    # valeur disparaîtrait silencieusement du total consolidé.
    df["Ministère"] = df["Ministère"].str.upper().str.replace(r"\s+", " ", regex=True).str.strip()
    df["Ministère"] = df["Ministère"].fillna("NON RENSEIGNÉ")
    df["Ministère (libellé)"] = df["Ministère"].map(
        lambda s: REFERENTIEL_MINISTERES.get(s, "Ministère ou institution non référencé"
                                             if s != "NON RENSEIGNÉ" else "Ministère non renseigné")
    )
    rapport["ministeres_presents"] = sorted(df["Ministère"].dropna().unique().tolist())
    rapport["ministeres_hors_referentiel"] = sorted(
        m for m in df["Ministère"].dropna().unique()
        if m not in REFERENTIEL_MINISTERES and m != "NON RENSEIGNÉ"
    )

    df["Structure"] = df["Structure"].fillna(df["Localisation"])
    df["Structure agrégée"] = (
        df["Structure"].fillna("Non renseignée").astype(str)
        .str.replace(r"\s*FD\s*N.*$", "", regex=True)
        .str.strip()
    )

    # Dates et ancienneté.
    df["Date d'acquisition"] = convert_dates(df["Date d'acquisition"])
    df["Date d'affectation"] = convert_dates(df["Date d'affectation"])
    ref = df["Date d'acquisition"].fillna(df["Date d'affectation"])
    df["Ancienneté (années)"] = ((pd.Timestamp.today() - ref).dt.days / 365.25).clip(lower=0)

    # Classification comptable.
    df["Classification"] = "Indéterminée"
    connu = df["Prix unitaire"].notna() & (df["Prix unitaire"] > 0)
    df.loc[connu & (df["Prix unitaire"] >= SEUIL_IMMOBILISATION_FCFA), "Classification"] = "Immobilisation"
    df.loc[connu & (df["Prix unitaire"] < SEUIL_IMMOBILISATION_FCFA), "Classification"] = "Stock"

    map_nature = {
        "IMMO": "Immobilisation", "IMMOBILISATION": "Immobilisation", "IMMOBILISATIONS": "Immobilisation",
        "STOCK": "Stock", "STOCKS": "Stock",
    }
    n1 = df["_nature_source"].map(_normaliser).map(map_nature)
    n2 = df["_nature_source2"].map(_normaliser).map(map_nature)
    df["Nature saisie"] = n1.fillna(n2).fillna("Non renseignée")
    df["Contrôle classification"] = "Indéterminé"
    comp = df["Nature saisie"].isin(["Immobilisation", "Stock"]) & (df["Classification"] != "Indéterminée")
    df.loc[comp & (df["Nature saisie"] == df["Classification"]), "Contrôle classification"] = "Conforme"
    df.loc[comp & (df["Nature saisie"] != df["Classification"]), "Contrôle classification"] = "À vérifier"
    rapport["n_divergences_classification"] = int((df["Contrôle classification"] == "À vérifier").sum())

    # Amortissement et risque (score transparent, PAS un modèle prédictif :
    # base transversale, aucune validation temporelle possible).
    va = df["Valeur d'acquisition"].where(df["Valeur d'acquisition"] > 0)
    df["Taux d'amortissement"] = (df["Amortissement"].fillna(0) / va).clip(0, 1)
    pts_etat = df["État du bien"].map({"Neuf": 0.0, "Bon état": 0.35, "À réformer": 0.9, "Réformé": 1.0})
    med_amort = df["Taux d'amortissement"].median() if df["Taux d'amortissement"].notna().any() else 0.5
    df["Score risque (0-100)"] = (100 * (0.6 * df["Taux d'amortissement"].fillna(med_amort) + 0.4 * pts_etat.fillna(0.5))).round(1)
    df["Niveau de risque"] = pd.cut(
        df["Score risque (0-100)"], bins=[-0.1, 40, 70, 100.1], labels=["Faible", "Moyen", "Élevé"]
    ).astype(str).replace("nan", "Non renseigné")

    # Géolocalisation en cascade : MAPS -> ville -> chef-lieu -> non géolocalisé.
    maps_xy = df["_maps_source"].map(parser_maps)
    ville_xy = df["Ville"].map(coords_ville)
    reg_xy = df["Région"].map(lambda r: COORD_REGIONS.get(r))
    lat, lon, src = [], [], []
    for m, v, r in zip(maps_xy, ville_xy, reg_xy):
        if m:
            lat.append(m[0]); lon.append(m[1]); src.append("Coordonnées MAPS")
        elif v:
            lat.append(v[0]); lon.append(v[1]); src.append("Ville")
        elif r:
            lat.append(r[0]); lon.append(r[1]); src.append("Chef-lieu de région")
        else:
            lat.append(np.nan); lon.append(np.nan); src.append("Non géolocalisé")
    df["lat"], df["lon"], df["Source géolocalisation"] = lat, lon, src
    rapport["n_non_geolocalises"] = int((df["Source géolocalisation"] == "Non géolocalisé").sum())
    rapport["villes_non_geolocalisees"] = sorted(
        df.loc[df["Ville"].notna() & df["Ville"].map(coords_ville).isna(), "Ville"].dropna().unique().tolist()
    )

    # Zone territoriale : la région, sauf pour l'administration centrale qui
    # est isolée. Sans cette distinction, les services centraux (tous situés à
    # Yaoundé) gonflent mécaniquement la région du Centre et rendent toute
    # comparaison interrégionale trompeuse. Les deux lectures restent
    # disponibles : « Région » (brute) et « Zone territoriale » (corrigée).
    df["Zone territoriale"] = np.where(
        df["Administration"].eq("Centrale"),
        "Administration centrale",
        df["Région"].fillna("Non renseigné"),
    )

    rapport["n_lignes_finales"] = len(df)
    return df, rapport


def _cle_entite(nom) -> str:
    """Clé de rapprochement entre un libellé de la base et un nom du fond.

    On retire accents, ponctuation et mots de liaison (« et », « de »…) :
    « NYONG ET SO'O » et « Nyong-et-So » doivent converger."""
    s = unicodedata.normalize("NFKD", str(nom)).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"\b(ET|AND|DE|DU|DES|LA|LE|LES)\b", " ", s.upper())
    return re.sub(r"[^A-Z0-9]+", "", s)


def apparier_au_fond(valeurs, noms_fond, corrections=None):
    """Rapproche les libellés de la base des entités du fond administratif.

    Trois passes, de la plus sûre à la plus permissive : clé normalisée, table
    de corrections explicite, puis rapprochement approximatif (seuil élevé)
    pour les fautes de frappe résiduelles. Retourne les correspondances et la
    liste des libellés non appariés — cette liste est montrée à l'utilisateur
    plutôt que masquée, car une entité non appariée est absente de la carte.
    """
    index_fond = {_cle_entite(n): n for n in noms_fond}
    corrections = {_cle_entite(k): v for k, v in (corrections or {}).items()}
    noms_valides = set(noms_fond)
    correspondances, absents = {}, []
    for valeur in valeurs:
        if valeur is None or (isinstance(valeur, float) and np.isnan(valeur)):
            continue
        cle = _cle_entite(valeur)
        if cle in index_fond:
            correspondances[valeur] = index_fond[cle]
        elif cle in corrections and corrections[cle] in noms_valides:
            correspondances[valeur] = corrections[cle]
        else:
            proches = difflib.get_close_matches(cle, list(index_fond), n=1, cutoff=0.88)
            if proches:
                correspondances[valeur] = index_fond[proches[0]]
            else:
                absents.append(str(valeur))
    return correspondances, sorted(set(absents))


def _bornes_entite(feature):
    """Boîte englobante d'une entité, pour recentrer la carte dessus."""
    lons, lats = [], []

    def parcourir(coords):
        if isinstance(coords[0], (int, float)):
            lons.append(coords[0])
            lats.append(coords[1])
        else:
            for sous in coords:
                parcourir(sous)

    parcourir(feature["geometry"]["coordinates"])
    return min(lons), min(lats), max(lons), max(lats)


def _hsl_vers_hex(h, s, l):
    s, l = s / 100.0, l / 100.0
    c = (1 - abs(2 * l - 1)) * s
    x = c * (1 - abs((h / 60.0) % 2 - 1))
    m = l - c / 2
    r, v, b = [(c, x, 0), (x, c, 0), (0, c, x),
               (0, x, c), (x, 0, c), (c, 0, x)][int(h // 60) % 6]
    return "#%02X%02X%02X" % (round((r + m) * 255), round((v + m) * 255),
                              round((b + m) * 255))


def _point_dans_anneau(lon, lat, anneau):
    """Test point-dans-polygone par lancer de rayon (algorithme pair-impair)."""
    dedans = False
    n = len(anneau)
    for i in range(n):
        x1, y1 = anneau[i][0], anneau[i][1]
        x2, y2 = anneau[(i + 1) % n][0], anneau[(i + 1) % n][1]
        if (y1 > lat) != (y2 > lat):
            # Abscisse de l'intersection du bord avec l'horizontale du point.
            x = x1 + (lat - y1) * (x2 - x1) / (y2 - y1)
            if lon < x:
                dedans = not dedans
    return dedans


def _point_dans_feature(lon, lat, feature):
    geom = feature["geometry"]
    polygones = ([geom["coordinates"]] if geom["type"] == "Polygon"
                 else geom["coordinates"])
    for polygone in polygones:
        if not polygone:
            continue
        if _point_dans_anneau(lon, lat, polygone[0]):
            # Les anneaux suivants sont des trous.
            if not any(_point_dans_anneau(lon, lat, trou) for trou in polygone[1:]):
                return True
    return False


def entites_contenues(centres_enfants, geojson_parent, nom_parent, avec_repli=True):
    """Noms des subdivisions dont le centre tombe dans l'entité parente.

    Sert à choisir les subdivisions à NOMMER dans la vue courante, pas à
    filtrer les polygones tracés : les fonds ADM2 et ADM3 proviennent de
    sources différentes et ne s'emboîtent pas exactement (le fond départemental
    place Yaoundé I, II, V et VII hors du Mfoundi). S'en servir pour découper
    l'affichage amputerait donc la carte ; s'en servir pour choisir des
    étiquettes ne fait, au pire, qu'en afficher une de trop.

    Le test porte sur le centre de chaque subdivision, avec repli sur la boîte
    englobante du parent, ce qui évite la dépendance à une bibliothèque
    géospatiale.
    """
    morceaux = [f for f in geojson_parent.get("features", []) if f["id"] == nom_parent]
    if not morceaux:
        return set()
    # Boîte englobante du parent, pour le repli.
    bornes_parent = None
    for f in morceaux:
        b = _bornes_entite(f)
        bornes_parent = b if bornes_parent is None else (
            min(bornes_parent[0], b[0]), min(bornes_parent[1], b[1]),
            max(bornes_parent[2], b[2]), max(bornes_parent[3], b[3]))

    retenues = set()
    for nom, centre in (centres_enfants or {}).items():
        if any(_point_dans_feature(centre["lon"], centre["lat"], f) for f in morceaux):
            retenues.add(nom)
            continue
        # Repli : le centre d'une entité très découpée ou concave peut tomber
        # hors de son propre parent (les polygones sont simplifiés). On retient
        # alors les entités dont le centre est dans la boîte du parent : mieux
        # vaut afficher une limite voisine de trop qu'amputer Yaoundé de quatre
        # de ses arrondissements.
        if avec_repli and bornes_parent:
            lon_min, lat_min, lon_max, lat_max = bornes_parent
            if (lon_min <= centre["lon"] <= lon_max
                    and lat_min <= centre["lat"] <= lat_max):
                retenues.add(nom)
    return retenues


def eclaircir(couleur_hex, part_blanc=0.62):
    """Version pastel d'une couleur : mélange avec du blanc.

    Sur la carte de points, le fond doit distinguer les territoires sans
    concurrencer les marqueurs. Les teintes pleines de la choroplèthe
    écraseraient les points (jaune sur jaune, vert sur vert), mais un
    éclaircissement trop fort les ramène toutes au même beige : à 80 % de
    blanc, l'Adamaoua (#D2E2DD) et l'Est (#D5E2E9) sont indiscernables. La
    valeur retenue, 62 %, garde des teintes franchement séparées tout en
    laissant les points au premier plan.
    """
    couleur_hex = couleur_hex.lstrip("#")
    r, v, b = (int(couleur_hex[i:i + 2], 16) for i in (0, 2, 4))
    melange = lambda c: round(c + (255 - c) * part_blanc)
    return "#%02X%02X%02X" % (melange(r), melange(v), melange(b))


def colorscale_discrete(couleurs):
    """Échelle en paliers : une couleur par entité, sur une seule trace.

    Tracer une entité par trace coûterait 360 traces pour les arrondissements.
    On code plutôt l'entité par un entier et on découpe l'échelle en autant de
    paliers nets, ce qui donne le même rendu en une seule couche.
    """
    n = max(len(couleurs), 1)
    echelle = []
    for i, couleur in enumerate(couleurs):
        echelle.append([i / n, couleur])
        echelle.append([(i + 1) / n, couleur])
    return echelle or [[0, "#F2EFE6"], [1, "#F2EFE6"]]


def palette_entites(noms):
    """Teintes distinctes pour un niveau administratif quelconque.

    Les dix régions gardent leurs couleurs historiques — ce sont des repères
    déjà acquis. Au-delà, les teintes sont réparties sur le cercle chromatique
    par l'angle d'or, ce qui maximise l'écart entre voisines même avec soixante
    entités ; deux niveaux de luminosité alternent pour séparer les teintes
    proches. Saturation et luminosité restent dans la plage qui garantit un
    contraste suffisant avec le texte de la légende.
    """
    couleurs = {}
    index = 0
    for nom in noms:
        if nom in COULEURS_REGIONS:
            couleurs[nom] = COULEURS_REGIONS[nom]
            continue
        teinte = (index * 137.508) % 360
        couleurs[nom] = _hsl_vers_hex(teinte, 46, 34 if index % 2 else 44)
        index += 1
    return couleurs


def charger_geojson_admin(chemin: str, corrections_noms=None):
    """Charge un fond administratif (régions, départements ou arrondissements).

    Retourne (geojson, noms triés, centres) où `centres` donne pour chaque
    entité son point de recentrage et le zoom adapté à son étendue.
    """
    try:
        with open(chemin, "r", encoding="utf-8") as fh:
            brut = json.load(fh)
    except (OSError, ValueError):
        return None, [], {}
    features, bornes = [], {}
    for feat in brut.get("features", []):
        nom = (feat.get("properties") or {}).get("shapeName")
        if not nom:
            continue
        nom = (corrections_noms or {}).get(nom, nom)
        feat = dict(feat)
        feat["properties"] = {"entite": nom}
        feat["id"] = nom
        lon_min, lat_min, lon_max, lat_max = _bornes_entite(feat)
        # Une même entité peut arriver en plusieurs morceaux (Ndian et
        # Sanaga-Maritime comptent chacun deux polygones dans le fond ADM2) :
        # les bornes sont fusionnées pour que le recadrage englobe le tout, et
        # le nom n'est compté qu'une fois — un doublon dans l'index rendrait
        # toute lecture de valeur ambiguë.
        if nom in bornes:
            a, b, c, d = bornes[nom]
            bornes[nom] = (min(a, lon_min), min(b, lat_min),
                           max(c, lon_max), max(d, lat_max))
        else:
            bornes[nom] = (lon_min, lat_min, lon_max, lat_max)
        features.append(feat)
    if not features:
        return None, [], {}

    centres = {}
    for nom, (lon_min, lat_min, lon_max, lat_max) in bornes.items():
        etendue = max(lon_max - lon_min, lat_max - lat_min, 0.05)
        centres[nom] = {
            "lat": (lat_min + lat_max) / 2,
            "lon": (lon_min + lon_max) / 2,
            "bornes": (lon_min, lat_min, lon_max, lat_max),
            # À un zoom z, la carte montre environ 360/2^z degrés de large.
            # Le zoom qui fait tenir l'entité est donc log2(360/étendue), moins
            # une marge pour ne pas coller aux bords.
            "zoom": float(np.clip(np.log2(360.0 / etendue) - 0.8, 4.4, 10.0)),
        }
    return {"type": "FeatureCollection", "features": features}, sorted(bornes), centres


def charger_geojson_regions(chemin: str):
    """Charge le fond administratif des 10 régions et renomme les entités en
    français. Retourne (geojson, [noms de régions disponibles])."""
    try:
        with open(chemin, "r", encoding="utf-8") as fh:
            brut = json.load(fh)
    except (OSError, ValueError):
        return None, []
    features, noms = [], []
    for feat in brut.get("features", []):
        nom_en = (feat.get("properties") or {}).get("shapeName")
        nom_fr = NOMS_REGIONS_GEOJSON.get(nom_en)
        if not nom_fr:
            continue
        feat = dict(feat)
        feat["properties"] = {"region": nom_fr}
        feat["id"] = nom_fr
        features.append(feat)
        noms.append(nom_fr)
    if not features:
        return None, []
    return {"type": "FeatureCollection", "features": features}, sorted(noms)

# =============================================================================
# 4. AUDIT AUTOMATIQUE
# =============================================================================

REGLES_ANOMALIES = [
    ("A01", "Bien à réformer avec VNC positive", "Moyenne", "Préparer la réforme et la sortie comptable", lambda d: (d["État du bien"] == "À réformer") & (d["Valeur nette comptable"] > 0)),
    ("A02", "Bien réformé avec VNC positive", "Haute", "Vérifier amortissement et régulariser la VNC", lambda d: (d["État du bien"] == "Réformé") & (d["Valeur nette comptable"] > 0)),
    ("A03", "VNC supérieure à la valeur d'acquisition", "Haute", "Vérifier les montants VA / amortissement / VNC", lambda d: d["Valeur nette comptable"] > d["Valeur d'acquisition"]),
    ("A04", "Amortissement supérieur à la valeur d'acquisition", "Haute", "Régulariser l'amortissement", lambda d: d["Amortissement"] > d["Valeur d'acquisition"]),
    ("A05", "Prix unitaire nul ou manquant", "Moyenne", "Compléter le prix unitaire", lambda d: d["Prix unitaire"].isna() | (d["Prix unitaire"] <= 0)),
    ("A06", "Quantité nulle ou manquante", "Moyenne", "Compléter la quantité réelle", lambda d: d["_qte_imputee"]),
    # L'état manquant étant désormais imputé à « Bon état », le contrôle porte
    # sur l'imputation elle-même : la valeur affichée est une convention, pas
    # un constat de terrain, et doit être confirmée par le service détenteur.
    ("A07", "État du bien imputé (non constaté au recensement)", "Moyenne", "Faire confirmer l'état physique par le service détenteur", lambda d: d["État imputé"]),
    ("A08", "Administration non renseignée", "Faible", "Compléter le type d'administration", lambda d: d["Administration"] == "Non renseigné"),
    ("A09", "Localisation administrative incomplète", "Faible", "Compléter région / département / ville", lambda d: (d["Région"] == "Non renseigné") | d["Département"].isna() | d["Ville"].isna()),
    ("A10", "Bien non géolocalisable", "Faible", "Compléter la ville ou les coordonnées MAPS", lambda d: d["Source géolocalisation"] == "Non géolocalisé"),
    ("A11", "Stock saisi mais PU ≥ 500 000 FCFA", "Haute", "Reclasser en immobilisation", lambda d: (d["Nature saisie"] == "Stock") & (d["Classification"] == "Immobilisation")),
    ("A12", "Immobilisation saisie mais PU < 500 000 FCFA", "Moyenne", "Reclasser en stock", lambda d: (d["Nature saisie"] == "Immobilisation") & (d["Classification"] == "Stock")),
    ("A13", "Libellé non harmonisé", "Faible", "Enrichir le dictionnaire de libellés", lambda d: ~d["_libelle_resolu"]),
    ("A14", "Doublon probable sur identifiant", "Haute", "Contrôler les doublons sur identifiant unique", lambda d: d["Identifiant"].notna() & d.duplicated(subset=["Identifiant"], keep=False)),
    ("A15", "Doublon probable par signature", "Moyenne", "Contrôler les doublons par signature composite", lambda d: d.duplicated(subset=["Libellé standardisé", "Structure agrégée", "Ville", "Description", "Prix unitaire", "Valeur nette comptable"], keep=False) & d["Prix unitaire"].notna()),
]
POIDS_PRIORITE = {"Haute": 3, "Moyenne": 2, "Faible": 1}


def detecter_anomalies(df: pd.DataFrame):
    d = df.copy()
    d["Anomalies détectées"] = ""
    d["Action recommandée principale"] = "Aucune action prioritaire"
    d["Priorité maximale"] = "Aucune"
    d["Nb anomalies"] = 0
    blocs = []

    for code, libelle, priorite, action, fn in REGLES_ANOMALIES:
        try:
            m = fn(d).fillna(False)
        except Exception as exc:
            raise RuntimeError(f"Contrôle {code} impossible : {libelle}") from exc
        if not m.any():
            continue
        sub = d.loc[m]
        blocs.append(pd.DataFrame({
            "Code": code,
            "Type d'anomalie": libelle,
            "Priorité": priorite,
            "Action recommandée": action,
            "Bien concerné": sub["Libellé standardisé"],
            "Identifiant": sub["Identifiant"],
            "Ministère": sub["Ministère"],
            "Région": sub["Région"],
            "Administration": sub["Administration"],
            "Structure": sub["Structure agrégée"],
            "État du bien": sub["État du bien"],
            "VNC exposée (FCFA)": sub["Valeur nette comptable"].fillna(0),
            "_index_ligne": sub.index,
        }))
        d.loc[m, "Nb anomalies"] += 1
        d.loc[m, "Anomalies détectées"] = d.loc[m, "Anomalies détectées"] + code + "; "
        update = d.loc[m, "Priorité maximale"].map(lambda x: POIDS_PRIORITE.get(x, 0)) < POIDS_PRIORITE[priorite]
        idx = d.loc[m].loc[update].index
        d.loc[idx, "Priorité maximale"] = priorite
        d.loc[idx, "Action recommandée principale"] = action

    anomalies = pd.concat(blocs, ignore_index=True) if blocs else pd.DataFrame(columns=[
        "Code", "Type d'anomalie", "Priorité", "Action recommandée", "Bien concerné", "Identifiant",
        "Ministère", "Région", "Administration", "Structure", "État du bien", "VNC exposée (FCFA)", "_index_ligne",
    ])
    d["Anomalies détectées"] = d["Anomalies détectées"].str.rstrip("; ").replace("", "Aucune")
    return d, anomalies


def construire_plan_action(anomalies: pd.DataFrame) -> pd.DataFrame:
    """Score de priorité (0-100), transparent et auditable :
    50 % gravité + 35 % VNC exposée (échelle log) + 15 % volume (échelle log)."""
    if anomalies.empty:
        return pd.DataFrame(columns=["Priorité", "Action recommandée", "Type d'anomalie", "Biens concernés", "VNC exposée (FCFA)", "Zones principales", "Score"])
    plan = anomalies.groupby(["Action recommandée", "Type d'anomalie", "Priorité"], as_index=False).agg(
        **{"Biens concernés": ("Code", "size"), "VNC exposée (FCFA)": ("VNC exposée (FCFA)", "sum")}
    )
    zones = []
    for _, row in plan.iterrows():
        s = anomalies[(anomalies["Action recommandée"] == row["Action recommandée"]) & (anomalies["Type d'anomalie"] == row["Type d'anomalie"])]
        vc = s["Région"].fillna("Non renseigné").value_counts().head(2)
        zones.append(", ".join([f"{k} ({v})" for k, v in vc.items()]))
    plan["Zones principales"] = zones
    grav = plan["Priorité"].map(POIDS_PRIORITE).fillna(1) / 3
    vnc = np.log10(1 + plan["VNC exposée (FCFA)"].clip(lower=0))
    vnc_n = vnc / vnc.max() if vnc.max() > 0 else 0
    vol = np.log10(1 + plan["Biens concernés"])
    vol_n = vol / vol.max() if vol.max() > 0 else 0
    plan["Score"] = (100 * (0.50 * grav + 0.35 * vnc_n + 0.15 * vol_n)).round(1)
    return plan[["Priorité", "Action recommandée", "Type d'anomalie", "Biens concernés", "VNC exposée (FCFA)", "Zones principales", "Score"]].sort_values("Score", ascending=False)

# =============================================================================
# 5. INDICATEURS ET GRAPHIQUES
# =============================================================================


def synthese_kpi(d: pd.DataFrame, anomalies: pd.DataFrame) -> dict:
    qte = d["Quantité"].fillna(0)
    n_biens = float(qte.sum())
    vnc_tot = float(d["Valeur nette comptable"].sum(skipna=True))
    va_tot = float(d["Valeur d'acquisition"].sum(skipna=True))
    n_biens_immo = float(d.loc[d["Classification"] == "Immobilisation", "Quantité"].sum(skipna=True))
    n_biens_stock = float(d.loc[d["Classification"] == "Stock", "Quantité"].sum(skipna=True))
    vnc_immo = float(d.loc[d["Classification"] == "Immobilisation", "Valeur nette comptable"].sum(skipna=True))
    vnc_stock = float(d.loc[d["Classification"] == "Stock", "Valeur nette comptable"].sum(skipna=True))
    return {
        "n_biens": n_biens,
        "n_lignes": len(d),
        "vnc_totale": vnc_tot,
        "va_totale": va_tot,
        "amort_total": float(d["Amortissement"].sum(skipna=True)),
        "n_a_reformer": float(d.loc[d["État du bien"] == "À réformer", "Quantité"].sum()),
        "n_reformes": float(d.loc[d["État du bien"] == "Réformé", "Quantité"].sum()),
        "part_a_reformer": float(d.loc[d["État du bien"] == "À réformer", "Quantité"].sum()) / n_biens if n_biens else 0,
        "vnc_immo": vnc_immo,
        "vnc_stock": vnc_stock,
        "part_immo_vnc": vnc_immo / vnc_tot if vnc_tot else 0,
        "part_stock_vnc": vnc_stock / vnc_tot if vnc_tot else 0,
        "part_immo_biens": n_biens_immo / n_biens if n_biens else 0,
        "part_stock_biens": n_biens_stock / n_biens if n_biens else 0,
        "indice_sante": indice_sante(d),
        "n_neufs": float(d.loc[d["État du bien"] == "Neuf", "Quantité"].sum()),
        "n_bon_etat": float(d.loc[d["État du bien"] == "Bon état", "Quantité"].sum()),
        "part_neufs": float(d.loc[d["État du bien"] == "Neuf", "Quantité"].sum()) / n_biens if n_biens else 0,
        "part_bon_etat": float(d.loc[d["État du bien"] == "Bon état", "Quantité"].sum()) / n_biens if n_biens else 0,
        "n_etat_impute": int(d["État imputé"].sum()) if "État imputé" in d.columns else 0,
        "taux_amortissement": (float(d["Amortissement"].sum(skipna=True)) / va_tot) if va_tot else 0,
        "n_structures": int(d["Structure agrégée"].nunique()),
        "n_regions": int(d.loc[d["Région"] != "Non renseigné", "Région"].nunique()),
        "n_ministeres": int(d["Ministère"].nunique()),
        "n_anomalies": len(anomalies),
        "n_anomalies_hautes": int((anomalies["Priorité"] == "Haute").sum()) if len(anomalies) else 0,
        "vnc_exposee": float(d.loc[d.index.isin(anomalies["_index_ligne"]), "Valeur nette comptable"].sum()) if len(anomalies) else 0,
    }


def kpi_par_ministere(d: pd.DataFrame, anomalies: pd.DataFrame) -> pd.DataFrame:
    """Tableau de bord ministère par ministère, avec la ligne « Ensemble »
    pour le global. Toutes les colonnes sont additives sauf l'indice de santé,
    recalculé sur chaque sous-population (une moyenne de moyennes pondérée à
    tort donnerait un indice faux)."""
    lignes = []
    anomalies_par_min = (
        anomalies.groupby("Ministère").size() if len(anomalies)
        else pd.Series(dtype=int)
    )

    def _ligne(nom, sub, n_anos):
        n = float(sub["Quantité"].sum())
        va = float(sub["Valeur d'acquisition"].sum(skipna=True))
        ind = indice_sante(sub)
        return {
            "Ministère": nom,
            "Libellé": REFERENTIEL_MINISTERES.get(nom, "—" if nom == "Ensemble du périmètre" else "Non référencé"),
            "Biens": n,
            "Lignes": len(sub),
            "VNC (FCFA)": float(sub["Valeur nette comptable"].sum(skipna=True)),
            "Valeur d'acquisition (FCFA)": va,
            "Amortissement (FCFA)": float(sub["Amortissement"].sum(skipna=True)),
            "Neufs (%)": round(100 * float(sub.loc[sub["État du bien"] == "Neuf", "Quantité"].sum()) / n, 1) if n else 0.0,
            "Bon état (%)": round(100 * float(sub.loc[sub["État du bien"] == "Bon état", "Quantité"].sum()) / n, 1) if n else 0.0,
            "À réformer (%)": round(100 * float(sub.loc[sub["État du bien"] == "À réformer", "Quantité"].sum()) / n, 1) if n else 0.0,
            "Indice santé": round(ind, 1) if ind is not None else np.nan,
            "Structures": int(sub["Structure agrégée"].nunique()),
            "Régions couvertes": int(sub.loc[sub["Région"] != "Non renseigné", "Région"].nunique()),
            "Anomalies": int(n_anos),
        }

    for nom, sub in d.groupby("Ministère", dropna=False):
        lignes.append(_ligne(str(nom), sub, anomalies_par_min.get(nom, 0)))
    t = pd.DataFrame(lignes).sort_values("VNC (FCFA)", ascending=False)
    if len(t) > 1:
        t = pd.concat([t, pd.DataFrame([_ligne("Ensemble du périmètre", d, len(anomalies))])],
                      ignore_index=True)
    return t.reset_index(drop=True)


# Métrique cartographiée -> fonction de formatage de la valeur affichée.
METRIQUES_CARTE = {
    "Nombre de biens": fmt_int,
    "Valeur nette comptable (FCFA)": fmt_fcfa,
    "Valeur d'acquisition (FCFA)": fmt_fcfa,
    "Biens à réformer": fmt_int,
    "À réformer (%)": lambda v: f"{float(v):.1f} %".replace(".", ","),
    "Indice de santé": lambda v: "n.d." if pd.isna(v) else f"{float(v):.1f}/100".replace(".", ","),
    "Nombre de lignes d'inventaire": fmt_int,
    "Structures": fmt_int,
}


def agreger_par_region(d: pd.DataFrame, colonne_region: str,
                       nom_colonne: str = "Région") -> pd.DataFrame:
    """Agrégat régional complet servant à la fois la choroplèthe, sa légende
    et le tableau de statistiques associé."""
    t = d.copy()
    t["_a_reformer"] = np.where(t["État du bien"] == "À réformer", t["Quantité"].fillna(0), 0)
    g = t.groupby(colonne_region, dropna=False)
    out = pd.DataFrame({
        "Nombre de biens": g["Quantité"].sum(),
        "Nombre de lignes d'inventaire": g.size(),
        "Valeur nette comptable (FCFA)": g["Valeur nette comptable"].sum(),
        "Valeur d'acquisition (FCFA)": g["Valeur d'acquisition"].sum(),
        "Biens à réformer": g["_a_reformer"].sum(),
        "Structures": g["Structure agrégée"].nunique(),
    }).fillna(0)
    out["Indice de santé"] = indice_sante_par(t, colonne_region).round(1)
    n = out["Nombre de biens"].replace(0, np.nan)
    out["À réformer (%)"] = (100 * out["Biens à réformer"] / n).round(1).fillna(0)
    out.index.name = nom_colonne
    return out.reset_index()


def statistiques_descriptions(d: pd.DataFrame) -> pd.DataFrame:
    """Une ligne par description, sur le seul périmètre transmis."""
    t = d.copy()
    t["Description"] = t["Description"].astype("string").str.strip().replace("", pd.NA)
    # Désignation du classeur, conservée dans « Libellé source » au chargement.
    # Une description peut couvrir plusieurs désignations : toutes sont listées
    # sans dédoubler les lignes ni leurs montants.
    t["Désignation"] = t["Libellé source"].astype("string").str.strip().replace("", pd.NA)

    def regrouper_designations(s):
        valeurs = sorted(s.dropna().unique().tolist())
        if s.isna().any():
            valeurs.append("(Désignation non renseignée)")
        return " ; ".join(valeurs)

    resultat = t.groupby("Description", dropna=False, sort=True).agg(**{
        "Désignation": ("Désignation", regrouper_designations),
        "Lignes d'inventaire": ("Description", "size"),
        "Quantité totale": ("Quantité", "sum"),
        "Valeur d'acquisition totale (FCFA)": ("Valeur d'acquisition", lambda s: s.sum(min_count=1)),
        "VNC totale (FCFA)": ("Valeur nette comptable", lambda s: s.sum(min_count=1)),
    }).reset_index()
    resultat["Description"] = resultat["Description"].fillna("(Description non renseignée)")
    # La désignation ouvre le tableau, et le tri la suit pour que les biens de
    # même désignation se lisent ensemble.
    autres = [c for c in resultat.columns if c not in ("Désignation", "Description")]
    return (resultat[["Désignation", "Description"] + autres]
            .sort_values(["Désignation", "Description"], kind="stable")
            .reset_index(drop=True))


def stats_par(d: pd.DataFrame, dim: str) -> pd.DataFrame:
    g = d.groupby(dim, dropna=False)
    out = pd.DataFrame({
        "Nombre de biens": g["Quantité"].sum(),
        "VNC totale (FCFA)": g["Valeur nette comptable"].sum(),
        "Valeur d'acquisition (FCFA)": g["Valeur d'acquisition"].sum(),
        "Anomalies": g["Nb anomalies"].sum(),
    }).fillna(0)
    out["Indice de santé"] = indice_sante_par(d, dim).round(1)
    return out.sort_values("VNC totale (FCFA)", ascending=False)


def repartition_etats(d: pd.DataFrame, dim=None):
    if dim is None:
        s = d.groupby("État du bien")["Quantité"].sum()
        return s.reindex([e for e in ORDRE_ETATS if e in s.index]).to_frame("Nombre de biens")
    p = d.pivot_table(index=dim, columns="État du bien", values="Quantité", aggfunc="sum", fill_value=0)
    return p[[c for c in ORDRE_ETATS if c in p.columns]]


def comparaison_administrations(d: pd.DataFrame) -> pd.DataFrame:
    """Encart Central vs Déconcentré (question n°4 du cahier des charges)."""
    lignes = []
    vnc_tot = float(d["Valeur nette comptable"].sum(skipna=True)) or 1.0
    for admin, sub in d.groupby("Administration"):
        n = float(sub["Quantité"].sum()) or 1.0
        vnc_a = float(sub["Valeur nette comptable"].sum(skipna=True))
        va_a = float(sub["Valeur d'acquisition"].sum(skipna=True))
        ind = indice_sante(sub)
        lignes.append({
            "Administration": admin,
            "Biens": float(sub["Quantité"].sum()),
            # Valeur d'acquisition affichée avant la VNC : les deux montants se
            # lisent ensemble, l'écart entre eux étant l'amortissement cumulé.
            "Valeur d'acquisition (FCFA)": va_a,
            "VNC (FCFA)": vnc_a,
            "Part VNC (%)": round(100 * vnc_a / vnc_tot, 1),
            "À réformer (%)": round(100 * float(sub.loc[sub["État du bien"] == "À réformer", "Quantité"].sum()) / n, 1),
            "Indice santé": round(ind, 1) if ind is not None else np.nan,
            "Anomalies": int(sub["Nb anomalies"].sum()),
        })
    t = pd.DataFrame(lignes).set_index("Administration")
    ordre = ["Centrale", "Déconcentrée", "Non renseigné"]
    return t.reindex([a for a in ordre if a in t.index])


TOTAL_MINISTERE = "Total du ministère"
MESURES_CENTRALE = ["VNC (FCFA)", "Valeur d'acquisition (FCFA)", "Biens"]


def statistiques_structures_centrales(d: pd.DataFrame) -> pd.DataFrame:
    """Administration centrale : une ligne par structure, regroupée par
    ministère de tutelle, précédée du total de ce ministère.

    Les services centraux siègent tous à Yaoundé : la carte les empile sur un
    seul point. Leur lecture passe donc par l'organigramme (ministère, puis
    direction ou service) et non par la géographie."""
    colonnes = ["Ministère", "Structure", "Lignes d'inventaire", "Biens",
                "Valeur d'acquisition (FCFA)", "VNC (FCFA)",
                "Part de la VNC du ministère (%)", "À réformer (%)", "Indice santé"]
    c = d[d["Administration"] == "Centrale"]
    if c.empty:
        return pd.DataFrame(columns=colonnes)

    def ligne(ministere, structure, sub, vnc_ministere):
        n = float(sub["Quantité"].sum())
        vnc = float(sub["Valeur nette comptable"].sum(skipna=True))
        a_reformer = float(sub.loc[sub["État du bien"] == "À réformer", "Quantité"].sum())
        ind = indice_sante(sub)
        return {
            "Ministère": ministere,
            "Structure": structure,
            "Lignes d'inventaire": len(sub),
            "Biens": n,
            "Valeur d'acquisition (FCFA)": float(sub["Valeur d'acquisition"].sum(skipna=True)),
            "VNC (FCFA)": vnc,
            "Part de la VNC du ministère (%)": (round(100 * vnc / vnc_ministere, 1)
                                                if vnc_ministere else np.nan),
            "À réformer (%)": round(100 * a_reformer / n, 1) if n else 0.0,
            "Indice santé": round(ind, 1) if ind is not None else np.nan,
        }

    lignes = []
    for ministere, bloc in c.groupby("Ministère", sort=True):
        vnc_ministere = float(bloc["Valeur nette comptable"].sum(skipna=True))
        lignes.append(ligne(ministere, TOTAL_MINISTERE, bloc, vnc_ministere))
        structures = [ligne(ministere, structure, sub, vnc_ministere)
                      for structure, sub in bloc.groupby("Structure agrégée", dropna=False)]
        lignes.extend(sorted(structures, key=lambda l: l["VNC (FCFA)"], reverse=True))
    return pd.DataFrame(lignes, columns=colonnes)


def treemap_structures_centrales(tableau: pd.DataFrame, mesure: str, titre=None):
    """Treemap à deux niveaux : un rectangle par ministère, découpé par
    structure, de surface proportionnelle à la mesure choisie."""
    detail = tableau[(tableau["Structure"] != TOTAL_MINISTERE) & (tableau[mesure] > 0)]
    if detail.empty:
        return apply_layout(go.Figure(), 360, titre)
    monetaire = "FCFA" in mesure
    court = fmt_compact if monetaire else fmt_int
    complet = fmt_fcfa_complet if monetaire else (lambda v: f"{fmt_int(v)} biens")
    # Une teinte par ministère ; ses structures la partagent pour que les deux
    # niveaux se lisent d'un seul regard.
    reserve = [THEME["or"], THEME["bleu"], "#7A5A07", "#5B6B4F", "#8C4A2F", "#6B7280"]
    ids, labels, parents, valeurs, couleurs, donnees = [], [], [], [], [], []
    for rang, (ministere, bloc) in enumerate(detail.groupby("Ministère", sort=False)):
        teinte = COULEURS_MINISTERE.get(ministere, reserve[rang % len(reserve)])
        total = float(bloc[mesure].sum())
        ids.append(ministere)
        labels.append(ministere)
        parents.append("")
        valeurs.append(total)
        couleurs.append(teinte)
        donnees.append([court(total), complet(total),
                        f"{len(bloc)} structure(s) · "
                        + REFERENTIEL_MINISTERES.get(ministere, "ministère non référencé")])
        for _, l in bloc.iterrows():
            part = f"{100 * l[mesure] / total:.1f}".replace(".", ",") if total else "0"
            ids.append(f"{ministere}/{l['Structure']}")
            labels.append(str(l["Structure"]))
            parents.append(ministere)
            valeurs.append(float(l[mesure]))
            couleurs.append(teinte)
            donnees.append([court(l[mesure]), complet(l[mesure]),
                            f"{part} % du {ministere} · {fmt_int(l['Biens'])} biens"])
    fig = go.Figure(go.Treemap(
        ids=ids, labels=labels, parents=parents, values=valeurs, branchvalues="total",
        customdata=donnees,
        texttemplate="<b>%{label}</b><br>%{customdata[0]}",
        hovertemplate=("<b>%{label}</b><br>%{customdata[1]}<br>%{customdata[2]}"
                       "<extra></extra>"),
        marker=dict(colors=couleurs, line=dict(color="#FFFFFF", width=1.5)),
        tiling=dict(pad=2), pathbar=dict(visible=True),
    ))
    return apply_layout(fig, 600, titre)


TYPES_GRAPHIQUES = ["Barres horizontales", "Barres verticales", "Anneau",
                    "Camembert", "Treemap"]


def figure_repartition(d, dimension, mesure, type_graphique, couleur, limite,
                       titre, monetaire=False):
    """Répartition d'une mesure selon une dimension, dans la forme demandée.

    Le choix de forme n'est pas cosmétique : les barres comparent des
    grandeurs entre elles, l'anneau et le camembert montrent des parts d'un
    tout, le treemap fait les deux mais se lit mal en dessous de quelques
    unités. La fonction produit la même donnée dans les cinq formes, avec la
    valeur toujours écrite à côté de la couleur (lecture possible sans
    distinguer les teintes).

    `limite` vaut « Toutes » par défaut : aucun classement n'est tronqué
    silencieusement, une modalité absente du graphique serait lue comme une
    modalité sans patrimoine.
    """
    agg = (d.groupby(dimension, dropna=False, as_index=False)
           .agg(Valeur=(mesure, "sum")))
    agg[dimension] = agg[dimension].fillna("Non renseigné").astype(str)
    agg = agg[agg["Valeur"].notna()].sort_values("Valeur", ascending=False)
    if limite != "Toutes":
        agg = agg.head(int(limite))
    if agg.empty:
        return apply_layout(go.Figure(), 360, titre)

    formateur = fmt_compact if monetaire else fmt_int
    agg["Étiquette"] = agg["Valeur"].map(formateur)
    agg["Survol"] = agg["Valeur"].map(fmt_fcfa_complet if monetaire else fmt_int)
    total = float(agg["Valeur"].sum()) or 1.0
    agg["Part"] = (100 * agg["Valeur"] / total).map(
        lambda x: f"{x:.1f}".replace(".", ",") + " %")
    # Hauteur proportionnelle au nombre de modalités : avec 58 départements,
    # une hauteur fixe écraserait les barres jusqu'à l'illisible.
    hauteur = min(max(380, 24 * len(agg) + 140), 1400)

    if type_graphique == "Barres horizontales":
        fig = px.bar(agg.sort_values("Valeur"), y=dimension, x="Valeur",
                     orientation="h", text="Étiquette",
                     color_discrete_sequence=[couleur])
        fig.update_traces(textposition="outside")
    elif type_graphique == "Barres verticales":
        fig = px.bar(agg, x=dimension, y="Valeur", text="Étiquette",
                     color_discrete_sequence=[couleur])
        fig.update_traces(textposition="outside")
        hauteur = min(max(420, 12 * len(agg) + 260), 900)
    elif type_graphique in ("Anneau", "Camembert"):
        fig = px.pie(agg, names=dimension, values="Valeur",
                     hole=0.55 if type_graphique == "Anneau" else 0)
        fig.update_traces(
            text=agg["Part"], textinfo="text",
            customdata=agg["Survol"],
            hovertemplate="%{label}<br>%{customdata} · %{text}<extra></extra>")
        if type_graphique == "Anneau":
            fig.add_annotation(
                text=f"<b>{formateur(total)}</b><br>"
                     f"<span style='font-size:11px'>total</span>",
                showarrow=False, font=dict(size=17, family="Archivo",
                                           color=THEME["texte"]))
        hauteur = min(max(420, 10 * len(agg) + 320), 900)
    else:  # Treemap
        fig = px.treemap(agg, path=[dimension], values="Valeur")
        fig.update_traces(
            texttemplate="%{label}<br>%{customdata[0]}",
            customdata=agg[["Étiquette", "Survol"]].values,
            hovertemplate="%{label}<br>%{customdata[1]}<extra></extra>")
        hauteur = min(max(420, 9 * len(agg) + 320), 900)

    fig.update_layout(yaxis_title=None, xaxis_title=None)
    return apply_layout(fig, hauteur, titre)


def apply_layout(fig, height=360, title=None):
    fig.update_layout(
        title=title,
        height=height,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color=THEME["texte"], size=13),
        title_font=dict(family="Archivo", size=15),
        margin=dict(l=10, r=10, t=50 if title else 20, b=20),
        legend=dict(orientation="h", yanchor="bottom", y=-0.25, xanchor="left", x=0),
        separators=", ",  # format FR : virgule décimale, espace pour les milliers
    )
    fig.update_xaxes(gridcolor="#E5E7EB", zerolinecolor="#E5E7EB")
    fig.update_yaxes(gridcolor="#E5E7EB", zerolinecolor="#E5E7EB")
    _survol_chiffres_complets(fig)
    return fig


def _survol_chiffres_complets(fig):
    """Infobulles des barres en toutes unités (« 900 000 », jamais « 900k »).

    Par défaut Plotly abrège les grands nombres au survol et y répète
    l'étiquette courte de la barre. Les petites valeurs (parts en %, indices)
    gardent leurs décimales."""
    for trace in fig.data:
        if trace.type != "bar":
            continue
        horizontal = trace.orientation == "h"
        valeurs = pd.to_numeric(pd.Series(list(trace.x if horizontal else trace.y)),
                                errors="coerce")
        if not (valeurs.abs() >= 1000).any():
            continue
        if horizontal:
            fig.update_xaxes(hoverformat=",.0f")
        else:
            fig.update_yaxes(hoverformat=",.0f")
        if trace.hovertemplate:
            trace.hovertemplate = re.sub(r"<br>[^<>]*=%\{text\}", "", trace.hovertemplate)


def jauge_sante(v, height=250):
    """Élément signature du dashboard : jauge de l'indice de santé
    patrimoniale avec les quatre seuils du barème (>=80 sain, 60-79
    vigilance, 40-59 critique, <40 intervention urgente)."""
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=round(v, 1),
        number=dict(suffix=" / 100", font=dict(size=32, family="Archivo",
                                               color=couleur_sante(v))),
        gauge=dict(
            axis=dict(range=[0, 100], tickvals=[0, 40, 60, 80, 100],
                      tickfont=dict(size=11)),
            bar=dict(color=couleur_sante(v), thickness=0.30),
            borderwidth=0,
            steps=[
                dict(range=[0, 40], color="#F6DAD5"),
                dict(range=[40, 60], color="#FBE7D0"),
                dict(range=[60, 80], color="#F3EAC9"),
                dict(range=[80, 100], color="#D5EBDC"),
            ],
        ),
    ))
    fig.update_layout(height=height, margin=dict(l=24, r=24, t=28, b=8),
                      paper_bgcolor="rgba(0,0,0,0)",
                      font=dict(family="Inter", color=THEME["texte"]),
                      separators=", ")
    return fig



# =============================================================================
# 7. FIGURES CARTOGRAPHIQUES
# =============================================================================

# Plotly < 5.24 ne connaît que les tracés « mapbox » ; les versions récentes
# leur préfèrent « map ». On s'adapte pour rester compatible des deux côtés.
USE_MAPLIBRE = hasattr(go, "Scattermap")


def trace_map(**kw):
    return go.Scattermap(**kw) if USE_MAPLIBRE else go.Scattermapbox(**kw)


# Fond de carte : « white-bg » = fond uni, sans la moindre requête réseau.
# Les styles à tuiles (carto-positron, open-street-map…) téléchargent leurs
# images chez un prestataire externe ; CARTO ayant rendu ses fonds payants,
# ses serveurs renvoient depuis peu des tuiles filigranées « API KEY
# REQUIRED » (HTTP 200, donc aucune erreur détectable côté application), et
# les serveurs OpenStreetMap refusent les usages applicatifs. Le découpage
# administratif venant de notre propre GeoJSON, le fond uni suffit : la carte
# devient autonome, insensible aux conditions commerciales d'un tiers et
# fonctionne sans accès Internet.
FOND_CARTE = "white-bg"

def trace_etiquettes(noms, centres, couleur="#3F3A2B", taille=11, gras=False):
    """Noms des entités, posés à leur centre.

    Plotly ne sait pas étiqueter un polygone : sans cette couche de texte, un
    découpage n'est qu'un réseau de traits muets. Les entités absentes du
    dictionnaire de centres sont ignorées silencieusement — mieux vaut une
    étiquette manquante qu'un nom posé au mauvais endroit.
    """
    points = [(n, centres[n]) for n in noms if n in (centres or {})]
    if not points:
        return None
    return trace_map(
        lat=[c["lat"] for _, c in points],
        lon=[c["lon"] for _, c in points],
        mode="text",
        text=[n for n, _ in points],
        textfont=dict(size=taille, color=couleur,
                      family="Archivo" if gras else "Inter"),
        hoverinfo="skip", showlegend=False,
    )


def carte_choroplethe(agg, geojson, metrique, region_selectionnee, mode_couleur,
                      hauteur=680, colonne="Région", ordre=None, couleurs=None,
                      centres=None, sous_decoupage=None, centres_sous=None,
                      afficher_noms=True, noms_sous=None):
    """Choroplèthe d'un niveau administratif quelconque.

    - Vue d'ensemble : toutes les entités du niveau sont tracées ; en mode « une
      couleur par entité » chacune reçoit sa teinte (une trace par entité), en
      mode « par valeur » une échelle séquentielle unique permet de comparer
      les volumes d'un coup d'œil.
    - Une entité sélectionnée : elle seule est dessinée et la carte est
      recentrée dessus — les autres ne sont pas grisées mais retirées.

    `sous_decoupage` superpose en trait fin les limites du niveau inférieur
    (départements sous les régions, arrondissements sous les départements) :
    le découpage administratif reste visible même quand la base ne permet pas
    de le chiffrer.
    """
    valeurs = agg.set_index(colonne)[metrique].to_dict()
    formateur = METRIQUES_CARTE.get(metrique, fmt_int)
    # Au survol, les montants s'affichent en toutes unités.
    if formateur is fmt_fcfa:
        formateur = fmt_fcfa_complet
    disponibles = {f["id"] for f in geojson["features"]}
    ordre = ordre or ORDRE_REGIONS
    couleurs = couleurs or COULEURS_REGIONS
    entites_tracees = ([region_selectionnee] if region_selectionnee != TOUT_LE_CAMEROUN
                       else [r for r in ordre if r in disponibles])

    fig = go.Figure()

    # Sous-découpage d'abord : il doit passer sous les aplats colorés.
    if sous_decoupage is not None and sous_decoupage.get("features"):
        ids_sous = [f["id"] for f in sous_decoupage["features"]]
        fig.add_trace(go.Choroplethmapbox(
            geojson=sous_decoupage, locations=ids_sous, z=[0] * len(ids_sous),
            showscale=False, colorscale=[[0, "#FBF8F0"], [1, "#FBF8F0"]], zmin=0, zmax=1,
            marker=dict(line=dict(color="#B9AC85", width=0.7)), marker_opacity=0.55,
            hovertemplate="<b>%{location}</b><extra></extra>", showlegend=False,
        ))

    if mode_couleur == "Par valeur (dégradé)" and region_selectionnee == TOUT_LE_CAMEROUN:
        sous_geo = {"type": "FeatureCollection",
                    "features": [f for f in geojson["features"] if f["id"] in entites_tracees]}
        series = pd.Series({r: float(valeurs.get(r, 0) or 0) for r in entites_tracees})
        fig.add_trace(go.Choroplethmapbox(
            geojson=sous_geo, locations=series.index.tolist(), z=series.values,
            colorscale=ECHELLE_CHOROPLETHE, marker=dict(line=dict(color="#FFFFFF", width=1.4)),
            marker_opacity=0.88,
            colorbar=dict(title=dict(text=metrique, side="right"), thickness=14, len=0.6,
                          x=0.99, xanchor="right", y=0.5),
            customdata=[[formateur(v)] for v in series.values],
            hovertemplate="<b>%{location}</b><br>" + metrique + " : %{customdata[0]}<extra></extra>",
        ))
    else:
        for entite in entites_tracees:
            # Toutes les parties de l'entité, pas seulement la première : un
            # département en deux morceaux doit être colorié en entier.
            morceaux = [f for f in geojson["features"] if f["id"] == entite]
            if not morceaux:
                continue
            couleur = couleurs.get(entite, THEME["or"])
            val = valeurs.get(entite, 0)
            fig.add_trace(go.Choroplethmapbox(
                geojson={"type": "FeatureCollection", "features": morceaux},
                locations=[entite] * len(morceaux), z=[1] * len(morceaux), showscale=False,
                colorscale=[[0, couleur], [1, couleur]], zmin=0, zmax=1,
                marker=dict(line=dict(color="#FFFFFF", width=1.6)),
                marker_opacity=0.92 if entite == region_selectionnee else 0.80,
                name=entite, customdata=[[formateur(val)]],
                hovertemplate="<b>%{location}</b><br>" + metrique + " : %{customdata[0]}<extra></extra>",
            ))

    # Étiquettes en dernier : elles doivent passer au-dessus des aplats.
    if sous_decoupage is not None and noms_sous and centres_sous:
        etiquettes = trace_etiquettes(sorted(set(noms_sous)), centres_sous,
                                      couleur="#6F6854", taille=9)
        if etiquettes is not None:
            fig.add_trace(etiquettes)
    if afficher_noms:
        etiquettes = trace_etiquettes(entites_tracees, centres or {},
                                      couleur="#2B2418", taille=11, gras=True)
        if etiquettes is not None:
            fig.add_trace(etiquettes)

    if region_selectionnee != TOUT_LE_CAMEROUN:
        # Le cadrage vient du fond lui-même : un arrondissement doit être vu
        # de bien plus près qu'une région.
        repere = (centres or {}).get(region_selectionnee)
        if repere:
            centre = dict(lat=repere["lat"], lon=repere["lon"])
            zoom = repere["zoom"]
        else:
            lat0, lon0 = COORD_REGIONS.get(region_selectionnee, (6.0, 12.3))
            centre, zoom = dict(lat=lat0, lon=lon0), 6.3
    else:
        centre, zoom = dict(lat=6.6, lon=12.5), 5.15
    fig.update_layout(
        mapbox=dict(style=FOND_CARTE, center=centre, zoom=zoom),
        height=hauteur, margin=dict(l=0, r=0, t=8, b=0), showlegend=False,
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color=THEME["texte"], size=13),
        hoverlabel=dict(bgcolor="#FFFFFF", bordercolor=THEME["vert"],
                        font=dict(size=13, color=THEME["texte"])),
        separators=", ",
    )
    return fig


def donnees_legende_regions(agg, metrique, region_selectionnee, mode_couleur,
                            colonne="Région", ordre=None, couleurs=None,
                            singulier="région", feminin=True):
    """Données de la légende cartographique, indépendantes du rendu.

    La légende native de Plotly ne sait ni afficher la valeur à côté du nom de
    région, ni mettre en évidence la région sélectionnée : chaque interface
    dessine donc sa propre légende à partir d'ici (HTML pour Streamlit,
    composants pour Dash), sans dupliquer le calcul.

    Retourne (titre, sous-titre, lignes) où chaque ligne est un dictionnaire
    {region, couleur, valeur, actif}.
    """
    formateur = METRIQUES_CARTE.get(metrique, fmt_int)
    colonne = colonne or "Région"
    ordre = ordre or ORDRE_REGIONS
    couleurs = couleurs or COULEURS_REGIONS
    t = agg.set_index(colonne)
    entites = ([region_selectionnee] if region_selectionnee != TOUT_LE_CAMEROUN
               else [r for r in ordre if r in t.index])
    maxi = max([float(t[metrique].get(r, 0) or 0) for r in entites] or [0]) or 1.0
    lignes = []
    for entite in entites:
        val = t[metrique].get(entite, 0)
        if mode_couleur == "Par valeur (dégradé)" and region_selectionnee == TOUT_LE_CAMEROUN:
            ratio = float(val or 0) / maxi
            couleur = ECHELLE_CHOROPLETHE[min(int(ratio * 4), 4)][1]
        else:
            couleur = couleurs.get(entite, THEME["or"])
        lignes.append({
            "region": entite,
            "couleur": couleur,
            "valeur": formateur(val),
            "actif": entite == region_selectionnee,
        })
    if region_selectionnee != TOUT_LE_CAMEROUN:
        titre = (f"{singulier.capitalize()} sélectionné{'e' if feminin else ''} : "
                 f"{region_selectionnee}")
    else:
        titre = f"Tout le Cameroun — une couleur par {singulier}"
    return titre, metrique, lignes


def layout_map(fig, height=690, zoom=5.2, center=None, style=FOND_CARTE):
    conf = dict(style=style, center=center or dict(lat=6.0, lon=12.3), zoom=zoom)
    fig.update_layout(
        **({"map": conf} if USE_MAPLIBRE else {"mapbox": conf}),
        height=height,
        margin=dict(l=0, r=0, t=40, b=0),
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color=THEME["texte"], size=13),
        title_font=dict(family="Archivo", size=15),
        legend=dict(bgcolor="rgba(255,255,255,.92)", bordercolor=THEME["bordure"], borderwidth=1, x=0.01, y=0.01, xanchor="left", yanchor="bottom"),
        hoverlabel=dict(bgcolor="#FFFFFF", bordercolor=THEME["vert"], font=dict(size=13, color=THEME["texte"])),
        separators=", ",
    )
    return fig




# Hover d'audit complet : chaque point doit permettre un contrôle sans quitter
# la carte (exigence du cahier des charges).
HOVER_CARTE_DETAIL = (
    "<b>%{customdata[0]}</b> · <i>%{customdata[1]}</i><br>"
    "%{customdata[2]}<br>"
    "Ministère de tutelle : <b>%{customdata[21]}</b><br>"
    "Délégation : %{customdata[22]}<br>"
    "Structure : %{customdata[9]}<br>"
    "État : <b>%{customdata[3]}</b> · %{customdata[4]} · %{customdata[5]}<br>"
    "%{customdata[6]} · %{customdata[7]} · %{customdata[8]}<br>"
    "Qté : %{customdata[10]} · PU : %{customdata[11]}<br>"
    "VA : %{customdata[12]} · Amort. : %{customdata[13]} · <b>VNC : %{customdata[14]}</b><br>"
    "Nature saisie : %{customdata[15]} · Contrôle : %{customdata[16]}<br>"
    "Mode d'acquisition : %{customdata[17]}<br>"
    "Anomalies : %{customdata[18]} · Action : %{customdata[19]}<br>"
    "Géolocalisation : %{customdata[20]}<extra></extra>"
)


def libelle_delegation(sub):
    """Délégation de rattachement de chaque ligne, pour l'infobulle de la carte.

    Le classeur n'a pas de colonne « délégation » : la structure y vaut
    « DELEGATION REGIONALE » ou « DELEGATION DEPARTEMENTALE » sans dire
    laquelle. Le territoire est donc repris de la région ou du département
    de la ligne."""
    structure = sub["Structure agrégée"].fillna("").astype(str).map(_normaliser)
    region = sub["Région"].fillna("Non renseigné").astype(str)
    departement = sub["Département"].fillna("Non renseigné").astype(str)
    deleg = pd.Series("Non précisée dans le classeur", index=sub.index, dtype="object")
    deleg[sub["Administration"] == "Centrale"] = "Sans objet (administration centrale)"
    regionale = structure.str.contains("DELEGATION REGIONALE", na=False)
    departementale = structure.str.contains("DELEGATION DEPARTEMENTALE", na=False)
    deleg[regionale] = "Délégation régionale — " + region[regionale]
    deleg[departementale] = ("Délégation départementale — " + departement[departementale]
                             + " (" + region[departementale] + ")")
    return deleg


def custom_data_detail(sub):
    """Colonnes d'audit injectées dans l'infobulle de la carte de points."""
    c = pd.DataFrame(index=sub.index)
    c[0] = sub["Libellé standardisé"].fillna("—")
    c[1] = sub["Libellé source"].fillna("—").astype(str).str[:45]
    c[2] = sub["Description"].fillna("—").astype(str).str[:60]
    c[3] = sub["État du bien"].fillna("—")
    c[4] = sub["Administration"].fillna("—")
    c[5] = sub["Classification"].fillna("—")
    c[6] = sub["Région"].fillna("—")
    c[7] = sub["Département"].fillna("—")
    c[8] = sub["Ville"].fillna("—")
    c[9] = sub["Structure agrégée"].fillna("—").astype(str).str[:55]
    c[10] = sub["Quantité"].fillna(1).map(fmt_int)
    c[11] = sub["Prix unitaire"].map(fmt_fcfa_complet)
    c[12] = sub["Valeur d'acquisition"].map(fmt_fcfa_complet)
    c[13] = sub["Amortissement"].map(fmt_fcfa_complet)
    c[14] = sub["Valeur nette comptable"].map(fmt_fcfa_complet)
    c[15] = sub["Nature saisie"].fillna("—")
    c[16] = sub["Contrôle classification"].fillna("—")
    c[17] = sub["Mode d'acquisition"].fillna("—")
    c[18] = sub["Anomalies détectées"].fillna("Aucune")
    c[19] = sub["Action recommandée principale"].fillna("—").astype(str).str[:55]
    c[20] = sub["Source géolocalisation"].fillna("—")
    c[21] = (sub["Ministère"].fillna("—").astype(str) + " — "
             + sub["Ministère (libellé)"].fillna("libellé non renseigné").astype(str))
    c[22] = libelle_delegation(sub)
    return c.values


def disperser_points(d, rayon_max=0.20):
    """Étale les points qui partagent exactement la même coordonnée.

    La géolocalisation se fait à la ville : tous les biens d'une même localité
    tombent sur un point unique, et 2 000 biens de Yaoundé s'empilent en un
    seul pâté illisible. Chaque groupe est donc réparti sur un disque selon la
    spirale de Vogel (angle d'or), qui répartit n points uniformément sans
    jamais les aligner ni les regrouper — contrairement à un bruit aléatoire,
    qui laisse des paquets et des trous.

    Le rayon croît avec l'effectif du groupe : une ville à 5 biens reste un
    point net, une ville à 2 000 biens devient une tache dont la surface se lit
    comme une densité. La correction en cos(latitude) évite l'ovalisation du
    disque quand on s'éloigne de l'équateur. Le calcul est déterministe : la
    même base produit toujours la même carte.
    """
    d = d.dropna(subset=["lat", "lon"]).copy()
    if d.empty:
        d["lat_j"], d["lon_j"] = [], []
        return d
    lat_j = d["lat"].astype(float).to_numpy().copy()
    lon_j = d["lon"].astype(float).to_numpy().copy()
    position = {cle: i for i, cle in enumerate(d.index)}
    angle_or = np.pi * (3.0 - np.sqrt(5.0))

    for _, groupe in d.groupby(["lat", "lon"], sort=False):
        n = len(groupe)
        if n < 2:
            continue
        indices = [position[i] for i in groupe.index]
        k = np.arange(n)
        rayon = min(rayon_max, 0.012 * np.sqrt(n)) * np.sqrt((k + 0.5) / n)
        theta = k * angle_or
        lat0 = float(groupe["lat"].iloc[0])
        # Un degré de longitude vaut cos(lat) degré de latitude en distance.
        facteur_lon = max(np.cos(np.radians(lat0)), 0.2)
        lat_j[indices] = lat0 + rayon * np.cos(theta)
        lon_j[indices] = float(groupe["lon"].iloc[0]) + rayon * np.sin(theta) / facteur_lon

    d["lat_j"], d["lon_j"] = lat_j, lon_j
    return d


def repartition_etats_carte(d):
    """Effectifs et parts par état, pour la légende de la carte de points.

    Compte les biens (somme des quantités) et non les points, afin de rester
    cohérent avec tous les autres indicateurs de l'application."""
    if d.empty:
        return []
    par_etat = d.groupby("État du bien")["Quantité"].sum()
    total = float(par_etat.sum()) or 1.0
    ordre = ([e for e in ORDRE_ETATS if e in par_etat.index]
             + [e for e in par_etat.index if e not in ORDRE_ETATS])
    return [{"etat": str(e), "biens": float(par_etat[e]),
             "part": 100 * float(par_etat[e]) / total,
             "couleur": COULEURS_ETAT.get(str(e), "#6B7280")} for e in ordre]


def carte_points(d, centre, zoom, titre, hauteur=760, style=FOND_CARTE, geojson=None,
                 regions_contours=None, rayon_dispersion=0.0, centres=None,
                 afficher_noms=True, sous_decoupage=None, centres_sous=None,
                 noms_sous=None, couleurs_fond=None):
    """Carte détaillée : un point = une ligne d'inventaire, colorée par état.

    Le fond étant uni, les contours des régions sont dessinés sous les points à
    partir du GeoJSON local : le repère géographique est conservé sans dépendre
    d'un serveur de tuiles externe. `regions_contours` restreint ce tracé aux
    régions nommées — quand une zone est sélectionnée, elle s'affiche seule,
    comme sur la choroplèthe.

    La légende porte, pour chaque état, le nombre de biens et sa part, calculés
    sur les seules données affichées.
    """
    d = disperser_points(d, rayon_dispersion)

    fig = go.Figure()
    noms_fond = []

    # Sous-découpage d'abord (trait le plus fin), puis le découpage principal.
    if sous_decoupage is not None and sous_decoupage.get("features"):
        entites_sous = sous_decoupage["features"]
        ids_sous = [f["id"] for f in entites_sous]
        fig.add_trace(go.Choroplethmapbox(
            geojson={"type": "FeatureCollection", "features": entites_sous},
            locations=ids_sous, z=[0] * len(ids_sous), showscale=False,
            colorscale=[[0, "#FBF8F0"], [1, "#FBF8F0"]], zmin=0, zmax=1,
            marker=dict(line=dict(color="#B9AC85", width=0.7)), marker_opacity=0.5,
            hovertemplate="<b>%{location}</b><extra></extra>", showlegend=False,
        ))

    if geojson is not None:
        entites = geojson["features"]
        if regions_contours:
            entites = [f for f in entites if f["id"] in set(regions_contours)]
        if entites:
            noms_fond = sorted({f["id"] for f in entites})
            if couleurs_fond:
                # Une teinte pastel par territoire : les découpages se lisent
                # d'un coup d'œil sans masquer les points.
                index = {nom: i for i, nom in enumerate(noms_fond)}
                echelle = colorscale_discrete(
                    [eclaircir(couleurs_fond.get(n, THEME["or"])) for n in noms_fond])
                valeurs_z = [index[f["id"]] for f in entites]
                bornes_z = (0, len(noms_fond))
                opacite = 0.85
            else:
                echelle = [[0, "#F2EFE6"], [1, "#F2EFE6"]]
                valeurs_z = [0] * len(entites)
                bornes_z = (0, 1)
                opacite = 0.55
            fig.add_trace(go.Choroplethmapbox(
                geojson={"type": "FeatureCollection", "features": entites},
                locations=[f["id"] for f in entites], z=valeurs_z, showscale=False,
                colorscale=echelle, zmin=bornes_z[0], zmax=bornes_z[1],
                marker=dict(line=dict(color="#FFFFFF", width=1.3)),
                marker_opacity=opacite,
                hovertemplate="<b>%{location}</b><extra></extra>", showlegend=False,
            ))

    # Noms avant les points : un nom masqué par un point reste lisible, alors
    # qu'un point masqué par un nom devient incliquable.
    if sous_decoupage is not None and noms_sous and centres_sous:
        etiquettes = trace_etiquettes(sorted(set(noms_sous)), centres_sous,
                                      couleur="#6F6854", taille=9)
        if etiquettes is not None:
            fig.add_trace(etiquettes)
    if afficher_noms and noms_fond and centres:
        etiquettes = trace_etiquettes(noms_fond, centres, couleur="#2B2418",
                                      taille=11, gras=True)
        if etiquettes is not None:
            fig.add_trace(etiquettes)

    parts = {p["etat"]: p for p in repartition_etats_carte(d)}
    etats = ([e for e in ORDRE_ETATS if e in d["État du bien"].unique()]
             + [e for e in d["État du bien"].dropna().unique() if e not in ORDRE_ETATS])
    for etat in etats:
        g = d[d["État du bien"] == etat]
        if g.empty:
            continue
        # Effectif et part portés par la légende : la couleur seule ne dit pas
        # combien de biens elle représente.
        info = parts.get(str(etat))
        if info:
            part = f"{info['part']:.1f}".replace(".", ",")
            etiquette = f"{etat} · {fmt_int(info['biens'])} biens ({part} %)"
        else:
            etiquette = str(etat)
        fig.add_trace(trace_map(
            lat=g["lat_j"], lon=g["lon_j"], mode="markers",
            marker=dict(size=8, color=COULEURS_ETAT.get(str(etat), "#6B7280"),
                        opacity=0.78),
            name=etiquette, customdata=custom_data_detail(g),
            hovertemplate=HOVER_CARTE_DETAIL,
        ))
    layout_map(fig, height=hauteur, zoom=zoom, center=centre, style=style)
    fig.update_layout(title=titre)
    return fig
