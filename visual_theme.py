"""Palette institutionnelle et couleurs sémantiques de l'application."""

# Palette unique de l'état des biens, commune à la carte de points, à sa
# légende et à tous les graphiques : vert = neuf, jaune = bon état,
# orange = à réformer.
COULEURS_ETAT = {
    "Neuf": "#159947",
    "Bon état": "#F2C744",
    "À réformer": "#E8710A",
    "Réformé": "#B42318",
    "Non renseigné": "#6B7280",
}

COULEURS_ADMIN = {
    "Centrale": "#006837",
    "Déconcentrée": "#2C6E91",
    "Non renseigné": "#6B7280",
}
COULEURS_CLASSIF = {
    "Immobilisation": "#006837",
    "Stock": "#2C6E91",
    "Indéterminée": "#6B7280",
}
COULEURS_MINISTERE = {"MINTP": "#006837", "MINAC": "#2C6E91"}
COULEURS_PRIORITE = {"Haute": "#B42318", "Moyenne": "#E8710A", "Faible": "#B8860B"}
COULEURS_RISQUE = {
    "Faible": "#159947",
    "Moyen": "#E8710A",
    "Élevé": "#B42318",
    "Non renseigné": "#6B7280",
}

THEME = {
    "fond": "#FFFFFF",
    "carte": "#FFFFFF",
    "texte": "#242116",
    "texte2": "#6F6854",
    "bordure": "#D8C17A",
    "vert": "#8A6508",
    "vert2": "#00A85A",
    "or": "#9A7209",
    "or_clair": "#D7B84B",
    "or_pale": "#F4EBCB",
    "noir": "#211E16",
    "orange": "#E8710A",
    "rouge": "#B42318",
    "bleu": "#2C6E91",
}
