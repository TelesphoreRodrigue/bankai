# -*- coding: utf-8 -*-
"""
Cartographie des biens de l'État — version 7 — revue indépendante.

Lancement : streamlit run app7.py
Données   : placer le fichier Excel dans le même dossier ou dans ./data/.

Nouveautés de la v6 par rapport à la v5 :
1. Filtre ministériel généralisé : un filtre par ministère / institution, KPI
   affichés ministère par ministère, et vue « global » sur l'ensemble du
   périmètre. Le référentiel des ministères et institutions est intégré pour
   que la liste déroulante reste valable quand la base s'étendra au-delà du
   pilote MINTP / MINAC.
2. État du bien : les biens dont l'état n'est pas renseigné sont désormais
   comptés en « Bon état » (règle métier demandée). La ligne reste tracée par
   l'indicateur « État imputé » pour ne pas perdre l'information de qualité.
3. Cartographie choroplèthe régionale : fond administratif réel (10 régions),
   couleur par région ou par valeur, isolement de la région sélectionnée, et
   légende détaillée qui met la région sélectionnée en évidence.
4. Espace de statistiques dédié au périmètre filtré (module « Statistiques »).

Principes métier :
- valeur du patrimoine = valeur nette comptable (la valeur d'acquisition est
  systématiquement affichée à côté) ;
- Administration centrale / Administration DÉCONCENTRÉE (jamais « décentralisée ») ;
- À réformer ≠ Réformé ;
- immobilisation si prix unitaire >= 500 000 FCFA (seuil sur le prix UNITAIRE) ;
- audit automatique (15 contrôles) + plan d'action priorisé.

Charte visuelle (justifications) :
- une teinte = un sens (Few, 2006) : l'axe vert→or→orange→rouge est réservé
  à la santé/priorité ; administration et classification utilisent un axe
  neutre vert foncé / bleu pétrole / gris ;
- état des biens : une seule palette partout, celle de la carte de points
  (vert = neuf, jaune #F2C744 = bon état, orange = à réformer). Le jaune est
  peu contrasté sur fond blanc (1,6:1) : chaque graphique d'état porte donc
  aussi ses valeurs chiffrées ;
- contrastes >= 3:1 (WCAG 2.1, SC 1.4.11) pour le reste : or #B8860B (3,3:1)
  et gris #6B7280 (4,8:1) remplacent #D4AF37 (2,1:1) et #94A3B8 (2,6:1) ;
- zones cartographiques en camaïeu désaturé, rouge réservé aux alertes
  (traitement préattentif, Healey & Enns 2012) ;
- montants toujours en français (« 1,64 Md », jamais « 1.6G ») et chiffres
  tabulaires (tabular-nums) pour l'alignement en colonnes ;
- encodage redondant couleur + valeur chiffrée (daltonisme, ~8 % des hommes).
"""

import re
import hashlib
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from auth import image_data_uri, logout, require_authentication
from visual_theme import (
    COULEURS_CLASSIF,
    COULEURS_ETAT,
    COULEURS_MINISTERE,
    THEME,
)


from patrimoine_core import (
    COORD_REGIONS,
    GLOBAL_MINISTERES,
    METRIQUES_CARTE,
    NIVEAUX_ADMIN,
    TYPES_GRAPHIQUES,
    NIVEAU_ARRONDISSEMENT,
    NIVEAU_DEPARTEMENT,
    NIVEAU_REGION,
    ORDRE_ETATS,
    ORDRE_MODES_ACQUISITION,
    ORDRE_REGIONS,
    REFERENTIEL_MINISTERES,
    TOUT_LE_CAMEROUN,
    agreger_par_region,
    apparier_au_fond,
    apply_layout,
    charger_geojson_admin,
    carte_choroplethe,
    carte_points,
    charger_et_nettoyer as _charger_et_nettoyer,
    charger_geojson_regions as _charger_geojson_regions,
    comparaison_administrations,
    construire_plan_action,
    couleur_sante,
    detecter_anomalies as _detecter_anomalies,
    detecter_strategies,
    donnees_legende_regions,
    entites_contenues,
    figure_repartition,
    MESURES_CENTRALE,
    TOTAL_MINISTERE,
    treemap_structures_centrales,
    fmt_compact,
    fmt_fcfa,
    fmt_fcfa_complet,
    fmt_int,
    indice_sante,
    indice_sante_par,
    jauge_sante,
    kpi_par_ministere,
    label_sante,
    open_excel,
    palette_entites,
    repartition_etats,
    repartition_etats_carte,
    stats_par,
    statistiques_descriptions,
    statistiques_structures_centrales,
    synthese_kpi,
    texte_sante,
    to_excel_bytes,
)

# =============================================================================
# CACHES STREAMLIT
# =============================================================================
# Les règles de gestion vivent dans `patrimoine_core` ; l'interface se contente
# d'y appliquer sa propre stratégie de cache.

charger_et_nettoyer = st.cache_data(ttl=3600, max_entries=8, show_spinner="Chargement et nettoyage de la base…")(_charger_et_nettoyer)
detecter_anomalies = st.cache_data(ttl=3600, max_entries=8, show_spinner="Audit automatique des anomalies…")(_detecter_anomalies)
charger_geojson_regions = st.cache_data(ttl=3600, max_entries=8, show_spinner=False)(_charger_geojson_regions)


def legende_regions_html(agg, metrique, region_selectionnee, mode_couleur, **options):
    """Rendu HTML de la légende cartographique (les données viennent du cœur).

    `options` transmet le niveau administratif courant : colonne, ordre,
    palette et accord grammatical."""
    titre, sous_titre, lignes = donnees_legende_regions(
        agg, metrique, region_selectionnee, mode_couleur, **options)
    corps = "".join(
        f"<div class='legend-row{' legend-row-active' if l['actif'] else ''}'>"
        f"<span class='legend-chip' style='background:{l['couleur']}'></span>"
        f"<span class='legend-name'>{l['region']}"
        f"{' — sélectionnée' if l['actif'] else ''}</span>"
        f"<span class='legend-val'>{l['valeur']}</span></div>"
        for l in lignes
    )
    return (f"<div class='legend-box'><div class='legend-title'>{titre}</div>"
            f"<div class='legend-sub'>{sous_titre}</div>{corps}</div>")


# =============================================================================
# 6. INTERFACE STREAMLIT
# =============================================================================

st.set_page_config(page_title="MINFI | Cartographie des biens de l'État", page_icon="🏛️", layout="wide", initial_sidebar_state="expanded")

base_dir = Path(__file__).resolve().parent
logo_path = base_dir / "assets" / "logo_minfi_officiel.png"


LOGO_URI = image_data_uri(logo_path)

# Charte : Archivo (display, titres et grandes valeurs) + Inter (texte).
# Ombre portée réservée au premier niveau (cartes KPI) ; les autres blocs
# passent en simple filet — quand tout est mis en valeur, rien ne l'est.
# tabular-nums : chiffres à chasse fixe pour l'alignement des montants.
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Archivo:wght@600;700;800&family=Inter:wght@400;500;600&display=swap');
html, body, [class*="css"] {{ font-family: 'Inter', sans-serif; }}
html, body, #root, .stApp,
[data-testid="stAppViewContainer"],
[data-testid="stMain"],
[data-testid="stMain"] > div {{ background-color: #FFFFFF !important; }}
[data-testid="stHeader"] {{ background-color: rgba(255,255,255,.96) !important; }}
.stApp {{ color: {THEME['texte']}; }}
section[data-testid="stSidebar"] {{ background: #211E16 !important; border-right: 4px solid {THEME['or']}; }}
section[data-testid="stSidebar"] label, section[data-testid="stSidebar"] p, section[data-testid="stSidebar"] span {{ color: {THEME['texte']} !important; }}
section[data-testid="stSidebar"] label, section[data-testid="stSidebar"] p, section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] h1, section[data-testid="stSidebar"] h2, section[data-testid="stSidebar"] h3 {{ color: #F8F3E5 !important; }}
/* Icônes des filtres : la loupe est sur le fond clair du champ, les bulles d'aide sur le fond sombre. */
section[data-testid="stSidebar"] [data-testid="stTextInputIcon"] span {{ color: #7A5A07 !important; }}
section[data-testid="stSidebar"] [data-testid="stTooltipIcon"] svg {{ stroke: #D7B84B !important; }}
/* Zone d'import du classeur : fond clair dans la barre sombre, donc texte foncé.
   Streamlit n'y propose que des libellés anglais : ils sont remplacés ici. */
section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] span,
section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] p {{ color: {THEME['texte']} !important; }}
[data-testid="stFileUploaderDropzone"] button[data-testid="stBaseButton-secondary"] p {{ font-size: 0 !important; }}
[data-testid="stFileUploaderDropzone"] button[data-testid="stBaseButton-secondary"] p::after {{ content: "Parcourir"; font-size: 14px; }}
[data-testid="stFileUploaderDropzoneInstructions"] span {{ font-size: 0 !important; }}
[data-testid="stFileUploaderDropzoneInstructions"] span::after {{ content: "ou glissez un fichier .xlsx ici"; font-size: 12px; }}
.block-container {{ padding-top: .8rem; padding-bottom: 2rem; background-color: #FFFFFF !important; }}
.utility-bar {{ background: #171712; color: #F4EBD3; border-radius: 10px 10px 0 0; padding: 9px 18px; display: flex; justify-content: space-between; gap: 20px; font-size: 11px; font-weight: 600; letter-spacing: .06em; text-transform: uppercase; }}
.utility-bar .gold {{ color: #D7B84B; }}
.main-header {{ background: #FFFFFF; border: 1px solid {THEME['bordure']}; border-top: 0; border-bottom: 4px solid {THEME['or']}; border-radius: 0 0 12px 12px; padding: 18px 28px; margin-bottom: 12px; display: flex; align-items: center; justify-content: space-between; gap: 24px; box-shadow: 0 8px 22px rgba(42, 32, 8, 0.08); }}
.main-header .brand-logo {{ width: 104px; height: 104px; flex: 0 0 104px; object-fit: contain; background: #FFFFFF; padding: 5px; }}
.main-header .header-content {{ flex: 1; text-align: center; }}
.main-header .hero-eyebrow {{ margin: 0 0 6px; font-size: 12px; font-weight: 800; letter-spacing: .18em; text-transform: uppercase; color: {THEME['or']}; }}
.main-header h1 {{ margin: 0; font-size: 32px; font-weight: 800; font-family: 'Archivo', sans-serif; color: #171712; letter-spacing: -0.035em; line-height: 1.15; text-transform: uppercase; }}
.main-header p {{ margin: 7px 0 0; font-size: 13px; color: {THEME['texte2']}; }}
.hero-badges {{ margin-top: 12px; display: flex; justify-content: center; gap: 8px; }}
.badge-mintp, .badge-minac {{ background: #F7F2E3; color: #6F5005; border: 1px solid #DCCB97; }}
@media (max-width: 760px) {{
  .utility-bar {{ flex-direction: column; text-align: center; gap: 3px; }}
  .main-header {{ padding: 18px 14px; gap: 10px; }}
  .main-header .brand-logo {{ width: 62px; height: 62px; flex-basis: 62px; padding: 4px; }}
  .main-header h1 {{ font-size: 21px; }}
  .main-header p {{ font-size: 12px; }}
}}
.kpi-ticker {{ background: #FFFFFF; border: 1px solid {THEME['bordure']}; border-left: 5px solid {THEME['or']}; color: {THEME['texte']}; border-radius: 8px; padding: 0; margin: 10px 0 14px; display: grid; grid-template-columns: repeat(6, 1fr); overflow: hidden; box-shadow: 0 5px 14px rgba(138,101,8,.08); }}
.ticker-item {{ padding: 11px 16px; border-right: 1px solid {THEME['bordure']}; min-width: 0; }}
.ticker-item:last-child {{ border-right: 0; }}
.ticker-label {{ color: #7A5A07; font-size: 10px; font-weight: 800; letter-spacing: .08em; text-transform: uppercase; white-space: nowrap; }}
.ticker-value {{ color: #211E16; font: 800 18px/1.2 'Archivo', sans-serif; margin-top: 3px; font-variant-numeric: tabular-nums; white-space: nowrap; }}
.ticker-value.gold {{ color: #8A6508; }}
.ticker-note {{ color: {THEME['texte2']}; font-size: 10px; margin-top: 2px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
@media (max-width: 900px) {{ .kpi-ticker {{ grid-template-columns: repeat(2, 1fr); }} .ticker-item {{ border-bottom: 1px solid {THEME['bordure']}; }} }}
.section-card {{ background: #FFFFFF; border: 1px solid {THEME['bordure']}; border-radius: 14px; padding: 16px 18px; margin: 10px 0; }}
.section-title {{ font-size: 19px; font-weight: 800; font-family: 'Archivo', sans-serif; margin-bottom: 10px; color: {THEME['texte']}; text-transform: uppercase; letter-spacing: -.015em; border-left: 4px solid {THEME['or']}; padding-left: 10px; }}
.small-muted {{ color: {THEME['texte2']}; font-size: 14px; line-height: 1.6; }}
/* Compatibilité : ces règles ne concernent que les sous-onglets secondaires.
   La navigation principale utilise l'API native segmented_control. */
.stTabs [data-baseweb="tab-list"] {{ gap: 6px; }}
.stTabs [data-baseweb="tab"] {{ background: #FFFFFF; border-radius: 999px; padding: 9px 18px; border: 1px solid {THEME['bordure']}; color: {THEME['texte2']} !important; font-weight: 600; }}
.stTabs [aria-selected="true"] {{ color: white !important; background: linear-gradient(135deg, #6F5005, #A87D0C) !important; border-color: #8A6508 !important; }}
[data-testid="stButton"] button, [data-testid="stFormSubmitButton"] button {{ background: linear-gradient(135deg, #6F5005, #A87D0C); color: white; border: 0; border-radius: 10px; font-weight: 700; min-height: 42px; box-shadow: 0 6px 16px rgba(111,80,5,.18); }}
[data-testid="stButton"] button:hover, [data-testid="stFormSubmitButton"] button:hover {{ background: #5B4104; color: white; border: 0; }}
.login-brand {{ text-align: center; padding: 8px 8px 2px; }}
.login-brand img {{ width: 150px; height: 150px; object-fit: contain; }}
.login-brand h1 {{ font: 800 30px/1.15 'Archivo', sans-serif; color: {THEME['noir']}; margin: 10px 0 6px; }}
.login-brand p {{ color: {THEME['texte2']}; margin: 0 0 12px; }}
.login-kicker {{ color: {THEME['or']}; text-transform: uppercase; letter-spacing: .16em; font-weight: 800; font-size: 11px; }}
.login-intro {{ min-height: 430px; padding: 48px 38px; border-radius: 16px; background: linear-gradient(145deg,#151512,#302A1C); color: white; display: flex; flex-direction: column; justify-content: center; box-shadow: 0 16px 40px rgba(23,23,18,.16); }}
.login-intro img {{ width: 112px; height: 112px; object-fit: contain; background: white; border-radius: 50%; padding: 5px; margin-bottom: 25px; }}
.login-intro .eyebrow {{ color: #D7B84B; font-size: 11px; font-weight: 800; letter-spacing: .16em; text-transform: uppercase; }}
.login-intro h1 {{ color: white; font: 800 32px/1.12 'Archivo', sans-serif; text-transform: uppercase; margin: 10px 0 14px; }}
.login-intro p {{ color: #D8D2C3; line-height: 1.7; font-size: 14px; }}
[data-testid="stMetric"] {{ background: #FFFFFF !important; border: 1px solid {THEME['bordure']} !important; border-radius: 14px !important; padding: 13px 16px !important; }}
[data-testid="stMetricValue"] {{ color: {THEME['texte']} !important; font-size: 23px !important; font-weight: 800 !important; font-family: 'Archivo', sans-serif !important; font-variant-numeric: tabular-nums; }}
[data-testid="stMetricLabel"] {{ color: {THEME['texte2']} !important; font-size: 13px !important; }}
[data-testid="stMetricDelta"] {{ color: {THEME['texte2']} !important; font-size: 13px !important; }}
div[data-testid="stDataFrame"] {{ border: 1px solid {THEME['bordure']}; border-radius: 12px; overflow: hidden; font-variant-numeric: tabular-nums; }}
div[data-testid="stPlotlyChart"] [data-testid="stElementToolbar"] {{ display: none !important; }}
.modebar-container .modebar {{ background: rgba(255,255,255,.92) !important; border: 1px solid {THEME['bordure']} !important; border-radius: 8px !important; padding: 2px 4px !important; }}
.modebar-btn svg {{ fill: #6F5005 !important; }}
.modebar-btn.active svg, .modebar-btn:hover svg {{ fill: #211E16 !important; }}
[data-testid="stVerticalBlockBorderWrapper"] {{ background: #FFFFFF; border-radius: 16px !important; box-shadow: 0 8px 22px rgba(15, 23, 42, 0.05); }}
.badge {{ display: inline-block; padding: 4px 9px; border-radius: 999px; font-size: 12px; font-weight: 700; }}
.badge-high {{ background: #FEE4E2; color: #B42318; }}
.badge-mid {{ background: #FEF0C7; color: #B54708; }}
.badge-low {{ background: #F3EAC9; color: #7A5A07; }}
/* Légende cartographique : rendue en HTML pour afficher la valeur à côté de
   chaque région et surligner la région sélectionnée. */
.legend-box {{ background: #FFFFFF; border: 1px solid {THEME['bordure']}; border-radius: 12px; padding: 12px 14px; }}
.legend-title {{ font: 800 13px/1.3 'Archivo', sans-serif; color: {THEME['texte']}; text-transform: uppercase; letter-spacing: .02em; }}
.legend-sub {{ color: {THEME['texte2']}; font-size: 11px; margin: 2px 0 9px; }}
.legend-row {{ display: flex; align-items: center; gap: 8px; padding: 4px 6px; border-radius: 7px; font-size: 12.5px; }}
.legend-row-active {{ background: {THEME['or_pale']}; font-weight: 700; box-shadow: inset 3px 0 0 {THEME['or']}; }}
.legend-chip {{ width: 13px; height: 13px; border-radius: 3px; flex: 0 0 13px; border: 1px solid rgba(0,0,0,.12); }}
.legend-name {{ flex: 1; color: {THEME['texte']}; }}
.legend-val {{ font-variant-numeric: tabular-nums; color: {THEME['texte2']}; white-space: nowrap; }}
/* Bandeau des filtres actifs, au-dessus des statistiques du périmètre. */
.filter-strip {{ display: flex; flex-wrap: wrap; gap: 7px; align-items: center; margin: 4px 0 12px; }}
.filter-pill {{ background: #FFFFFF; border: 1px solid {THEME['bordure']}; border-left: 3px solid {THEME['or']}; border-radius: 8px; padding: 4px 10px; font-size: 12px; color: {THEME['texte']}; }}
.filter-pill b {{ color: #7A5A07; }}
.filter-none {{ color: {THEME['texte2']}; font-size: 12px; }}
/* Bandeau du périmètre ministériel affiché au-dessus de la navigation. */
.ministry-strip {{ background: linear-gradient(135deg,#171712,#3A3118); color: #F7F0DC; border-radius: 10px; padding: 11px 18px; margin: 2px 0 10px; display: flex; align-items: baseline; justify-content: space-between; gap: 16px; flex-wrap: wrap; }}
.ministry-strip .ms-label {{ font-size: 10px; font-weight: 800; letter-spacing: .14em; text-transform: uppercase; color: #D7B84B; }}
.ministry-strip .ms-name {{ font: 800 19px/1.2 'Archivo', sans-serif; text-transform: uppercase; }}
.ministry-strip .ms-sub {{ font-size: 12px; color: #D8D2C3; }}
.st-key-anti_traduction {{ display: none; }}
</style>
""", unsafe_allow_html=True)

# Streamlit déclare la page en anglais (<html lang="en">) : un navigateur réglé
# en français propose alors de la traduire. La traduction automatique réécrit
# le texte à l'insu de l'interface, qui plante au rafraîchissement suivant
# (« removeChild » : le nœud à supprimer n'est pas un enfant de ce nœud).
# La page est donc déclarée en français et non traduisible.
with st.container(key="anti_traduction"):
    st.html("""
<script>
(() => {
  const racine = document.documentElement;
  racine.lang = "fr";
  racine.setAttribute("translate", "no");
  racine.classList.add("notranslate");
  if (!document.querySelector('meta[name="google"]')) {
    const balise = document.createElement("meta");
    balise.name = "google";
    balise.content = "notranslate";
    document.head.appendChild(balise);
  }
})();
</script>
""", unsafe_allow_javascript=True)


require_authentication(LOGO_URI)

# scrollZoom conservé pour le zoom interactif ; barre d'outils Plotly (export
# PNG, sélection...) masquée à la demande, tout comme le bouton plein écran
# ajouté par Streamlit au survol des graphiques (masqué en CSS ci-dessus).
CONFIG_PLOTLY = {"displayModeBar": False, "scrollZoom": True}

# Les cartes, elles, gardent leur barre d'outils : Plotly n'active le zoom à la
# molette sur les fonds cartographiques qu'en option (son défaut est
# « gl3d+geo+map », sans « mapbox »), et sur une page longue la molette sert
# d'abord à faire défiler. Sans boutons visibles, la carte paraît figée.
# Les boutons de sélection rectangulaire et lasso sont retirés : sur une carte
# ils désactivent le déplacement, ce qui donne l'impression d'un blocage.
CONFIG_CARTE = {
    "displayModeBar": True,
    "displaylogo": False,
    "scrollZoom": True,
    "doubleClick": "reset",
    "modeBarButtonsToRemove": ["select2d", "lasso2d", "select", "lasso"],
    "toImageButtonOptions": {"format": "png", "scale": 2,
                             "filename": "cartographie_biens_etat"},
}
USE_MAPLIBRE = hasattr(go, "Scattermap")


def priority_badge(p):
    cls = "badge-high" if p == "Haute" else "badge-mid" if p == "Moyenne" else "badge-low"
    return f"<span class='badge {cls}'>{p}</span>"

# --- Sélection du classeur Excel ------------------------------------------------
# Deux sources possibles : un classeur importé depuis le poste de l'utilisateur
# (prioritaire, gardé en mémoire le temps de la session, jamais écrit sur le
# disque) ou, à défaut, un classeur rangé dans le dossier de l'application.
fichiers = sorted(base_dir.glob("*.xlsx")) + sorted((base_dir / "data").glob("*.xlsx"))
fichiers = [f for f in fichiers if not f.name.startswith("~$")]

logo_header = f'<img class="brand-logo" src="{LOGO_URI}" alt="Logo du MINFI">' if LOGO_URI else '<div style="font-size:46px">🏛️</div>'
st.markdown(f"""
<div class="utility-bar">
  <span><span class="gold">République du Cameroun</span> · Paix — Travail — Patrie</span>
  <span>Plateforme interne · Accès sécurisé</span>
</div>
<div class="main-header">
  {logo_header}
  <div class="header-content">
    <p class="hero-eyebrow">MINISTÈRE DES FINANCES</p>
    <h1>Cartographie des biens de l'État</h1>
    <p>Filtrage par ministère · cartographie régionale · statistiques · audit · plan d'action</p>
  </div>
  {logo_header}
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    if logo_path.exists():
        st.image(str(logo_path), width=112)
    st.caption(f"Connecté : {st.session_state.get('utilisateur', 'utilisateur')}")
    if st.button("Se déconnecter", width="stretch", icon="🚪"):
        logout()
    st.divider()
    # Emplacement réservé : la recherche et les filtres s'affichent ici, au-dessus
    # de la source de données, mais ne peuvent être dessinés qu'une fois le
    # classeur chargé (leurs listes en dépendent).
    zone_filtres = st.container()

source_obj, nom_source, source_importee = None, None, False
with st.sidebar:
    st.markdown("### 🏛️ Source de données")
    classeur_importe = st.file_uploader(
        "Importer un classeur Excel", type=["xlsx"], key="classeur_importe",
        help="Le classeur importé remplace celui du dossier de l'application pour "
             "cette session. Il reste en mémoire et n'est pas enregistré sur le disque.")
    if classeur_importe is not None:
        source_obj, nom_source, source_importee = (
            classeur_importe.getvalue(), classeur_importe.name, True)
        st.caption(f"Classeur lu : {nom_source} (importé). Retirez-le avec la croix "
                   "pour revenir au classeur du dossier de l'application.")
    elif fichiers:
        nom_source = st.selectbox("Classeur du dossier de l'application",
                                  [f.name for f in fichiers], key="classeur_local")
        source_obj = next(f for f in fichiers if f.name == nom_source).read_bytes()

if source_obj is None:
    st.info("Importez un classeur Excel depuis la barre latérale, ou placez-le dans "
            "le même dossier que `app7.py`.")
    st.stop()

# Changement de classeur : les choix faits sur l'ancien (filtres, zone, ministère
# cartographié…) peuvent ne plus exister dans le nouveau. On repart d'un état
# neutre, en gardant la connexion, la source choisie et le module affiché.
empreinte_source = hashlib.sha256(source_obj).hexdigest()
CLES_CONSERVEES = {"authentifie", "utilisateur", "classeur_importe", "classeur_local",
                   "module_principal", "_empreinte_source"}
if st.session_state.get("_empreinte_source") not in (None, empreinte_source):
    for _cle in list(st.session_state.keys()):
        if _cle not in CLES_CONSERVEES:
            del st.session_state[_cle]
st.session_state["_empreinte_source"] = empreinte_source

RETOUR_SOURCE = (" Retirez le classeur importé dans la barre latérale pour revenir "
                 "à celui du dossier de l'application." if source_importee else "")
try:
    xls0 = open_excel(source_obj)
    noms_feuilles = xls0.sheet_names
except Exception as e:
    st.error(f"Impossible d'ouvrir « {nom_source} » : ce fichier n'est pas un classeur Excel "
             f"(.xlsx) lisible. Détail technique : {str(e).rstrip('.')}.{RETOUR_SOURCE}")
    st.stop()

strategies = detecter_strategies(noms_feuilles)
with st.sidebar:
    choix_strat = st.selectbox("Stratégie de chargement", [s[0] for s in strategies], index=0, key="strategie")
feuilles = dict(strategies)[choix_strat]

try:
    df0, rapport = charger_et_nettoyer(source_obj, tuple(feuilles))
    df0, anomalies0 = detecter_anomalies(df0)
except Exception as exc:
    st.error(f"Chargement ou audit interrompu : {exc}. Aucun résultat partiel ne sera "
             f"présenté.{RETOUR_SOURCE}")
    st.stop()
# Colonnes sans lesquelles les indicateurs n'ont plus de sens (les colonnes
# facultatives, comme l'arrondissement, ne déclenchent pas d'alerte).
COLONNES_ESSENTIELLES = {
    "Ministère": "ministère", "Structure": "structure",
    "_libelle_source": "désignation", "_type_admin_source": "type d'administration",
    "Région": "région", "Ville": "ville", "Quantité": "quantité",
    "Prix unitaire": "prix unitaire", "Valeur d'acquisition": "montant total",
    "Valeur nette comptable": "valeur actuelle", "_etat_source": "état du bien",
}
essentielles_absentes = [libelle for colonne, libelle in COLONNES_ESSENTIELLES.items()
                         if colonne in rapport.get("colonnes_manquantes", [])]
if essentielles_absentes:
    st.warning(f"Le classeur « {nom_source} » ne contient pas toutes les colonnes "
               "attendues : les indicateurs qui en dépendent seront vides ou faux. "
               "Colonnes absentes : " + ", ".join(essentielles_absentes) + "." + RETOUR_SOURCE)
st.caption(f"Lignes entièrement vides écartées : {rapport['n_lignes_vides_ecartees']}")
st.caption(f"Version 7 · Classeur : {nom_source}"
           f"{' (importé)' if source_importee else ''} · Source SHA-256 : {empreinte_source[:16]}")

ministeres_presents = sorted(df0["Ministère"].dropna().unique().tolist())


def libelle_ministere(sigle: str) -> str:
    """Étiquette de la liste déroulante : sigle + libellé complet quand il est
    connu du référentiel, pour qu'un décideur n'ait pas à déchiffrer un sigle."""
    if sigle == GLOBAL_MINISTERES:
        return GLOBAL_MINISTERES
    plein = REFERENTIEL_MINISTERES.get(sigle)
    return f"{sigle} — {plein}" if plein else sigle


# --- Recherche et filtres, dans la barre latérale -------------------------------
# Ils sont toujours visibles, au même endroit quel que soit le module consulté.
# Le bandeau de pilules affiché en tête de chaque module rappelle les filtres
# actifs, pour qu'aucun chiffre restreint ne soit lu sans son périmètre.
VALEURS_FILTRES_DEFAUT = {
    "ministere_sel": GLOBAL_MINISTERES, "admin_sel": "Toutes", "structure_sel": [],
    "classif_sel": "Toutes", "reg_sel": [], "etat_sel": [], "nature_sel": [],
    "cat_sel": [], "libelle_sel": [], "mode_sel": [], "recherche_global": "",
}
for _cle, _valeur in VALEURS_FILTRES_DEFAUT.items():
    st.session_state.setdefault(_cle, _valeur)


def reinitialiser_filtres():
    """Remet la recherche et tous les filtres à leur valeur neutre."""
    for cle, valeur in VALEURS_FILTRES_DEFAUT.items():
        st.session_state[cle] = valeur


def compter_filtres_actifs():
    e = st.session_state
    return sum([
        e["ministere_sel"] != GLOBAL_MINISTERES, e["admin_sel"] != "Toutes",
        bool(e["structure_sel"]),
        e["classif_sel"] != "Toutes", bool(e["reg_sel"]), bool(e["etat_sel"]),
        bool(e["nature_sel"]), bool(e["cat_sel"]), bool(e["libelle_sel"]),
        bool(e["mode_sel"]), bool((e["recherche_global"] or "").strip()),
    ])


with zone_filtres:
    st.markdown("### 🔍 Recherche")
    recherche_global = st.text_input(
        "Recherche libre", key="recherche_global", icon=":material/search:",
        placeholder="ordinateur, Yaoundé, MINTP…", label_visibility="collapsed")
    st.caption("Libellés, identifiants, structures, villes, régions, ministères… "
               "Validez avec Entrée.")

    st.markdown("### Filtres")
    actifs = compter_filtres_actifs()
    st.caption(f"{actifs} critère(s) actif(s), recherche comprise." if actifs
               else "Aucun critère actif : toute la base est affichée.")
    st.button("Réinitialiser", width="stretch", icon="↩️",
              on_click=reinitialiser_filtres, disabled=not actifs,
              key="reset_filtres")
    ministere_sel = st.selectbox(
        "Ministère / institution", [GLOBAL_MINISTERES] + ministeres_presents,
        format_func=libelle_ministere, key="ministere_sel",
        help="« Global » consolide l'ensemble du périmètre chargé.")
    admin_sel = st.selectbox(
        "Administration",
        ["Toutes"] + [a for a in ["Centrale", "Déconcentrée", "Non renseigné"]
                      if a in df0["Administration"].unique()],
        key="admin_sel")
    structure_sel = st.multiselect(
        "Structure", sorted(df0["Structure agrégée"].dropna().unique()),
        placeholder="Toutes", key="structure_sel",
        help="Direction, service ou délégation, tel que saisi dans le classeur. "
             "Un même sigle peut exister dans plusieurs ministères : combinez "
             "avec le filtre « Ministère / institution ».")
    classif_sel = st.selectbox(
        "Classification",
        ["Toutes"] + [c for c in ["Immobilisation", "Stock", "Indéterminée"]
                      if c in df0["Classification"].unique()],
        key="classif_sel")
    reg_sel = st.multiselect("Région(s)", sorted(df0["Région"].dropna().unique()),
                             placeholder="Toutes", key="reg_sel")
    etat_sel = st.multiselect(
        "État du bien",
        [e for e in ORDRE_ETATS if e in df0["État du bien"].unique()],
        placeholder="Tous", key="etat_sel")
    nature_sel = st.multiselect(
        "Nature du matériel", sorted(df0["Nature du matériel"].dropna().unique()),
        placeholder="Toutes", key="nature_sel")
    cat_sel = st.multiselect("Catégorie", sorted(df0["Catégorie"].dropna().unique()),
                             placeholder="Toutes", key="cat_sel")
    libelle_sel = st.multiselect(
        "Objet précis", sorted(df0["Libellé standardisé"].dropna().unique()),
        placeholder="Tous les objets", key="libelle_sel",
        help="Restreint toute l'application à un ou plusieurs objets "
             "(chaise, véhicule, ordinateur…).")
    modes = [m for m in ORDRE_MODES_ACQUISITION
             if m in df0["Mode d'acquisition"].unique()]
    autres = sorted([m for m in df0["Mode d'acquisition"].dropna().unique()
                     if m not in modes])
    mode_sel = st.multiselect("Mode d'acquisition", modes + autres,
                              placeholder="Tous", key="mode_sel")
    st.divider()

with st.sidebar:
    with st.expander(f"Référentiel des ministères ({len(REFERENTIEL_MINISTERES)})"):
        couverture = pd.DataFrame({
            "Sigle": list(REFERENTIEL_MINISTERES.keys()),
            "Administration": list(REFERENTIEL_MINISTERES.values()),
        })
        couverture["Recensé"] = np.where(
            couverture["Sigle"].isin(ministeres_presents), "✅ Oui", "— Pas encore")
        st.dataframe(couverture, width="stretch", height=240, hide_index=True)
        if rapport.get("ministeres_hors_referentiel"):
            st.caption("Hors référentiel dans le classeur : "
                       + ", ".join(rapport["ministeres_hors_referentiel"]))

df = df0.copy()
min_sel = [] if ministere_sel == GLOBAL_MINISTERES else [ministere_sel]
if min_sel:
    df = df[df["Ministère"].isin(min_sel)]
if nature_sel:
    df = df[df["Nature du matériel"].isin(nature_sel)]
if libelle_sel:
    df = df[df["Libellé standardisé"].isin(libelle_sel)]
if admin_sel != "Toutes":
    df = df[df["Administration"] == admin_sel]
if structure_sel:
    df = df[df["Structure agrégée"].isin(structure_sel)]
if classif_sel != "Toutes":
    df = df[df["Classification"] == classif_sel]
if reg_sel:
    df = df[df["Région"].isin(reg_sel)]
if etat_sel:
    df = df[df["État du bien"].isin(etat_sel)]
if cat_sel:
    df = df[df["Catégorie"].isin(cat_sel)]
if mode_sel:
    df = df[df["Mode d'acquisition"].isin(mode_sel)]
if recherche_global.strip():
    pat = re.escape(recherche_global.strip())
    m = pd.Series(False, index=df.index)
    for c in ["Libellé standardisé", "Libellé source", "Description", "Ministère",
              "Ministère (libellé)", "Mode d'acquisition", "Mode d'acquisition source",
              "Nature du matériel", "Catégorie", "Identifiant", "Localisation",
              "Structure agrégée", "Ville", "Département", "Région", "État du bien"]:
        if c in df.columns:
            m |= df[c].astype(str).str.contains(pat, case=False, na=False)
    df = df[m]

if df.empty:
    st.warning("Aucun bien ne correspond à la recherche et aux filtres sélectionnés. "
               "Modifiez-les ou réinitialisez-les dans la barre latérale.")
    st.stop()

anomalies = anomalies0[anomalies0["_index_ligne"].isin(df.index)].copy()
plan = construire_plan_action(anomalies)
kpi = synthese_kpi(df, anomalies)
ministeres_perimetre = sorted(df["Ministère"].dropna().unique().tolist()) or ["n.d."]

with st.expander("Fiabilité des données du périmètre — conventions et couverture", expanded=True):
    q_total = float(df["Quantité"].sum())
    q_imputee = float(df.loc[df["État imputé"], "Quantité"].sum())
    observe = indice_sante(df.loc[~df["État imputé"]])
    score_obs = "Non calculable" if observe is None else f"{observe:.1f}/100"
    st.warning(f"État imputé à Bon état : {fmt_int(q_imputee)} biens sur {fmt_int(q_total)} "
               f"({100*q_imputee/q_total if q_total else 0:.1f} %). "
               "Les graphiques d’état et l’indice conventionnel incluent cette hypothèse.")
    st.write(f"Indice sur états renseignés uniquement : **{score_obs}**. "
             "Ce sous-ensemble n’est pas nécessairement représentatif du parc.")
    n_gps = int(df["Source géolocalisation"].eq("Coordonnées MAPS").sum())
    st.caption(f"Coordonnées MAPS renseignées : {n_gps} / {len(df)} lignes (non vérifiées sur le terrain). "
               "Les autres positions sont des repères de ville ou de chef-lieu, pas l’emplacement des biens.")
    st.caption(f"Quantités corrigées : {int(df['_qte_imputee'].sum())} lignes. "
               f"VNC reconstituées : {int(df['VNC imputée'].sum())} lignes. "
               f"VNC manquantes : {int(df['Valeur nette comptable'].isna().sum())} lignes. "
               "Les sommes monétaires ignorent les montants manquants.")


# Inventaire des filtres actifs : il alimente le bandeau affiché au-dessus des
# statistiques, pour qu'aucun chiffre ne soit lu sans son périmètre.
filtres_actifs = []
if ministere_sel != GLOBAL_MINISTERES:
    filtres_actifs.append(("Ministère", ministere_sel))
if admin_sel != "Toutes":
    filtres_actifs.append(("Administration", admin_sel))
if classif_sel != "Toutes":
    filtres_actifs.append(("Classification", classif_sel))
for etiquette, valeurs in [("Structure", structure_sel),
                           ("Région", reg_sel), ("État", etat_sel),
                           ("Nature du matériel", nature_sel), ("Catégorie", cat_sel),
                           ("Bien", libelle_sel), ("Mode d'acquisition", mode_sel)]:
    if valeurs:
        filtres_actifs.append((etiquette, ", ".join(map(str, valeurs))))
if recherche_global.strip():
    filtres_actifs.append(("Recherche", recherche_global.strip()))

if filtres_actifs:
    pilules = "".join(f"<span class='filter-pill'><b>{k}</b> · {v}</span>"
                      for k, v in filtres_actifs)
    bandeau_filtres = f"<div class='filter-strip'>{pilules}</div>"
else:
    bandeau_filtres = ("<div class='filter-strip'><span class='filter-none'>"
                       "Aucun filtre actif — lecture sur l'ensemble du périmètre chargé."
                       "</span></div>")

# Titre du périmètre ministériel, affiché au-dessus de la navigation : le
# lecteur sait toujours de quelle administration proviennent les chiffres.
if ministere_sel == GLOBAL_MINISTERES:
    titre_perimetre = "Ensemble du périmètre recensé"
    sous_titre_perimetre = (f"{len(ministeres_perimetre)} administration(s) : "
                            f"{', '.join(ministeres_perimetre)}")
else:
    titre_perimetre = ministere_sel
    sous_titre_perimetre = REFERENTIEL_MINISTERES.get(
        ministere_sel, "Administration non référencée dans le référentiel intégré")

st.markdown(f"""
<div class="ministry-strip">
  <div>
    <div class="ms-label">Périmètre analysé</div>
    <div class="ms-name">{titre_perimetre}</div>
    <div class="ms-sub">{sous_titre_perimetre}</div>
  </div>
  <div class="ms-sub">
    {fmt_int(kpi['n_biens'])} biens · {fmt_fcfa(kpi['vnc_totale'])} de VNC ·
    {fmt_fcfa(kpi['va_totale'])} de valeur d'acquisition
  </div>
</div>
""", unsafe_allow_html=True)

# Ruban synthétique affiché au-dessus de la navigation principale.
pct_immo_biens = f"{kpi['part_immo_biens']*100:.1f}".replace(".", ",")
pct_immo_vnc = f"{kpi['part_immo_vnc']*100:.1f}".replace(".", ",")
pct_stock_biens = f"{kpi['part_stock_biens']*100:.1f}".replace(".", ",")
pct_stock_vnc = f"{kpi['part_stock_vnc']*100:.1f}".replace(".", ",")
pct_a_reformer = f"{kpi['part_a_reformer']*100:.1f}".replace(".", ",")

st.markdown(f"""
<div class="kpi-ticker">
  <div class="ticker-item">
    <div class="ticker-label">Valeur nette comptable</div>
    <div class="ticker-value gold">{fmt_fcfa(kpi['vnc_totale'])}</div>
    <div class="ticker-note">Valeur de référence</div>
  </div>
  <div class="ticker-item">
    <div class="ticker-label">Valeur d'acquisition</div>
    <div class="ticker-value">{fmt_fcfa(kpi['va_totale'])}</div>
    <div class="ticker-note">Amorti à {f"{kpi['taux_amortissement']*100:.1f}".replace(".", ",")} %</div>
  </div>
  <div class="ticker-item">
    <div class="ticker-label">Biens recensés</div>
    <div class="ticker-value">{fmt_int(kpi['n_biens'])}</div>
    <div class="ticker-note">{fmt_int(kpi['n_lignes'])} lignes d'inventaire</div>
  </div>
  <div class="ticker-item">
    <div class="ticker-label">Immobilisations</div>
    <div class="ticker-value">{pct_immo_biens} %</div>
    <div class="ticker-note">{fmt_fcfa(kpi['vnc_immo'])} · {pct_immo_vnc} % VNC</div>
  </div>
  <div class="ticker-item">
    <div class="ticker-label">Stocks</div>
    <div class="ticker-value">{pct_stock_biens} %</div>
    <div class="ticker-note">{fmt_fcfa(kpi['vnc_stock'])} · {pct_stock_vnc} % VNC</div>
  </div>
  <div class="ticker-item">
    <div class="ticker-label">Biens à réformer</div>
    <div class="ticker-value">{fmt_int(kpi['n_a_reformer'])}</div>
    <div class="ticker-note">{pct_a_reformer} % du parc</div>
  </div>
</div>
""", unsafe_allow_html=True)

# =============================================================================
# 7. ONGLETS — 4 MODULES
# =============================================================================

module = st.segmented_control(
    "Navigation principale",
    ["Vue exécutive", "Cartographie", "Statistiques", "Analyse du patrimoine",
     "Données et exports"],
    default="Vue exécutive",
    key="module_principal",
    selection_mode="single",
    label_visibility="collapsed",
    width="stretch",
)

# -----------------------------------------------------------------------------
# Onglet 1 : Vue exécutive
# -----------------------------------------------------------------------------
if module == "Vue exécutive":
    st.markdown(bandeau_filtres, unsafe_allow_html=True)
    sante = kpi["indice_sante"]

    # Élément signature : jauge de santé + répartitions VNC / état, chacun dans sa propre box.
    j1, j2, j3 = st.columns(3)
    with j1, st.container(border=True):
        st.markdown("<div class='section-title'>Indice de santé patrimoniale</div>", unsafe_allow_html=True)
        if sante is not None:
            st.plotly_chart(jauge_sante(sante), width='stretch',
                            config={"displayModeBar": False})
            st.markdown(f"<div class='small-muted' style='text-align:center;margin-top:-14px'>"
                        f"<b style='color:{couleur_sante(sante)}'>{label_sante(sante)}</b> · "
                        "Barème : Neuf 100 · Bon état 70 · À réformer 20 · Réformé 0, "
                        "pondéré par les quantités</div>", unsafe_allow_html=True)
        else:
            st.info("Indice non calculable : aucun état renseigné dans le périmètre filtré.")
    with j2, st.container(border=True):
        vnc_tot_aff = kpi["vnc_totale"] or 1.0
        rep_type = df.groupby("Classification", dropna=False).agg(
            Biens=("Quantité", "sum"),
            VNC=("Valeur nette comptable", "sum"),
        ).reset_index().rename(columns={"Classification": "Type de bien"})
        rep_type["Étiquette %"] = (rep_type["VNC"] / vnc_tot_aff * 100).map(lambda x: f"{x:.1f}".replace(".", ",") + " %")
        fig = px.pie(rep_type, names="Type de bien", values="VNC", hole=0.58,
                     color="Type de bien", color_discrete_map=COULEURS_CLASSIF)
        fig.update_traces(text=rep_type["Étiquette %"], textinfo="text", textfont_size=12,
                           customdata=rep_type["VNC"].map(fmt_fcfa_complet),
                           hovertemplate="%{label}<br>%{customdata} · %{text}<extra></extra>")
        fig.add_annotation(text=f"<b>{fmt_compact(vnc_tot_aff)}</b><br><span style='font-size:11px'>VNC totale</span>",
                           showarrow=False, font=dict(size=18, family="Archivo", color=THEME["texte"]))
        st.plotly_chart(apply_layout(fig, 340, "Répartition de la VNC par type de bien"), width='stretch', config=CONFIG_PLOTLY)
    with j3, st.container(border=True):
        d_etat = repartition_etats(df).reset_index()
        tot_biens_etat = d_etat["Nombre de biens"].sum() or 1
        d_etat["Étiquette %"] = (d_etat["Nombre de biens"] / tot_biens_etat * 100).map(lambda x: f"{x:.1f}".replace(".", ",") + " %")
        fig = px.pie(d_etat, names="État du bien", values="Nombre de biens", hole=0.58, color="État du bien", color_discrete_map=COULEURS_ETAT)
        fig.update_traces(text=d_etat["Étiquette %"], textinfo="text", textfont_size=12)
        centre = texte_sante(sante).split(" ")[0] if sante is not None else fmt_int(kpi["n_biens"])
        sous = "santé" if sante is not None else "biens"
        fig.add_annotation(text=f"<b>{centre}</b><br><span style='font-size:11px'>{sous}</span>",
                           showarrow=False, font=dict(size=20, family="Archivo",
                                                      color=couleur_sante(sante) if sante is not None else THEME["texte"]))
        st.plotly_chart(apply_layout(fig, 340, "État du patrimoine"), width='stretch', config=CONFIG_PLOTLY)

    sante_reg = indice_sante_par(df, "Région").dropna().sort_values()
    zones_prioritaires = ", ".join([f"{r} ({v:.0f}/100)" for r, v in sante_reg.head(3).items()]) or "aucune zone prioritaire isolée"
    if not anomalies.empty:
        top_anos = anomalies.groupby("Type d'anomalie")["VNC exposée (FCFA)"].sum().sort_values(ascending=False).head(3)
        anos_txt = " ; ".join([f"{a} ({fmt_fcfa(v)})" for a, v in top_anos.items()])
    else:
        anos_txt = "aucune anomalie détectée sur le périmètre filtré"

    st.markdown(f"""
    <div class="section-card">
      <div class="section-title">Résumé exécutif</div>
      <div class="small-muted">
        Le patrimoine inventorié ({', '.join(ministeres_perimetre)}) présente une valeur nette comptable de
        <b>{fmt_fcfa(kpi['vnc_totale'])}</b>, pour une valeur d'acquisition de <b>{fmt_fcfa(kpi['va_totale'])}</b>
        et un amortissement cumulé de <b>{fmt_fcfa(kpi['amort_total'])}</b>. Les immobilisations représentent
        <b>{f"{kpi['part_immo_vnc']*100:.1f}".replace(".", ",")} %</b> de la valeur nette, contre
        <b>{f"{kpi['part_stock_vnc']*100:.1f}".replace(".", ",")} %</b> pour les stocks.
        L'indice de santé patrimoniale est de <b>{texte_sante(sante)}</b>.
        Les zones à regarder en priorité sont : <b>{zones_prioritaires}</b>.
        Les principales anomalies concernent : {anos_txt}.
      </div>
    </div>
    """, unsafe_allow_html=True)

    # Encart Central vs Déconcentré (question n°4 du cahier des charges).
    st.markdown("<div class='section-title'>Administration centrale vs déconcentrée</div>", unsafe_allow_html=True)
    st.markdown("<div class='small-muted'>Les services régionaux et départementaux de l'État constituent "
                "l'administration <b>déconcentrée</b> (et non « décentralisée », terme réservé aux "
                "collectivités territoriales, absentes de cette base).</div>", unsafe_allow_html=True)
    comp_admin = comparaison_administrations(df)
    st.dataframe(comp_admin.style.format({
        "Biens": "{:,.0f}", "Valeur d'acquisition (FCFA)": "{:,.0f}",
        "VNC (FCFA)": "{:,.0f}", "Part VNC (%)": "{:.1f}",
        "À réformer (%)": "{:.1f}", "Indice santé": "{:.1f}", "Anomalies": "{:,.0f}"},
        thousands=" ", decimal=",", na_rep="n.d."), width='stretch')

    # --- KPI ministère par ministère + ligne « Ensemble » ---------------------
    st.markdown("<div class='section-title'>Indicateurs par ministère</div>",
                unsafe_allow_html=True)
    st.markdown("<div class='small-muted'>Choisissez le ministère dont vous voulez lire les "
                "indicateurs ; par défaut ils portent sur l'ensemble des ministères de la "
                "base. Le tableau qui suit reste complet, ministère par ministère. Le filtre "
                "« Ministère / institution » de la barre latérale recalcule en plus la totalité "
                "de l'application sur un seul ministère.</div>", unsafe_allow_html=True)

    # Calculé sur le périmètre filtré courant : le tableau doit rester cohérent
    # avec le ruban de KPI affiché juste au-dessus (mêmes filtres, mêmes totaux).
    tableau_min = kpi_par_ministere(df, anomalies)
    lignes_min = {r["Ministère"]: r for r in tableau_min.to_dict("records")}

    # Une seule boîte de KPI, pilotée par une liste déroulante : avec une base
    # étendue à l'ensemble de l'État (plusieurs dizaines d'administrations), une
    # boîte par ministère produirait un mur d'encarts illisible.
    OPTION_ENSEMBLE = "Tous les ministères (ensemble du périmètre)"
    ministeres_dispo = [m for m in lignes_min if m != "Ensemble du périmètre"]
    options_kpi = [OPTION_ENSEMBLE] + ministeres_dispo
    kpi_ministere_sel = st.selectbox(
        "Ministère dont afficher les indicateurs", options_kpi, index=0,
        key="kpi_ministere",
        format_func=lambda m: m if m == OPTION_ENSEMBLE else libelle_ministere(m),
        help="Les indicateurs sont recalculés sur le seul ministère choisi. "
             "« Tous les ministères » consolide le périmètre filtré courant.",
    )

    if kpi_ministere_sel == OPTION_ENSEMBLE:
        # kpi_par_ministere n'ajoute la ligne consolidée que s'il y a plusieurs
        # ministères : avec un seul, sa propre ligne EST l'ensemble.
        ligne_kpi = lignes_min.get("Ensemble du périmètre") or next(iter(lignes_min.values()))
        nom_affiche = OPTION_ENSEMBLE
        libelle_affiche = (f"{len(ministeres_dispo)} administration(s) : "
                           f"{', '.join(ministeres_dispo)}")
    else:
        ligne_kpi = lignes_min[kpi_ministere_sel]
        nom_affiche = kpi_ministere_sel
        libelle_affiche = ligne_kpi["Libellé"]

    with st.container(border=True):
        indice_kpi = ligne_kpi["Indice santé"]
        st.markdown(
            f"<div style='font:800 19px/1.2 Archivo,sans-serif;text-transform:uppercase'>"
            f"{nom_affiche}</div>"
            f"<div class='small-muted' style='font-size:12.5px;margin-bottom:10px'>"
            f"{libelle_affiche}</div>", unsafe_allow_html=True)
        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Biens", fmt_int(ligne_kpi["Biens"]),
                  delta=f"{fmt_int(ligne_kpi['Lignes'])} lignes", delta_color="off")
        k2.metric("Valeur d'acquisition", fmt_fcfa(ligne_kpi["Valeur d'acquisition (FCFA)"]))
        k3.metric("VNC", fmt_fcfa(ligne_kpi["VNC (FCFA)"]))
        k4.metric("Amortissement cumulé", fmt_fcfa(ligne_kpi["Amortissement (FCFA)"]))
        k5.metric("Indice de santé",
                  "n.d." if pd.isna(indice_kpi) else f"{indice_kpi:.1f}".replace(".", ","),
                  delta=label_sante(indice_kpi), delta_color="off")
        e1, e2, e3, e4, e5 = st.columns(5)
        e1.metric("Neufs", f"{str(ligne_kpi['Neufs (%)']).replace('.', ',')} %")
        e2.metric("Bon état", f"{str(ligne_kpi['Bon état (%)']).replace('.', ',')} %")
        e3.metric("À réformer", f"{str(ligne_kpi['À réformer (%)']).replace('.', ',')} %")
        e4.metric("Structures", fmt_int(ligne_kpi["Structures"]))
        e5.metric("Régions couvertes", fmt_int(ligne_kpi["Régions couvertes"]),
                  delta=f"{fmt_int(ligne_kpi['Anomalies'])} anomalies", delta_color="off")

    st.dataframe(
        tableau_min.set_index("Ministère").style.format({
            "Biens": "{:,.0f}", "Lignes": "{:,.0f}", "VNC (FCFA)": "{:,.0f}",
            "Valeur d'acquisition (FCFA)": "{:,.0f}", "Amortissement (FCFA)": "{:,.0f}",
            "Neufs (%)": "{:.1f}", "Bon état (%)": "{:.1f}", "À réformer (%)": "{:.1f}",
            "Indice santé": "{:.1f}", "Structures": "{:,.0f}",
            "Régions couvertes": "{:,.0f}", "Anomalies": "{:,.0f}"},
            thousands=" ", decimal=",", na_rep="n.d."), width='stretch')

    if len(ministeres_dispo) > 1:
        comparatif = pd.DataFrame([lignes_min[m] for m in ministeres_dispo])
        g1, g2 = st.columns(2)
        with g1:
            fondu = comparatif.melt(id_vars="Ministère",
                                    value_vars=["Valeur d'acquisition (FCFA)", "VNC (FCFA)"],
                                    var_name="Valeur", value_name="Montant")
            fondu["Étiquette"] = fondu["Montant"].map(fmt_compact)
            fig = px.bar(fondu, x="Ministère", y="Montant", color="Valeur", barmode="group",
                         text="Étiquette",
                         color_discrete_map={"Valeur d'acquisition (FCFA)": THEME["bleu"],
                                             "VNC (FCFA)": THEME["vert"]})
            fig.update_traces(textposition="outside")
            st.plotly_chart(apply_layout(fig, 380, "Valeur d'acquisition et VNC par ministère"),
                            width='stretch', config=CONFIG_PLOTLY)
        with g2:
            etat_min = df.groupby(["Ministère", "État du bien"], as_index=False).agg(
                Biens=("Quantité", "sum"))
            total_min = etat_min.groupby("Ministère")["Biens"].transform("sum")
            etat_min["Part"] = 100 * etat_min["Biens"] / total_min
            etat_min["Étiquette"] = etat_min["Part"].map(lambda x: f"{x:.1f}".replace(".", ",") + " %")
            fig = px.bar(etat_min, x="Ministère", y="Part", color="État du bien",
                         barmode="stack", text="Étiquette", color_discrete_map=COULEURS_ETAT,
                         category_orders={"État du bien": ORDRE_ETATS})
            fig.update_layout(yaxis_title="Part des biens (%)")
            st.plotly_chart(apply_layout(fig, 380, "État du parc par ministère (%)"),
                            width='stretch', config=CONFIG_PLOTLY)

# -----------------------------------------------------------------------------
# Onglet 2 : Cartographie (choroplèthe régionale + carte détaillée)
# -----------------------------------------------------------------------------
TOUS_MINISTERES_CARTE = "Tous les ministères du périmètre"
TOUS_BIENS_CARTE = "Tous"

@st.cache_data(ttl=3600, max_entries=8, show_spinner=False)
def charger_fond(niveau: str):
    """Fond administratif d'un niveau, mis en cache (les arrondissements
    pèsent 1,4 Mo : on évite de les relire à chaque interaction)."""
    cfg = NIVEAUX_ADMIN[niveau]
    return charger_geojson_admin(str(base_dir / "assets" / cfg["fichier"]),
                                 cfg["traduction"])


if module == "Cartographie":
    st.markdown(bandeau_filtres, unsafe_allow_html=True)

    vue_carte = st.segmented_control(
        "Type de carte",
        ["Carte choroplèthe", "Carte détaillée (points)"],
        default="Carte choroplèthe",
        key="vue_carte",
        selection_mode="single",
        label_visibility="collapsed",
        width="stretch",
    )

    # -------------------------------------------------------------------------
    # 2.a Choroplèthe régionale
    # -------------------------------------------------------------------------
    if vue_carte == "Carte choroplèthe":
        # Niveau de découpage : régions, départements ou arrondissements. Seuls
        # les niveaux que la base peut réellement chiffrer sont proposés à la
        # coloration — un niveau sans données ne produirait qu'une carte vide.
        niveaux_possibles = []
        for niveau, cfg in NIVEAUX_ADMIN.items():
            colonne = cfg["colonne"]
            renseigne = (colonne in df.columns
                         and df[colonne].dropna().replace("Non renseigné", pd.NA).notna().any())
            if renseigne:
                niveaux_possibles.append(niveau)
        niveau_admin = st.radio(
            "Niveau de découpage", niveaux_possibles, index=0, horizontal=True,
            key="niveau_admin",
            help="La coloration porte sur le niveau choisi. Les limites du "
                 "niveau inférieur peuvent être superposées en trait fin.")
        cfg_niveau = NIVEAUX_ADMIN[niveau_admin]
        colonne_admin = cfg_niveau["colonne"]
        geojson_admin, entites_fond, centres_fond = charger_fond(niveau_admin)

        niveaux_absents = [n for n in NIVEAUX_ADMIN if n not in niveaux_possibles]
        if niveaux_absents:
            st.caption("Niveau(x) non cartographiable(s) faute de données dans le "
                       f"classeur : {', '.join(niveaux_absents).lower()}. "
                       "Les limites restent affichables en surimpression.")

        if geojson_admin is None:
            st.error(
                "Fond administratif introuvable. Placez le fichier "
                f"`assets/{cfg_niveau['fichier']}` à côté de l'application.")
        else:
            c1, c2, c3, c4 = st.columns([1.3, 1.2, 1.3, 1.1])
            with c1:
                # Filtre ministériel propre à la carte : il se combine avec le
                # filtre global de la barre latérale (la liste ne propose que
                # les ministères du périmètre courant, jamais une intersection
                # vide). Il pilote la carte, ses statistiques et son tableau.
                ministeres_carte = sorted(df["Ministère"].dropna().unique().tolist())
                ministere_carte = st.selectbox(
                    "Ministère cartographié",
                    [TOUS_MINISTERES_CARTE] + ministeres_carte, index=0,
                    key="ministere_carte",
                    format_func=lambda m: (m if m == TOUS_MINISTERES_CARTE
                                           else libelle_ministere(m)),
                    help="Restreint la carte, ses statistiques et son tableau à "
                         "un seul ministère, sans toucher au reste de l'application.")
            with c2:
                entite_carte = st.selectbox(
                    f"{cfg_niveau['singulier'].capitalize()} à afficher",
                    [TOUT_LE_CAMEROUN] + entites_fond, index=0,
                    key=f"entite_carte_{niveau_admin}",
                    help="Une entité sélectionnée est affichée seule ; le reste du "
                         "pays est retiré de la carte.")
            with c3:
                metrique_carte = st.selectbox("Indicateur cartographié",
                                              list(METRIQUES_CARTE.keys()), index=0,
                                              key="metrique_carte")
            with c4:
                mode_couleur = st.selectbox(
                    "Coloration",
                    [f"Une couleur par {cfg_niveau['singulier']}", "Par valeur (dégradé)"],
                    index=0, key=f"mode_couleur_{niveau_admin}",
                    help="La coloration par entité distingue les territoires ; "
                         "« par valeur » les classe du plus faible au plus fort.")

            o1, o2 = st.columns([1.1, 1])
            with o1:
                isoler_centrale = st.toggle(
                    "Isoler l'administration centrale de la région du Centre",
                    value=True, key="isoler_centrale",
                    help="Les services centraux sont tous implantés à Yaoundé : les "
                         "laisser dans la région du Centre gonfle celle-ci et fausse "
                         "la comparaison interrégionale. Ils apparaissent alors sur "
                         "une ligne propre du tableau, et restent comptés dans les "
                         "totaux du périmètre.")
            with o2:
                # Le niveau inférieur en surimpression : c'est le seul moyen de
                # voir les arrondissements, que la base ne renseigne pas encore.
                niveaux_inferieurs = {NIVEAU_REGION: NIVEAU_DEPARTEMENT,
                                      NIVEAU_DEPARTEMENT: NIVEAU_ARRONDISSEMENT}
                niveau_sous = niveaux_inferieurs.get(niveau_admin)
                afficher_sous = False
                if niveau_sous:
                    afficher_sous = st.toggle(
                        f"Afficher les limites des {niveau_sous.lower()}",
                        value=False, key=f"sous_decoupage_{niveau_admin}",
                        help="Trace le découpage du niveau inférieur en trait fin, "
                             "sous les aplats de couleur.")
                afficher_noms_carte = st.toggle(
                    "Afficher les noms sur la carte", value=True,
                    key=f"noms_carte_{niveau_admin}",
                    help="Pose le nom de chaque entité en son centre. Les noms du "
                         "niveau inférieur ne sont écrits que sur une entité "
                         "sélectionnée, où ils restent lisibles.")

            # Périmètre de la carte : filtre ministériel local, puis mise à
            # l'écart éventuelle des services centraux.
            df_carte = df if ministere_carte == TOUS_MINISTERES_CARTE else \
                df[df["Ministère"] == ministere_carte]
            df_choro = (df_carte[df_carte["Administration"] != "Centrale"]
                        if isoler_centrale else df_carte)

            # Rapprochement des libellés de la base et des entités du fond :
            # les orthographes divergent (BAMBOUTOUS / Bamboutos, KADEY / Kadei).
            valeurs_base = (df_choro[colonne_admin].dropna().unique()
                            if colonne_admin in df_choro.columns else [])
            correspondances, non_apparies = apparier_au_fond(
                [v for v in valeurs_base if v != "Non renseigné"],
                entites_fond, cfg_niveau["corrections"])
            df_zone = df_choro.copy()
            df_zone["_entite"] = (df_zone[colonne_admin].map(correspondances)
                                  if colonne_admin in df_zone.columns else pd.NA)
            if non_apparies:
                st.warning("Non rattaché(s) au fond administratif, donc absent(s) de la "
                           f"carte : {', '.join(non_apparies[:12])}"
                           + (" …" if len(non_apparies) > 12 else ""))

            def completer_entites(agg, colonne):
                """Une entité sans bien doit s'afficher à zéro : la faire
                disparaître la rendrait indiscernable d'une absence de donnée."""
                manquantes = [r for r in entites_fond if r not in set(agg[colonne])]
                if not manquantes:
                    return agg
                vide = pd.DataFrame({colonne: manquantes})
                for col in agg.columns:
                    if col != colonne:
                        vide[col] = 0
                return pd.concat([agg, vide], ignore_index=True)

            agg_regions = agreger_par_region(
                df_zone[df_zone["_entite"].notna()], "_entite", nom_colonne=colonne_admin)
            agg_regions = completer_entites(agg_regions, colonne_admin)
            reference_ordre = (ORDRE_REGIONS if niveau_admin == NIVEAU_REGION
                               else entites_fond)
            ordre_entites = [e for e in reference_ordre
                             if e in set(agg_regions[colonne_admin])]
            palette_admin = palette_entites(entites_fond)

            # Le découpage fin est tracé en entier : le cadrage montre ce qui
            # compte. Seuls les NOMS sont restreints à la zone affichée, sans
            # quoi 360 étiquettes se superposeraient.
            geojson_sous, centres_sous, noms_sous = None, None, None
            if afficher_sous and niveau_sous:
                geojson_sous, _, centres_sous = charger_fond(niveau_sous)
                if afficher_noms_carte and entite_carte != TOUT_LE_CAMEROUN:
                    noms_sous = sorted(entites_contenues(
                        centres_sous, geojson_admin, entite_carte))

            carte_col, legende_col = st.columns([3, 1])
            with carte_col, st.container(border=True):
                fig_choro = carte_choroplethe(
                    agg_regions, geojson_admin, metrique_carte, entite_carte, mode_couleur,
                    colonne=colonne_admin, ordre=ordre_entites, couleurs=palette_admin,
                    centres=centres_fond, sous_decoupage=geojson_sous,
                    centres_sous=centres_sous, afficher_noms=afficher_noms_carte,
                    noms_sous=noms_sous)
                st.plotly_chart(fig_choro, width='stretch', config=CONFIG_CARTE)
            with legende_col:
                st.markdown(
                    legende_regions_html(
                        agg_regions, metrique_carte, entite_carte, mode_couleur,
                        colonne=colonne_admin, ordre=ordre_entites, couleurs=palette_admin,
                        singulier=cfg_niveau["singulier"], feminin=cfg_niveau["feminin"]),
                    unsafe_allow_html=True)
                if isoler_centrale:
                    n_exclus = int(df_carte.loc[df_carte["Administration"] == "Centrale",
                                                "Quantité"].sum())
                    st.caption(f"{fmt_int(n_exclus)} biens de l'administration centrale "
                               "ne sont pas coloriés sur la carte (aucun territoire ne "
                               "leur correspond) ; ils figurent sur leur propre ligne "
                               "dans le tableau ci-dessous.")

            # --- Statistiques de la zone cartographiée --------------------------
            st.markdown("<div class='section-title'>Statistiques de la zone cartographiée</div>",
                        unsafe_allow_html=True)
            if entite_carte == TOUT_LE_CAMEROUN:
                zone_stats = df_choro
                etiquette_zone = "Tout le Cameroun"
            else:
                # Le filtrage passe par l'entité appariée du fond, pas par le
                # libellé brut : les deux orthographes coexistent dans la base.
                zone_stats = df_zone[df_zone["_entite"] == entite_carte]
                etiquette_zone = entite_carte

            if zone_stats.empty:
                st.info(f"Aucun bien recensé dans « {etiquette_zone} » avec les filtres actuels.")
            else:
                k = synthese_kpi(zone_stats, anomalies[anomalies["_index_ligne"].isin(zone_stats.index)])
                m1, m2, m3, m4, m5 = st.columns(5)
                m1.metric("Biens", fmt_int(k["n_biens"]))
                m2.metric("VNC", fmt_fcfa(k["vnc_totale"]))
                m3.metric("Valeur d'acquisition", fmt_fcfa(k["va_totale"]))
                m4.metric("À réformer", fmt_int(k["n_a_reformer"]),
                          delta=f"{k['part_a_reformer']*100:.1f} %".replace(".", ","),
                          delta_color="inverse")
                m5.metric("Indice de santé",
                          "n.d." if k["indice_sante"] is None
                          else f"{k['indice_sante']:.1f}".replace(".", ","))

                s1, s2 = st.columns(2)
                with s1:
                    par_etat = repartition_etats(zone_stats).reset_index()
                    total_etat = par_etat["Nombre de biens"].sum() or 1
                    par_etat["Étiquette"] = (par_etat["Nombre de biens"] / total_etat * 100).map(
                        lambda x: f"{x:.1f}".replace(".", ",") + " %")
                    fig = px.bar(par_etat, x="État du bien", y="Nombre de biens",
                                 color="État du bien", text="Étiquette",
                                 color_discrete_map=COULEURS_ETAT,
                                 category_orders={"État du bien": ORDRE_ETATS})
                    fig.update_traces(textposition="outside")
                    st.plotly_chart(apply_layout(fig, 360, f"État des biens · {etiquette_zone}"),
                                    width='stretch', config=CONFIG_PLOTLY)
                with s2:
                    par_nature = (zone_stats.groupby("Nature du matériel", as_index=False)
                                  .agg(Biens=("Quantité", "sum"))
                                  .sort_values("Biens", ascending=False).head(10))
                    fig = px.bar(par_nature.sort_values("Biens"), y="Nature du matériel",
                                 x="Biens", orientation="h", text="Biens",
                                 color_discrete_sequence=[THEME["vert"]])
                    fig.update_traces(textposition="outside")
                    st.plotly_chart(apply_layout(fig, 360, f"Nature du matériel · {etiquette_zone}"),
                                    width='stretch', config=CONFIG_PLOTLY)

                # Tableau territorial : quand les services centraux sont isolés,
                # ils forment une ligne à part entière et sont déduits du Centre.
                # Sans cette ligne, leurs biens — les plus nombreux de la base —
                # disparaîtraient purement et simplement du tableau.
                # La ligne « Administration centrale » n'a de sens qu'au niveau
                # régional : plus bas, les services centraux se répartissent
                # déjà entre le Mfoundi et les départements voisins.
                if isoler_centrale and niveau_admin == NIVEAU_REGION:
                    agg_tableau = agreger_par_region(df_carte, "Zone territoriale",
                                                     nom_colonne="Zone territoriale")
                    agg_tableau = agg_tableau[
                        agg_tableau["Zone territoriale"].isin(
                            entites_fond + ["Administration centrale"])]
                    agg_tableau = completer_entites(agg_tableau, "Zone territoriale")
                    colonne_zone = "Zone territoriale"
                    ordre_zones = (["Administration centrale"]
                                   + [r for r in ORDRE_REGIONS
                                      if r in set(agg_tableau[colonne_zone])])
                else:
                    agg_tableau = agg_regions
                    colonne_zone = colonne_admin
                    ordre_zones = list(ordre_entites)

                if entite_carte != TOUT_LE_CAMEROUN:
                    # Entité sélectionnée : sa ligne, plus celle des services
                    # centraux quand la région choisie est le Centre.
                    ordre_zones = [entite_carte] + (
                        ["Administration centrale"]
                        if (isoler_centrale and niveau_admin == NIVEAU_REGION
                            and entite_carte == "Centre") else [])
                ordre_zones = [z for z in ordre_zones if z in set(agg_tableau[colonne_zone])]

                st.dataframe(
                    agg_tableau.set_index(colonne_zone).reindex(ordre_zones).style.format({
                        "Nombre de biens": "{:,.0f}", "Nombre de lignes d'inventaire": "{:,.0f}",
                        "Valeur nette comptable (FCFA)": "{:,.0f}",
                        "Valeur d'acquisition (FCFA)": "{:,.0f}",
                        "Biens à réformer": "{:,.0f}", "Structures": "{:,.0f}",
                        "Indice de santé": "{:.1f}", "À réformer (%)": "{:.1f}"},
                        thousands=" ", decimal=",", na_rep="n.d."), width='stretch')

    # -------------------------------------------------------------------------
    # 2.b Carte détaillée : 1 point = 1 ligne d'inventaire
    # -------------------------------------------------------------------------
    if vue_carte == "Carte détaillée (points)":
        ministeres_points = sorted(df["Ministère"].dropna().unique().tolist())
        p1, p2, p3 = st.columns(3)
        with p1:
            ministere_points = st.selectbox(
                "Ministère cartographié", [TOUS_MINISTERES_CARTE] + ministeres_points,
                index=0, key="ministere_points",
                format_func=lambda m: (m if m == TOUS_MINISTERES_CARTE
                                       else libelle_ministere(m)),
                help="Restreint les points affichés à un seul ministère, sans "
                     "toucher au reste de l'application.")
        with p2:
            # Projeter un objet précis (les ordinateurs, les véhicules…) pour
            # lire sa répartition sur le territoire : c'est la question que se
            # pose un décideur avant d'arbitrer une dotation.
            natures_points = sorted(df["Nature du matériel"].dropna().unique().tolist())
            nature_points = st.selectbox(
                "Nature du matériel", [TOUS_BIENS_CARTE] + natures_points,
                index=0, key="nature_points",
                help="Filtre par famille de matériel : informatique, roulant, "
                     "mobilier…")
        with p3:
            biens_dispo_carte = sorted(
                (df if nature_points == TOUS_BIENS_CARTE
                 else df[df["Nature du matériel"] == nature_points])
                ["Libellé standardisé"].dropna().unique().tolist())
            bien_points = st.selectbox(
                "Bien projeté", [TOUS_BIENS_CARTE] + biens_dispo_carte,
                index=0, key="bien_points",
                help="Projette un seul objet sur la carte — par exemple les "
                     "ordinateurs — pour voir comment il se répartit.")

        df_map = (df if ministere_points == TOUS_MINISTERES_CARTE
                  else df[df["Ministère"] == ministere_points])
        if nature_points != TOUS_BIENS_CARTE:
            df_map = df_map[df_map["Nature du matériel"] == nature_points]
        if bien_points != TOUS_BIENS_CARTE:
            df_map = df_map[df_map["Libellé standardisé"] == bien_points]
        df_map = df_map.copy()
        df_map["Zone cartographique"] = np.where(
            df_map["Administration"].eq("Centrale"),
            "Direction centrale - Yaoundé",
            df_map["Région"].fillna("Non renseigné"),
        )
        df_map["Zone cartographique"] = df_map["Zone cartographique"].replace({"Centre": "Centre - services déconcentrés"})

        zone_order = [
            "Direction centrale - Yaoundé", "Adamaoua", "Centre - services déconcentrés", "Est", "Extrême-Nord",
            "Littoral", "Nord", "Nord-Ouest", "Ouest", "Sud", "Sud-Ouest"
        ]
        zone_coords = {
            "Direction centrale - Yaoundé": (3.875, 11.56),
            "Adamaoua": COORD_REGIONS["Adamaoua"],
            "Centre - services déconcentrés": (3.55, 11.20),
            "Est": COORD_REGIONS["Est"],
            "Extrême-Nord": COORD_REGIONS["Extrême-Nord"],
            "Littoral": COORD_REGIONS["Littoral"],
            "Nord": COORD_REGIONS["Nord"],
            "Nord-Ouest": COORD_REGIONS["Nord-Ouest"],
            "Ouest": COORD_REGIONS["Ouest"],
            "Sud": COORD_REGIONS["Sud"],
            "Sud-Ouest": COORD_REGIONS["Sud-Ouest"],
        }
        zone_list = [TOUT_LE_CAMEROUN] + [z for z in zone_order if z in df_map["Zone cartographique"].unique()]

        c1, c2, c3 = st.columns([1.2, 1, 1.3])
        with c1:
            zone_sel = st.selectbox("Zone à visualiser", zone_list, index=0, key="zone_points")
        with c2:
            type_bien_sel = st.selectbox("Type de bien", ["Tous", "Immobilisation", "Stock"],
                                         index=0, key="type_bien_points")
        with c3:
            # Le fond de cette carte se règle comme celui de la choroplèthe :
            # on peut y lire les départements ou les arrondissements, que la
            # base les chiffre ou non.
            fond_points = st.selectbox(
                "Découpage affiché en fond", list(NIVEAUX_ADMIN.keys()), index=0,
                key="fond_points",
                help="Trace les limites du niveau choisi sous les points.")
        n1, n2, n3 = st.columns(3)
        with n1:
            afficher_noms_points = st.toggle(
                "Afficher les noms des subdivisions", value=True, key="noms_points")
        with n2:
            colorer_fond_points = st.toggle(
                "Colorer chaque subdivision", value=True, key="couleur_fond_points",
                help="Donne à chaque territoire sa propre teinte, en pastel pour "
                     "ne pas masquer les points. Décochez pour un fond neutre.")
        with n3:
            sous_fond_points = st.toggle(
                "Ajouter les limites du niveau inférieur", value=False,
                key="sous_fond_points",
                help="Superpose un découpage plus fin en trait léger.")

        # Chaque zone de cette carte correspond à une région du fond
        # administratif : les deux zones « Yaoundé » et « Centre - services
        # déconcentrés » découpent la même région du Centre.
        REGION_DE_LA_ZONE = {
            "Direction centrale - Yaoundé": "Centre",
            "Centre - services déconcentrés": "Centre",
        }

        # Cadrage resserré sur le Cameroun (moins de pays voisins visibles par
        # défaut) pour davantage mettre le pays en évidence.
        centre_carte = dict(lat=7.3, lon=12.3)
        zoom_carte = 6.0
        regions_contours = None
        if zone_sel != TOUT_LE_CAMEROUN:
            lat0, lon0 = zone_coords.get(zone_sel, (6.0, 12.3))
            centre_carte = dict(lat=lat0, lon=lon0)
            zoom_carte = 7.2 if zone_sel == "Direction centrale - Yaoundé" else 6.4
            # Comme sur la choroplèthe : la zone choisie s'affiche seule, le
            # reste du pays est retiré du fond.
            regions_contours = [REGION_DE_LA_ZONE.get(zone_sel, zone_sel)]

        d = df_map.dropna(subset=["lat", "lon"]).copy()
        if zone_sel != TOUT_LE_CAMEROUN:
            d = d[d["Zone cartographique"] == zone_sel]
        if type_bien_sel != "Tous":
            d = d[d["Classification"] == type_bien_sel]

        libelle_projection = (bien_points if bien_points != TOUS_BIENS_CARTE
                              else nature_points if nature_points != TOUS_BIENS_CARTE
                              else "tous les biens")

        with st.container(border=True):
            geojson_pts, entites_pts, centres_pts = charger_fond(fond_points)
            # Une zone sélectionnée s'affiche seule, comme sur la choroplèthe.
            # Aux niveaux fins, on ne garde que les entités dont le centre tombe
            # réellement dans la région (sans repli, qui ferait entrer les
            # départements voisins dans la vue).
            contours_pts = regions_contours
            if zone_sel != TOUT_LE_CAMEROUN and fond_points != NIVEAU_REGION:
                region_zone = REGION_DE_LA_ZONE.get(zone_sel, zone_sel)
                contenues = entites_contenues(centres_pts, charger_fond(NIVEAU_REGION)[0],
                                              region_zone, avec_repli=False)
                contours_pts = sorted(contenues) or None
            elif zone_sel == TOUT_LE_CAMEROUN:
                contours_pts = None

            niveaux_inferieurs_pts = {NIVEAU_REGION: NIVEAU_DEPARTEMENT,
                                      NIVEAU_DEPARTEMENT: NIVEAU_ARRONDISSEMENT}
            niveau_sous_pts = niveaux_inferieurs_pts.get(fond_points)
            geojson_sous_pts, centres_sous_pts, noms_sous_pts = None, None, None
            if sous_fond_points and niveau_sous_pts:
                geojson_sous_pts, _, centres_sous_pts = charger_fond(niveau_sous_pts)
                if afficher_noms_points and zone_sel != TOUT_LE_CAMEROUN:
                    region_zone = REGION_DE_LA_ZONE.get(zone_sel, zone_sel)
                    noms_sous_pts = sorted(entites_contenues(
                        centres_sous_pts, charger_fond(NIVEAU_REGION)[0], region_zone))

            disperser = st.toggle("Séparer visuellement les points superposés (positions artificielles)",
                                  value=False, key="dispersion_v7")
            st.caption("Sans dispersion, plusieurs lignes peuvent se superposer sur un même repère. "
                       "La carte ne permet pas de localiser précisément un bien.")
            if disperser:
                st.warning("Dispersion illustrative : jusqu’à environ 22 km du repère. "
                           "Ne pas interpréter les points comme des positions ou une densité réelles.")
            fig_map = carte_points(
                d, centre_carte, zoom_carte,
                f"{zone_sel} · {libelle_projection} · {fmt_int(len(d))} lignes "
                "· couleur : état du bien",
                geojson=geojson_pts, regions_contours=contours_pts,
                rayon_dispersion=0.20 if disperser else 0.0,
                centres=centres_pts, afficher_noms=afficher_noms_points,
                sous_decoupage=geojson_sous_pts, centres_sous=centres_sous_pts,
                noms_sous=noms_sous_pts,
                couleurs_fond=(palette_entites(contours_pts or entites_pts)
                               if colorer_fond_points else None))
            st.plotly_chart(fig_map, width='stretch', config=CONFIG_CARTE)

        st.subheader("Statistiques par description des biens cartographiés")
        st.caption(
            "Ce tableau reprend les lignes géolocalisées de la carte, après les filtres "
            "globaux et ceux de la carte : ministère, nature, bien, zone et type. "
            "Le zoom et la dispersion visuelle ne modifient pas les statistiques. "
            "Les descriptions absentes sont regroupées ; n.d. indique un montant non renseigné."
        )
        tableau_descriptions_carte = statistiques_descriptions(d)
        if tableau_descriptions_carte.empty:
            st.info("Aucun bien cartographié ne correspond à ces filtres.")
        else:
            st.dataframe(
                tableau_descriptions_carte.style.format({
                    "Lignes d'inventaire": "{:,.0f}",
                    "Quantité totale": "{:,.0f}",
                    "Valeur d'acquisition totale (FCFA)": "{:,.0f}",
                    "VNC totale (FCFA)": "{:,.0f}",
                }, thousands=" ", decimal=",", na_rep="n.d."),
                hide_index=True, width="stretch", height=480,
            )
            st.download_button(
                "Télécharger les statistiques de la carte (Excel)",
                to_excel_bytes({"Descriptions carte": tableau_descriptions_carte}),
                file_name="statistiques_descriptions_carte.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="export_descriptions_carte",
            )

        # Légende détaillée sous la carte : une ligne par couleur, avec le
        # nombre de biens et la part qu'elle représente dans la zone affichée.
        lignes_legende = repartition_etats_carte(d)
        if lignes_legende:
            total_zone = sum(l["biens"] for l in lignes_legende)
            morceaux = []
            for l in lignes_legende:
                part = f"{l['part']:.1f}".replace(".", ",")
                morceaux.append(
                    f"<div class='legend-row'>"
                    f"<span class='legend-chip' style='background:{l['couleur']}'></span>"
                    f"<span class='legend-name'>{l['etat']}</span>"
                    f"<span class='legend-val'>{fmt_int(l['biens'])} biens · {part} %"
                    f"</span></div>")
            corps = "".join(morceaux)
            st.markdown(
                f"<div class='legend-box'><div class='legend-title'>"
                f"Répartition par état · {zone_sel}</div>"
                f"<div class='legend-sub'>{fmt_int(total_zone)} biens "
                f"({fmt_int(len(d))} lignes d'inventaire) — couleurs de la carte"
                f"</div>{corps}</div>", unsafe_allow_html=True)

        df_non_geo = df_map[df_map["Source géolocalisation"] == "Non géolocalisé"]
        if len(df_non_geo) or rapport["villes_non_geolocalisees"]:
            with st.expander("Voir les biens non géolocalisés / villes à compléter"):
                if rapport["villes_non_geolocalisees"]:
                    st.write("Villes sans coordonnées GPS : " + ", ".join(rapport["villes_non_geolocalisees"][:80]))
                if len(df_non_geo):
                    st.dataframe(df_non_geo[["Libellé standardisé", "Région", "Département", "Ville", "Structure agrégée", "Valeur nette comptable"]], width='stretch', height=260)

# -----------------------------------------------------------------------------
# Onglet 3 : Statistiques du périmètre filtré
# -----------------------------------------------------------------------------
if module == "Statistiques":
    st.markdown("<div class='section-title'>Statistiques du périmètre filtré</div>",
                unsafe_allow_html=True)
    st.markdown(bandeau_filtres, unsafe_allow_html=True)
    if filtres_actifs:
        part_biens = kpi["n_biens"] / max(float(df0["Quantité"].fillna(0).sum()), 1.0)
        part_vnc = kpi["vnc_totale"] / max(float(df0["Valeur nette comptable"].sum(skipna=True)), 1.0)
        st.markdown(
            "<div class='small-muted'>Le périmètre filtré représente "
            f"<b>{f'{part_biens*100:.1f}'.replace('.', ',')} %</b> des biens et "
            f"<b>{f'{part_vnc*100:.1f}'.replace('.', ',')} %</b> de la valeur nette "
            "comptable de la base chargée.</div>", unsafe_allow_html=True)

    # --- Compteurs principaux -------------------------------------------------
    s1, s2, s3, s4, s5, s6 = st.columns(6)
    s1.metric("Biens", fmt_int(kpi["n_biens"]))
    s2.metric("Lignes d'inventaire", fmt_int(kpi["n_lignes"]))
    s3.metric("VNC", fmt_fcfa(kpi["vnc_totale"]))
    s4.metric("Valeur d'acquisition", fmt_fcfa(kpi["va_totale"]))
    s5.metric("Structures", fmt_int(kpi["n_structures"]))
    s6.metric("Régions couvertes", fmt_int(kpi["n_regions"]))

    e1, e2, e3, e4 = st.columns(4)
    e1.metric("Neufs", fmt_int(kpi["n_neufs"]),
              delta=f"{kpi['part_neufs']*100:.1f} %".replace(".", ","))
    e2.metric("Bon état", fmt_int(kpi["n_bon_etat"]),
              delta=f"{kpi['part_bon_etat']*100:.1f} %".replace(".", ","))
    e3.metric("À réformer", fmt_int(kpi["n_a_reformer"]),
              delta=f"{kpi['part_a_reformer']*100:.1f} %".replace(".", ","),
              delta_color="inverse")
    e4.metric("Taux d'amortissement",
              f"{kpi['taux_amortissement']*100:.1f} %".replace(".", ","))

    st.subheader("Statistiques par description des biens")
    tableau_descriptions = statistiques_descriptions(df)
    st.caption(
        f"{len(tableau_descriptions)} descriptions ou groupes affichés · "
        "Une ligne par description identique, après retrait des espaces en début et fin. "
        "Les descriptions absentes restent regroupées. Les filtres sélectionnés s’appliquent. "
        "Les montants manquants sont ignorés dans les sommes ; n.d. signifie qu’aucun montant n’est renseigné."
    )
    st.dataframe(
        tableau_descriptions.style.format({
            "Lignes d'inventaire": "{:,.0f}",
            "Quantité totale": "{:,.0f}",
            "Valeur d'acquisition totale (FCFA)": "{:,.0f}",
            "VNC totale (FCFA)": "{:,.0f}",
        }, thousands=" ", decimal=",", na_rep="n.d."),
        hide_index=True, width="stretch", height=480,
    )
    st.download_button(
        "Télécharger le tableau par description (Excel)",
        to_excel_bytes({"Statistiques descriptions": tableau_descriptions}),
        file_name="statistiques_descriptions.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="export_statistiques_descriptions",
    )

    # Le décompte des états imputés n'est plus rappelé ici : il encombrait la
    # lecture des statistiques. L'information reste disponible là où elle sert
    # vraiment — le contrôle A07 de l'audit et la colonne « État imputé » de la
    # page Données.

    # --- Répartitions ---------------------------------------------------------
    r1, r2 = st.columns(2)
    with r1, st.container(border=True):
        par_etat = repartition_etats(df).reset_index()
        total_etat = par_etat["Nombre de biens"].sum() or 1
        par_etat["Part"] = 100 * par_etat["Nombre de biens"] / total_etat
        par_etat["Étiquette"] = par_etat.apply(
            lambda r: f"{fmt_int(r['Nombre de biens'])} ({r['Part']:.1f} %)".replace(".", ","), axis=1)
        fig = px.bar(par_etat, x="État du bien", y="Nombre de biens", color="État du bien",
                     text="Étiquette", color_discrete_map=COULEURS_ETAT,
                     category_orders={"État du bien": ORDRE_ETATS})
        fig.update_traces(textposition="outside")
        st.plotly_chart(apply_layout(fig, 380, "Répartition par état du bien"),
                        width='stretch', config=CONFIG_PLOTLY)
    with r2, st.container(border=True):
        par_nature = (df.groupby("Nature du matériel", as_index=False)
                      .agg(Biens=("Quantité", "sum"), VNC=("Valeur nette comptable", "sum"))
                      .sort_values("Biens", ascending=False))
        total_nature = par_nature["Biens"].sum() or 1
        par_nature["Étiquette"] = par_nature.apply(
            lambda r: f"{fmt_int(r['Biens'])} ({100*r['Biens']/total_nature:.1f} %)".replace(".", ","), axis=1)
        fig = px.bar(par_nature.sort_values("Biens"), y="Nature du matériel", x="Biens",
                     orientation="h", text="Étiquette", color_discrete_sequence=[THEME["vert"]])
        fig.update_traces(textposition="outside")
        st.plotly_chart(apply_layout(fig, 380, "Répartition par nature du matériel"),
                        width='stretch', config=CONFIG_PLOTLY)

    # --- Classements ----------------------------------------------------------
    st.markdown("<div class='section-title'>Classements</div>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1.4, 1.4, 1.2])
    with c1:
        dimension = st.selectbox(
            "Dimension classée",
            ["Zone territoriale", "Région", "Département", "Ville", "Structure agrégée",
             "Libellé standardisé", "Catégorie", "Nature du matériel", "Ministère",
             "Administration"], key="dimension_stats",
            help="« Zone territoriale » isole les services centraux de la région "
                 "du Centre, qu'ils gonfleraient sinon artificiellement.")
    with c2:
        critere = st.selectbox("Classé par",
                               ["Nombre de biens", "VNC totale (FCFA)",
                                "Valeur d'acquisition (FCFA)", "Anomalies"])
    with c3:
        # Le top 10 masquait des lignes utiles ; le nombre d'entrées est réglable
        # et « Tout afficher » retire complètement la troncature.
        nb_lignes = st.selectbox("Entrées affichées", ["Tout afficher", 10, 20, 50],
                                 index=0, key="nb_lignes_stats")

    classement = stats_par(df, dimension).sort_values(critere, ascending=False)
    classement_complet = classement.copy()
    if nb_lignes != "Tout afficher":
        classement = classement.head(int(nb_lignes))

    g1, g2 = st.columns([1.3, 1])
    with g1:
        graphe = classement.reset_index().rename(columns={dimension: "Modalité"})
        graphe["Modalité"] = graphe["Modalité"].fillna("Non renseigné").astype(str)
        graphe["Étiquette"] = (graphe[critere].map(fmt_compact)
                               if "FCFA" in critere else graphe[critere].map(fmt_int))
        hauteur = max(360, 26 * len(graphe) + 90)
        fig = px.bar(graphe.sort_values(critere), y="Modalité", x=critere, orientation="h",
                     text="Étiquette", color_discrete_sequence=[THEME["or"]])
        fig.update_traces(textposition="outside")
        st.plotly_chart(apply_layout(fig, min(hauteur, 1200), f"{critere} par {dimension.lower()}"),
                        width='stretch', config=CONFIG_PLOTLY)
    with g2:
        st.dataframe(classement.style.format({
            "Nombre de biens": "{:,.0f}", "VNC totale (FCFA)": "{:,.0f}",
            "Valeur d'acquisition (FCFA)": "{:,.0f}", "Anomalies": "{:,.0f}",
            "Indice de santé": "{:.1f}"},
            thousands=" ", decimal=",", na_rep="n.d."),
            width='stretch', height=min(max(360, 26 * len(classement) + 90), 1200))

    # Les modalités absentes ou à zéro sont une information de gestion à part
    # entière (zones non recensées, structures sans dotation).
    vides = classement_complet[classement_complet["Nombre de biens"] <= 0]
    if len(vides):
        with st.expander(f"{len(vides)} {dimension.lower()}(s) sans aucun bien recensé"):
            st.dataframe(vides, width='stretch')

    # --- Détail d'un type d'objet --------------------------------------------
    st.markdown("<div class='section-title'>Analyser un type d'objet "
                "(chaises, ordinateurs, véhicules…)</div>", unsafe_allow_html=True)
    st.markdown(
        "<div class='small-muted'>Choisissez un objet ci-dessous : l'application "
        "indique <b>combien l'État en possède</b>, <b>dans quelles régions et quelles "
        "structures ils se trouvent</b>, <b>dans quel état</b> et <b>pour quelle "
        "valeur</b>.<br>Exemple d'usage : un ministère demande l'achat de chaises. "
        "Vous sélectionnez « Chaise » et vous voyez combien il en détient déjà, où "
        "elles sont et combien sont à réformer — de quoi accorder ou refuser la "
        "dotation sur pièces.</div>", unsafe_allow_html=True)

    biens_dispo = sorted(df["Libellé standardisé"].dropna().unique().tolist())
    o1, o2 = st.columns([1.4, 2])
    with o1:
        bien_focus = st.selectbox(
            "👉 Objet à analyser", biens_dispo,
            index=0 if biens_dispo else None, key="bien_focus",
            help="Ce choix ne concerne que cette section ; il ne modifie pas le "
                 "reste de la page. Pour restreindre toute la page à un objet, "
                 "utilisez « Objet précis » dans les filtres en haut.")
    with o2:
        if bien_focus:
            sous_ensemble = df[df["Libellé standardisé"] == bien_focus]
            st.markdown(
                "<div class='filter-strip' style='margin-top:26px'>"
                f"<span class='filter-pill'><b>Objet analysé</b> · {bien_focus}</span>"
                f"<span class='filter-pill'>{fmt_int(sous_ensemble['Quantité'].sum())} "
                f"biens · {fmt_int(len(sous_ensemble))} lignes</span>"
                f"<span class='filter-pill'>{sous_ensemble['Région'].nunique()} "
                f"région(s) · {sous_ensemble['Structure agrégée'].nunique()} "
                "structure(s)</span></div>", unsafe_allow_html=True)
    if bien_focus:
        sub = df[df["Libellé standardisé"] == bien_focus]
        f1, f2, f3, f4 = st.columns(4)
        f1.metric("Quantité", fmt_int(sub["Quantité"].sum()))
        f2.metric("VNC", fmt_fcfa(sub["Valeur nette comptable"].sum()))
        f3.metric("Valeur d'acquisition", fmt_fcfa(sub["Valeur d'acquisition"].sum()))
        pu_median = sub["Prix unitaire"].median()
        f4.metric("Prix unitaire médian", fmt_fcfa(pu_median))

        b1, b2 = st.columns(2)
        with b1:
            rep_geo = (sub.groupby("Région", as_index=False).agg(Biens=("Quantité", "sum"))
                       .sort_values("Biens", ascending=False))
            fig = px.bar(rep_geo.sort_values("Biens"), y="Région", x="Biens",
                         orientation="h", text="Biens",
                         color_discrete_sequence=[THEME["bleu"]])
            fig.update_traces(textposition="outside")
            st.plotly_chart(apply_layout(fig, 360, f"{bien_focus} par région"),
                            width='stretch', config=CONFIG_PLOTLY)
        with b2:
            rep_etat = (sub.groupby("État du bien", as_index=False).agg(Biens=("Quantité", "sum")))
            total_focus = rep_etat["Biens"].sum() or 1
            rep_etat["Étiquette"] = rep_etat["Biens"].map(
                lambda x: f"{100*x/total_focus:.1f}".replace(".", ",") + " %")
            fig = px.pie(rep_etat, names="État du bien", values="Biens", hole=0.55,
                         color="État du bien", color_discrete_map=COULEURS_ETAT)
            fig.update_traces(text=rep_etat["Étiquette"], textinfo="text")
            st.plotly_chart(apply_layout(fig, 360, f"État des « {bien_focus} »"),
                            width='stretch', config=CONFIG_PLOTLY)

        st.dataframe(
            sub.groupby(["Région", "Ville", "Structure agrégée"], dropna=False)
            .agg(Biens=("Quantité", "sum"),
                 VNC=("Valeur nette comptable", "sum"),
                 Lignes=("Quantité", "size"))
            .sort_values("Biens", ascending=False)
            .style.format({"Biens": "{:,.0f}", "VNC": "{:,.0f}", "Lignes": "{:,.0f}"},
                          thousands=" ", decimal=",", na_rep="n.d."),
            width='stretch', height=340)

# -----------------------------------------------------------------------------
# Onglet 4 : Analyse du patrimoine
# -----------------------------------------------------------------------------
if module == "Analyse du patrimoine":
    st.markdown("<div class='section-title'>Analyse du patrimoine</div>", unsafe_allow_html=True)
    st.markdown(bandeau_filtres, unsafe_allow_html=True)
    st.markdown("<div class='small-muted'>Analyses séparées entre immobilisations, stocks et comparaison globale. "
                "Les valeurs chiffrées sont affichées sur chaque barre (lecture possible sans référence à la couleur).</div>", unsafe_allow_html=True)

    # --- Réglages communs aux analyses ---------------------------------------
    a1, a2, a3 = st.columns([1.2, 1.4, 1.4])
    with a1:
        type_graphique = st.selectbox(
            "Type de graphique", TYPES_GRAPHIQUES, index=0, key="type_graphique_analyse",
            help="Les barres comparent des grandeurs, l'anneau et le camembert "
                 "montrent des parts d'un tout, le treemap combine les deux.")
    with a2:
        isoler_centrale_analyse = st.toggle(
            "Isoler les services centraux de la région du Centre", value=True,
            key="isoler_centrale_analyse",
            help="Les services centraux siègent tous à Yaoundé : les laisser dans "
                 "le Centre surestime cette région. Ils forment alors une "
                 "modalité propre, « Administration centrale ».")
    with a3:
        limite_modalites = st.selectbox(
            "Modalités affichées", ["Toutes", 10, 20, 30], index=0,
            key="limite_modalites_analyse",
            help="Par défaut, chaque graphique affiche la totalité des régions, "
                 "départements ou villes — aucun classement tronqué.")

    def colonne_territoire(d):
        """Région, sauf pour les services centraux qui forment leur propre
        modalité — sans quoi le Centre absorbe tout le patrimoine de l'État."""
        if isoler_centrale_analyse and "Zone territoriale" in d.columns:
            return "Zone territoriale"
        return "Région"

    def bloc_descriptif(d, titre, couleur):
        if d.empty:
            st.info(f"Aucune donnée pour {titre.lower()}.")
            return
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Biens", fmt_int(d["Quantité"].sum()))
        k2.metric("Lignes", fmt_int(len(d)))
        k3.metric("VNC", fmt_fcfa(d["Valeur nette comptable"].sum()))
        k4.metric("Valeur d'acquisition", fmt_fcfa(d["Valeur d'acquisition"].sum()))

        territoire = colonne_territoire(d)
        g1, g2 = st.columns(2)
        with g1:
            st.plotly_chart(
                figure_repartition(d, territoire, "Valeur nette comptable",
                                   type_graphique, couleur, limite_modalites,
                                   f"{titre} par {territoire.lower()} (VNC, FCFA)",
                                   monetaire=True),
                width='stretch', config=CONFIG_PLOTLY)
        with g2:
            st.plotly_chart(
                figure_repartition(d, "Ville", "Valeur nette comptable",
                                   type_graphique, THEME["bleu"], limite_modalites,
                                   f"{titre} par ville (VNC, FCFA)", monetaire=True),
                width='stretch', config=CONFIG_PLOTLY)

        g3, g4 = st.columns(2)
        with g3:
            st.plotly_chart(
                figure_repartition(d, "Département", "Quantité", type_graphique,
                                   THEME["vert2"], limite_modalites,
                                   f"{titre} par département (biens)"),
                width='stretch', config=CONFIG_PLOTLY)
        with g4:
            etat = d.groupby([territoire, "État du bien"], as_index=False).agg(
                Biens=("Quantité", "sum"))
            ordre_territoires = (etat.groupby(territoire)["Biens"].sum()
                                 .sort_values(ascending=False).index.tolist())
            if limite_modalites != "Toutes":
                ordre_territoires = ordre_territoires[:int(limite_modalites)]
                etat = etat[etat[territoire].isin(ordre_territoires)]
            hauteur = max(420, 22 * len(ordre_territoires) + 160)
            fig = px.bar(etat, x=territoire, y="Biens", color="État du bien",
                         barmode="stack", color_discrete_map=COULEURS_ETAT,
                         category_orders={"État du bien": ORDRE_ETATS,
                                          territoire: ordre_territoires})
            st.plotly_chart(apply_layout(fig, min(hauteur, 900),
                                         f"{titre} par état du bien"),
                            width='stretch', config=CONFIG_PLOTLY)

    onglets_analyse = ["Immobilisations", "Stocks", "Comparaison immo / stock",
                       "Administration centrale"]
    if len(ministeres_perimetre) > 1:
        onglets_analyse.append("Comparaison ministères")
    vue_analyse = st.segmented_control(
        "Type d'analyse",
        onglets_analyse,
        default=onglets_analyse[0],
        key="vue_analyse",
        selection_mode="single",
        label_visibility="collapsed",
        width="stretch",
    )

    if vue_analyse == "Immobilisations":
        bloc_descriptif(df[df["Classification"] == "Immobilisation"].copy(), "Immobilisations", THEME["vert"])
    if vue_analyse == "Stocks":
        bloc_descriptif(df[df["Classification"] == "Stock"].copy(), "Stocks", THEME["orange"])
    if vue_analyse == "Comparaison immo / stock":
        comp = df.groupby("Classification", as_index=False).agg(Biens=("Quantité", "sum"), VNC=("Valeur nette comptable", "sum"))
        comp["Étiquette VNC"] = comp["VNC"].map(fmt_compact)
        c1, c2 = st.columns(2)
        with c1:
            fig = px.bar(comp, x="Classification", y="VNC", color="Classification", color_discrete_map=COULEURS_CLASSIF, text="Étiquette VNC")
            fig.update_traces(textposition="outside")
            fig.update_layout(bargap=0.05)
            st.plotly_chart(apply_layout(fig, 380, "Comparaison de la VNC (FCFA)"), width='stretch', config=CONFIG_PLOTLY)
        with c2:
            fig = px.bar(comp, x="Classification", y="Biens", color="Classification", color_discrete_map=COULEURS_CLASSIF, text="Biens")
            fig.update_traces(textposition="outside")
            fig.update_layout(bargap=0.05)
            st.plotly_chart(apply_layout(fig, 380, "Comparaison du nombre de biens"), width='stretch', config=CONFIG_PLOTLY)

        etat_comp = df.groupby(["État du bien", "Classification"], as_index=False).agg(Biens=("Quantité", "sum"))
        fig = px.bar(etat_comp, x="État du bien", y="Biens", color="Classification", barmode="group", color_discrete_map=COULEURS_CLASSIF, category_orders={"État du bien": ORDRE_ETATS})
        st.plotly_chart(apply_layout(fig, 400, "Comparaison par état du bien"), width='stretch', config=CONFIG_PLOTLY)

    if vue_analyse == "Administration centrale":
        tableau_centrale = statistiques_structures_centrales(df)
        if tableau_centrale.empty:
            st.info("Aucun bien de l'administration centrale dans le périmètre filtré.")
        else:
            st.markdown(
                "<div class='small-muted'>Les services centraux siègent tous à Yaoundé : "
                "la carte les empile sur un seul point. Ils sont donc présentés ici par "
                "ministère de tutelle, puis par structure (directions et services, tels "
                "que saisis dans le classeur).</div>", unsafe_allow_html=True)
            totaux_centrale = tableau_centrale[tableau_centrale["Structure"] == TOTAL_MINISTERE]
            structures_centrale = tableau_centrale[tableau_centrale["Structure"] != TOTAL_MINISTERE]
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Ministères", fmt_int(len(totaux_centrale)))
            c2.metric("Structures", fmt_int(len(structures_centrale)))
            c3.metric("Biens", fmt_int(totaux_centrale["Biens"].sum()))
            c4.metric("VNC", fmt_fcfa(totaux_centrale["VNC (FCFA)"].sum()))

            mesure_centrale = st.segmented_control(
                "Taille des rectangles", MESURES_CENTRALE, default=MESURES_CENTRALE[0],
                key="mesure_centrale", selection_mode="single",
            ) or MESURES_CENTRALE[0]
            st.plotly_chart(
                treemap_structures_centrales(
                    tableau_centrale, mesure_centrale,
                    f"Administration centrale par ministère et structure — {mesure_centrale}"),
                width="stretch", config=CONFIG_PLOTLY)
            st.caption("Cliquez sur un ministère pour l'agrandir, puis sur le bandeau du haut "
                       "pour revenir. Les structures dont la mesure est nulle ne sont pas "
                       "dessinées ; elles figurent dans le tableau.")

            st.subheader("Ministères et structures de l'administration centrale")
            st.caption("La première ligne de chaque ministère donne son total ; ses structures "
                       "suivent, de la VNC la plus élevée à la plus faible.")
            st.dataframe(
                tableau_centrale.style
                .apply(lambda l: [f"background-color: {THEME['or_pale']}"
                                  if l["Structure"] == TOTAL_MINISTERE else ""] * len(l), axis=1)
                .format({
                    "Lignes d'inventaire": "{:,.0f}", "Biens": "{:,.0f}",
                    "Valeur d'acquisition (FCFA)": "{:,.0f}", "VNC (FCFA)": "{:,.0f}",
                    "Part de la VNC du ministère (%)": "{:.1f}", "À réformer (%)": "{:.1f}",
                    "Indice santé": "{:.1f}",
                }, thousands=" ", decimal=",", na_rep="n.d."),
                hide_index=True, width="stretch", height=520)
            st.download_button(
                "Télécharger le tableau de l'administration centrale (Excel)",
                to_excel_bytes({"Administration centrale": tableau_centrale}),
                file_name="administration_centrale_structures.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="export_administration_centrale",
            )

    if vue_analyse == "Comparaison ministères":
            st.markdown("<div class='small-muted'>Comparaison inter-ministérielle du périmètre chargé "
                        "(la feuille CENTRAL contient les administrations centrales du MINTP et du MINAC).</div>",
                        unsafe_allow_html=True)
            comp_m = df.groupby("Ministère", as_index=False).agg(
                Biens=("Quantité", "sum"), VNC=("Valeur nette comptable", "sum"),
                Lignes=("Libellé standardisé", "size"))
            comp_m["Étiquette VNC"] = comp_m["VNC"].map(fmt_compact)
            sante_m = indice_sante_par(df, "Ministère")
            comp_m["Indice santé"] = comp_m["Ministère"].map(sante_m).round(1)
            m1, m2 = st.columns(2)
            with m1:
                fig = px.bar(comp_m, x="Ministère", y="VNC", color="Ministère",
                             color_discrete_map=COULEURS_MINISTERE, text="Étiquette VNC")
                fig.update_traces(textposition="outside")
                st.plotly_chart(apply_layout(fig, 380, "VNC par ministère (FCFA)"), width='stretch', config=CONFIG_PLOTLY)
            with m2:
                fig = px.bar(comp_m, x="Ministère", y="Biens", color="Ministère",
                             color_discrete_map=COULEURS_MINISTERE, text="Biens")
                fig.update_traces(textposition="outside")
                st.plotly_chart(apply_layout(fig, 380, "Biens par ministère"), width='stretch', config=CONFIG_PLOTLY)
            st.dataframe(comp_m.drop(columns=["Étiquette VNC"]).set_index("Ministère").style.format(
                {"Biens": "{:,.0f}", "VNC": "{:,.0f}", "Lignes": "{:,.0f}", "Indice santé": "{:.1f}"},
                thousands=" ", decimal=",", na_rep="n.d."), width='stretch')


# -----------------------------------------------------------------------------
# Onglet 4 : Données & exports
# -----------------------------------------------------------------------------
if module == "Données et exports":
    st.markdown("<div class='section-title'>Données et exports</div>", unsafe_allow_html=True)
    st.markdown(bandeau_filtres, unsafe_allow_html=True)
    vue_donnees = st.segmented_control(
        "Type de données",
        ["Exploration", "Exports"],
        default="Exploration",
        key="vue_donnees",
        selection_mode="single",
        label_visibility="collapsed",
        width="stretch",
    )

    if vue_donnees == "Exploration":
        cols = [c for c in [
            "Identifiant", "Libellé source", "Libellé standardisé", "Catégorie",
            "Nature du matériel", "Description",
            "Ministère", "Ministère (libellé)", "Administration", "Structure agrégée",
            "Région", "Zone territoriale", "Département", "Ville", "Localisation",
            "Mode d'acquisition", "Mode d'acquisition source",
            "État du bien", "État imputé", "VNC imputée", "Classification", "Nature saisie", "Contrôle classification", "Quantité", "Prix unitaire",
            "Valeur d'acquisition", "Amortissement", "Valeur nette comptable", "Date d'acquisition", "Date d'affectation",
            "Ancienneté (années)", "Score santé", "Score risque (0-100)", "Niveau de risque", "Nb anomalies", "Anomalies détectées",
            "Action recommandée principale", "Source géolocalisation", "Feuille source",
        ] if c in df.columns]
        d_view = df[cols]
        st.caption(f"{fmt_int(len(d_view))} lignes dans le périmètre filtré")
        st.dataframe(d_view, width='stretch', height=520)

    if vue_donnees == "Exports":
        st.markdown("<div class='small-muted'>Les exports portent sur le périmètre FILTRÉ courant "
                    f"(ministère(s) : {', '.join(ministeres_perimetre)}).</div>", unsafe_allow_html=True)
        base_exp = df[[c for c in df.columns if not c.startswith("_")]].copy()
        anos_exp = anomalies.drop(columns=["_index_ligne"], errors="ignore")
        plan_exp = plan.copy()
        nomenclature_objets_exp = (
            df.groupby(["Libellé source", "Libellé standardisé", "Catégorie", "Feuille source"], dropna=False)
            .size().reset_index(name="Occurrences")
            .sort_values(["Libellé standardisé", "Occurrences"], ascending=[True, False])
        )
        nomenclature_modes_exp = (
            df.groupby(["Mode d'acquisition source", "Mode d'acquisition", "Feuille source"], dropna=False)
            .size().reset_index(name="Occurrences")
            .sort_values(["Mode d'acquisition", "Occurrences"], ascending=[True, False])
        )
        kpi_min_exp = kpi_par_ministere(df, anomalies)
        stats_reg_exp = agreger_par_region(df, "Région")
        e1, e2, e3 = st.columns(3)
        with e1:
            st.markdown("**Base nettoyée filtrée**")
            st.download_button("CSV base", base_exp.to_csv(index=False).encode("utf-8-sig"), "patrimoine_nettoye.csv", "text/csv")
            st.download_button(
                "Excel base",
                to_excel_bytes({
                    "Base": base_exp,
                    "KPI par ministère": kpi_min_exp,
                    "Statistiques régionales": stats_reg_exp,
                    "Nomenclature objets": nomenclature_objets_exp,
                    "Modes acquisition": nomenclature_modes_exp,
                }),
                "patrimoine_nettoye.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        with e2:
            st.markdown("**Anomalies**")
            st.download_button("CSV anomalies", anos_exp.to_csv(index=False).encode("utf-8-sig"), "anomalies_audit.csv", "text/csv")
            st.download_button("Excel anomalies", to_excel_bytes({"Anomalies": anos_exp}), "anomalies_audit.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with e3:
            st.markdown("**Plan d'action**")
            st.download_button("CSV plan", plan_exp.to_csv(index=False).encode("utf-8-sig"), "plan_action.csv", "text/csv")
            st.download_button("Excel plan", to_excel_bytes({"Plan": plan_exp}), "plan_action.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
