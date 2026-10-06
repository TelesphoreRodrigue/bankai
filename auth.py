"""Authentification légère de l'application Streamlit MINFI."""

import base64
import hmac
import os

import streamlit as st


def image_data_uri(path):
    """Retourne une image locale sous forme de data URI pour les blocs HTML."""
    if not path.exists():
        return ""
    mime = "image/png" if path.suffix.lower() == ".png" else "image/jpeg"
    contenu = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:{mime};base64,{contenu}"


def _lire_identifiants():
    secrets_auth = {}
    try:
        secrets_auth = dict(st.secrets.get("auth", {}))
    except Exception:
        pass
    utilisateur = secrets_auth.get("username") or os.getenv("MINFI_USERNAME")
    mot_de_passe = secrets_auth.get("password") or os.getenv("MINFI_PASSWORD")
    if not utilisateur or not mot_de_passe:
        st.error("Accès désactivé : configurez MINFI_USERNAME et MINFI_PASSWORD, ou la section auth des secrets Streamlit.")
        st.stop()
    return str(utilisateur), str(mot_de_passe)


def require_authentication(logo_uri):
    """Affiche la connexion et interrompt la page tant que l'accès est refusé."""
    utilisateur, mot_de_passe = _lire_identifiants()
    if st.session_state.get("authentifie", False):
        return

    st.markdown(
        "<style>section[data-testid='stSidebar']{display:none}"
        ".block-container{max-width:1050px;padding-top:7vh}</style>",
        unsafe_allow_html=True,
    )
    logo_html = (
        f'<img src="{logo_uri}" alt="Logo du MINFI">'
        if logo_uri
        else "<div style='font-size:64px'>🏛️</div>"
    )
    presentation, formulaire = st.columns([1.12, 0.88], gap="large")
    with presentation:
        st.markdown(
            f"""
            <div class="login-intro">
              {logo_html}
              <div class="eyebrow">Ministère des Finances</div>
              <h1>Pilotage du patrimoine de l'État</h1>
              <p>Une lecture consolidée des immobilisations pour éclairer les décisions,
              suivre les risques et améliorer la gestion du patrimoine public.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with formulaire:
        st.markdown(
            """
            <div class="login-brand">
              <div class="login-kicker">Accès sécurisé</div>
              <h1>Bienvenue</h1>
              <p>Connectez-vous à votre espace de pilotage.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        with st.form("connexion", clear_on_submit=False, border=True):
            st.markdown("### Connexion")
            utilisateur_saisi = st.text_input("Identifiant", placeholder="Votre identifiant")
            mot_de_passe_saisi = st.text_input(
                "Mot de passe", type="password", placeholder="Votre mot de passe"
            )
            valider = st.form_submit_button("Se connecter", width="stretch")

    if valider:
        utilisateur, mot_de_passe = _lire_identifiants()
        acces_valide = hmac.compare_digest(
            utilisateur_saisi.encode("utf-8"), utilisateur.encode("utf-8")
        ) and hmac.compare_digest(mot_de_passe_saisi.encode("utf-8"), mot_de_passe.encode("utf-8"))
        if acces_valide:
            st.session_state["authentifie"] = True
            st.session_state["utilisateur"] = utilisateur_saisi
            st.rerun()
        with formulaire:
            st.error("Identifiant ou mot de passe incorrect.")
    st.stop()


def logout():
    """Ferme la session Streamlit courante."""
    st.session_state.clear()
    st.rerun()
