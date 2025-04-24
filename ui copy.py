# File: ui.py

import streamlit as st
import pandas as pd
import numpy as np
from config import PLACEHOLDERS_DEFAULT
from presentation_generator import generate_presentation

def run_app():
    st.set_page_config(page_title='Générateur ISOEDRE', layout='wide')
    st.title('🚀 Générateur Document de référence - ISOEDRE')

    # Sidebar
    with st.sidebar:
        st.header('⚙️ Paramètres')
        img_dir = st.text_input('Dossier des images', value='img')
        logos_dir = st.text_input('Dossier blasons', value='logos')
        valeur_blason = st.text_input('Valeur blason actif', value='x')
        debug = st.checkbox('Mode debug')
        if debug:
            st.info('Logs mode debug activé')

    # Upload des fichiers
    col1, col2 = st.columns(2)
    with col1:
        fichier_excel = st.file_uploader('📥 Fichier Excel', type=['xlsx'])
    with col2:
        fichier_pptx = st.file_uploader('📥 Modèle PPTX', type=['pptx'])

    # Si on a les fichiers, on affiche la data et propose filtres
    if fichier_excel and fichier_pptx:
        df = pd.read_excel(fichier_excel, header=1, dtype=str)

        # SECTION FILTRAGE & TRI
        st.subheader("🔍 Filtres et tri")
        filtres = {}
        with st.expander("Filtres et options de tri", expanded=True):
            # Filtre sur une colonne
            col_filtre = st.selectbox("Filtrer par colonne", options=["Aucun"] + list(df.columns))
            if col_filtre != "Aucun":
                vals = sorted(df[col_filtre].dropna().unique())
                sel = st.multiselect(f"Valeurs pour {col_filtre}", options=vals)
                if sel:
                    filtres[col_filtre] = sel

            # Tri
            colonnes_tri = ["Aucun"] + list(df.select_dtypes(include=[np.number, 'datetime']).columns)
            col_tri = st.selectbox("Trier par", options=colonnes_tri)
            sens_tri = st.radio("Sens du tri", options=["Croissant", "Décroissant"])

            # Nombre maximum de slides
            max_slides = st.slider("Nombre max de slides", min_value=1, max_value=len(df), value=len(df))

        # Appliquer filtres et tri
        df_filtre = df.copy()
        for c, vals in filtres.items():
            df_filtre = df_filtre[df_filtre[c].isin(vals)]
        if col_tri != "Aucun":
            df_filtre = df_filtre.sort_values(
                by=col_tri, ascending=(sens_tri == "Croissant")
            )
        df_filtre = df_filtre.head(max_slides)

        st.dataframe(df_filtre)

        # Mapping des placeholders
        placeholders_map = {}
        with st.expander('🔄 Mapping placeholders'):
            for ph, default in PLACEHOLDERS_DEFAULT.items():
                idx = list(df_filtre.columns).index(default) + 1 if default in df_filtre.columns else 0
                col_sel = st.selectbox(ph, options=[''] + list(df_filtre.columns), index=idx)
                placeholders_map[ph] = col_sel

        # Génération
        if st.button('🚀 Générer la présentation'):
            buf = generate_presentation(
                fichier_pptx,
                df_filtre,
                placeholders_map,
                img_dir,
                logos_dir,
                valeur_blason
            )
            st.download_button(
                '🔽 Télécharger la présentation',
                data=buf,
                file_name='presentation_reference.pptx',
                mime='application/vnd.openxmlformats-officedocument.presentationml.presentation'
            )
