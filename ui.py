# File: ui.py

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from config import PLACEHOLDERS_DEFAULT
from presentation_generator import generate_presentation

def run_app():
    st.set_page_config(page_title='Générateur ISOEDRE', layout='wide')
    st.title('🚀 Générateur Document de référence - ISOEDRE')

    # Sidebar : paramètres
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
        fichier_excel = st.file_uploader('📥 Fichier Excel (référence)', type=['xlsx'])
    with col2:
        fichier_pptx = st.file_uploader('📥 Modèle PPTX', type=['pptx'])

    # Buffer pour l’Excel mis à jour
    excel_buf = None

    if fichier_excel and fichier_pptx:
        # Lecture
        df = pd.read_excel(fichier_excel, header=1, dtype=str)

        # Initialisation du compteur de filtres
        if 'n_filters' not in st.session_state:
            st.session_state.n_filters = 1

        # --- Formulaire d’ajout de projet ---
        with st.expander('➕ Ajouter un projet', expanded=False):
            with st.form('form_new_project'):
                st.write('Remplissez les champs pour le nouveau projet :')
                new_data = {}
                for col in df.columns:
                    key = col.lower().strip()
                    if key in ['travaux intérieurs', 'travaux interieurs', 
                               'travaux extérieurs', 'travaux exterieurs']:
                        new_data[col] = st.text_area(col, height=120)
                    else:
                        new_data[col] = st.text_input(col)
                submitted = st.form_submit_button('Ajouter')

            if submitted:
                df.loc[len(df)] = new_data
                st.success('✅ Projet ajouté !')
                # Préparer buffer Excel hors du formulaire
                excel_buf = BytesIO()
                df.to_excel(excel_buf, index=False)
                excel_buf.seek(0)

        # Bouton de téléchargement du nouvel Excel (hors du form)
        if excel_buf:
            st.download_button(
                '🔄 Télécharger Excel mis à jour',
                data=excel_buf,
                file_name='reference_mise_a_jour.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

        # --- Section filtres et tri ---
        st.subheader("🔍 Filtres et tri")
        filtres = {}

        with st.expander("Filtres et options de tri", expanded=True):
            # Pour chaque filtre actif
            for i in range(st.session_state.n_filters):
                col_filtre = st.selectbox(
                    f"Filtre {i+1} — sélectionner une colonne",
                    options=["Aucun"] + list(df.columns),
                    key=f"filter_col_{i}"
                )
                if col_filtre and col_filtre != "Aucun":
                    vals = sorted(df[col_filtre].dropna().unique())
                    sel = st.multiselect(
                        f"Valeurs pour {col_filtre}",
                        options=vals,
                        key=f"filter_vals_{i}"
                    )
                    if sel:
                        filtres[col_filtre] = sel

            # Bouton pour ajouter un filtre supplémentaire
            if st.button("➕ Ajouter un filtre"):
                st.session_state.n_filters += 1

            # Section tri
            st.markdown("---")
            colonnes_tri = ["Aucun"] + list(df.select_dtypes(include=[np.number, 'datetime']).columns)
            col_tri = st.selectbox("Trier par", options=colonnes_tri, key="col_tri")
            sens_tri = st.radio("Sens du tri", options=["Croissant", "Décroissant"], key="sens_tri")

            # Nombre max de slides
            max_slides = st.slider(
                "Nombre max de slides",
                min_value=1,
                max_value=len(df),
                value=len(df),
                key="max_slides"
            )

        # Application des filtres
        df_filtre = df.copy()
        for colonne, valeurs in filtres.items():
            df_filtre = df_filtre[df_filtre[colonne].isin(valeurs)]

        # Application du tri
        if st.session_state.col_tri and st.session_state.col_tri != "Aucun":
            ascending = st.session_state.sens_tri == "Croissant"
            df_filtre = df_filtre.sort_values(by=st.session_state.col_tri, ascending=ascending)

        # Limitation du nombre de slides
        df_filtre = df_filtre.head(st.session_state.max_slides)

        # Affichage du DataFrame filtré
        st.dataframe(df_filtre)

        # --- Mapping des placeholders ---
        placeholders_map = {}
        with st.expander('🔄 Mapping placeholders', expanded=True):
            for ph, default in PLACEHOLDERS_DEFAULT.items():
                idx = (list(df_filtre.columns).index(default) + 1) if default in df_filtre.columns else 0
                col_sel = st.selectbox(ph, options=[''] + list(df_filtre.columns), index=idx, key=f"map_{ph}")
                placeholders_map[ph] = col_sel

        # --- Génération PPTX ---
        if st.button('🚀 Générer PPTX'):
            buf = generate_presentation(
                model_path=fichier_pptx,
                df=df_filtre,
                placeholders_mapping=placeholders_map,
                img_dir=img_dir,
                logos_dir=logos_dir,
                valeur_blason_active=valeur_blason
            )
            st.download_button(
                '🔽 Télécharger PPTX',
                data=buf,
                file_name='presentation_output.pptx',
                mime='application/vnd.openxmlformats-officedocument.presentationml.presentation'
            )
