import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import datetime
import io

# --- CONFIGURATION DE LA PAGE ---
st.set_page_config(page_title="Ponceblanc Business Intelligence", layout="wide", page_icon="📊")

# Style CSS pour des cartes d'indicateurs épurées
st.markdown("""
    <style>
    .stMetric {
        background-color: #ffffff;
        padding: 15px;
        border-radius: 8px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    </style>
""", unsafe_allow_html=True)

# Liste de référence unique pour toute l'application
liste_mois_noms = [
    "Janvier", "Février", "Mars", "Avril", "Mai", "Juin", 
    "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"
]

# --- CHARGEMENT DE LA FEUILLE DE SYNTHÈSE DEPUIS GITHUB ---
@st.cache_data
def load_synthese_data():
    # ⚠️ REMPLACEZ CETTE URL PAR VOTRE LIEN RAW GITHUB REEL
    url_github = "https://github.com/maximeponceblanc-a11y/dashboard-ponceblanc/blob/f17a8cc3e9d64da0c40c1097dd0a42daa82c35b8/dashboard_commercial.py"
    
    try:
        # Pandas est capable de lire un fichier Excel directement depuis une URL web
        df = pd.read_excel(url_github, sheet_name="PONCEBLANC")
    except Exception as e:
        st.error(f"Erreur lors de la récupération du fichier sur GitHub : {e}")
        st.info("Vérifiez que l'URL 'raw.githubusercontent.com' est correcte et que le dépôt est accessible.")
        st.stop()

    df.columns = df.columns.str.strip()

    # CORRECTION DES DATES : Gestion des numéros de série Excel (ex: 46012)
    if 'Dates' in df.columns:
        s_numeric = pd.to_numeric(df['Dates'], errors='coerce')
        df['Dates_Propres'] = pd.to_datetime(s_numeric, unit='D', origin='1899-12-30', errors='coerce')
        df['Dates_Propres'] = df['Dates_Propres'].fillna(pd.to_datetime(df['Dates'], errors='coerce'))
        df = df.dropna(subset=['Dates_Propres'])

    # SÉCURISATION : Utilisation prioritaire des colonnes natives pré-calculées 'Mois devis' et 'Année Devis'
    if 'Mois devis' in df.columns:
        df['Mois_Num'] = pd.to_numeric(df['Mois devis'], errors='coerce').fillna(1).astype(int)
    elif 'Dates_Propres' in df.columns:
        df['Mois_Num'] = df['Dates_Propres'].dt.month.fillna(1).astype(int)
    else:
        df['Mois_Num'] = 1
        
    if 'Année Devis' in df.columns:
        df['Année Devis'] = pd.to_numeric(df['Année Devis'], errors='coerce').fillna(2026).astype(int)
    elif 'Dates_Propres' in df.columns:
        df['Année Devis'] = df['Dates_Propres'].dt.year.fillna(2026).astype(int)
    else:
        df['Année Devis'] = 2026

    # Traduction en Français pour l'affichage des graphiques
    mois_fr = {1: "Janvier", 2: "Février", 3: "Mars", 4: "Avril", 5: "Mai", 6: "Juin", 
               7: "Juillet", 8: "Août", 9: "Septembre", 10: "Octobre", 11: "Novembre", 12: "Décembre"}
    df['Mois_Nom'] = df['Mois_Num'].map(mois_fr).fillna("Janvier")
    
    df['Prix total'] = pd.to_numeric(df['Prix total'], errors='coerce').fillna(0)
    
    return df

try:
    df = load_synthese_data()
except Exception as e:
    st.error(f"Impossible de charger les données : {e}")
    st.stop()


# ========================================================
# LOGIQUE DE CLASSIFICATION ABC (12 MOIS ROULANTS GLISSANTS)
# ========================================================
def calculer_classification_abc(dataframe):
    max_date_all = dataframe['Dates_Propres'].max()
    if pd.isnull(max_date_all):
        dataframe['Catégorie Client ABC'] = "NOUVEAU CLIENT"
        return dataframe
        
    date_limite_12m = max_date_all - pd.Timedelta(days=365)
    
    mask_12m_signe = (dataframe['Dates_Propres'] >= date_limite_12m) & (dataframe['Signé?'].astype(str).str.upper().str.strip() == "O")
    df_12m_signe = dataframe[mask_12m_signe]
    
    ca_par_client = df_12m_signe.groupby('Nom Client')['Prix total'].sum().reset_index()
    ca_par_client = ca_par_client.sort_values(by='Prix total', ascending=False).reset_index(drop=True)
    
    total_ca_signe_12m = ca_par_client['Prix total'].sum()
    
    dict_classification = {}
    if total_ca_signe_12m > 0:
        ca_par_client['CA_Cumule_Pct'] = (ca_par_client['Prix total'].cumsum() / total_ca_signe_12m) * 100
        
        for _, row in ca_par_client.iterrows():
            client = row['Nom Client']
            pct = row['CA_Cumule_Pct']
            
            if pct <= 80.001:
                dict_classification[client] = "GRAND COMPTE"
            elif pct <= 90.001:
                dict_classification[client] = "CLIENT INTERMEDIAIRE"
            else:
                dict_classification[client] = "PETITS CLIENTS"
                
    def attribuer_classe(client_name):
        return dict_classification.get(client_name, "NOUVEAU CLIENT")
        
    dataframe['Catégorie Client ABC'] = dataframe['Nom Client'].apply(attribuer_classe)
    return dataframe

df = calculer_classification_abc(df)


# ========================================================
# ENTÊTE : SYSTEME DE FILTRES TEMPORELS GLOBAUX
# ========================================================
st.title("📊 Ponceblanc Business Intelligence")

with st.container(border=True):
    st.markdown("#### 📅 Filtres Temporels Globaux")
    col_top_annee, col_top_mois = st.columns([4, 6])
    
    with col_top_annee:
        if 'Année Devis' in df.columns:
            annees_possibles = sorted([int(a) for a in df['Année Devis'].unique() if a > 2000], reverse=True)
        else:
            annees_possibles = [2026, 2025]
        
        annees_selectionnees = st.multiselect(
            "Années d'analyse :",
            options=annees_possibles,
            default=annees_possibles
        )
        
    with col_top_mois:
        start_month, end_month = st.select_slider(
            "Intervalle des mois inclus :",
            options=liste_mois_noms,
            value=("Janvier", "Décembre")
        )

mois_debut_num = liste_mois_noms.index(start_month) + 1
mois_fin_num = liste_mois_noms.index(end_month) + 1
mois_selectionnes_bornes = liste_mois_noms[mois_debut_num - 1:mois_fin_num]

mask_temporel = (df['Année Devis'].isin(annees_selectionnees)) & (df['Mois_Num'] >= mois_debut_num) & (df['Mois_Num'] <= mois_fin_num)


# --- BARRE LATÉRALE : FILTRES METIER ET NAVIGATION ---
st.sidebar.header("Filtres Métier & Profils")

selected_abc = st.sidebar.multiselect(
    "Catégorie Client (ABC)",
    options=["GRAND COMPTE", "CLIENT INTERMEDIAIRE", "PETITS CLIENTS", "NOUVEAU CLIENT"],
    default=["GRAND COMPTE", "CLIENT INTERMEDIAIRE", "PETITS CLIENTS", "NOUVEAU CLIENT"]
)

search_client = st.sidebar.text_input("Rechercher un Nom Client")
search_devis = st.sidebar.text_input("Rechercher un N° de Devis")

mask_metier = pd.Series(True, index=df.index)

if selected_abc and "Catégorie Client ABC" in df.columns:
    mask_metier &= df["Catégorie Client ABC"].isin(selected_abc)
if search_client and "Nom Client" in df.columns:
    mask_metier &= df["Nom Client"].astype(str).str.contains(search_client, case=False, na=False)
if search_devis and "DEVIS N°" in df.columns:
    mask_metier &= df["DEVIS N°"].astype(str).str.contains(search_devis, case=False, na=False)

for col_df, label in [("Commercial", "Commercial"), ("Type de produit", "Type de produit"), ("Matière", "Matière")]:
    if col_df in df.columns:
        options = sorted([str(x) for x in df[col_df].dropna().unique()])
        selected = st.sidebar.multiselect(label, options=options)
        if selected:
            mask_metier &= df[col_df].astype(str).isin(selected)

df_filtered_base = df[mask_metier & mask_temporel].copy()
st.sidebar.success(f"Données filtrées : {len(df_filtered_base)} lignes.")

page = st.sidebar.radio("Navigation", ["📈 Performance Commerciale", "💰 Analyse Marge & Prix"])


# --- FONCTION DE CALCUL DES KPIS ---
def extraire_kpis_annee(dataframe, annee_cible):
    df_cible = dataframe[dataframe['Année Devis'] == annee_cible]
    col_id = "DEVIS N°" if "DEVIS N°" in dataframe.columns else dataframe.columns[0]
    
    if df_cible.empty:
        return None
        
    mask_signe = df_cible["Signé?"].astype(str).str.upper().str.strip() == "O"
    
    ca_devis = df_cible["Prix total"].sum()
    ca_signe = df_cible[mask_signe]["Prix total"].sum()
    tx_succes_ca = (ca_signe / ca_devis * 100) if ca_devis > 0 else 0
    
    vol_devis = df_cible[col_id].nunique()
    vol_signe = df_cible[mask_signe][col_id].nunique()
    tx_succes_vol = (vol_signe / vol_devis * 100) if vol_devis > 0 else 0
    
    return {
        "ca_devis": ca_devis, "ca_signe": ca_signe, "tx_ca": tx_succes_ca,
        "vol_devis": vol_devis, "vol_signe": vol_signe, "tx_vol": tx_succes_vol
    }


# ========================================================
# RENDER : VUE PERFORMANCE COMMERCIALE
# ========================================================
if page == "📈 Performance Commerciale":
    st.subheader("Performance Commerciale")
    
    if len(annees_selectionnees) > 0:
        tabs_val = st.tabs([f"Année {annee}" for annee in annees_selectionnees])
        
        for idx, annee_selectionnee in enumerate(annees_selectionnees):
            with tabs_val[idx]:
                st.markdown(f"##### 📌 Indicateurs Clés — {annee_selectionnee} ({start_month} à {end_month})")
                
                annee_prev = annee_selectionnee - 1
                annee_next = annee_selectionnee + 1
                
                kpis_current = extraire_kpis_annee(df_filtered_base, annee_selectionnee)
                kpis_prev = extraire_kpis_annee(df_filtered_base, annee_prev)
                kpis_next = extraire_kpis_annee(df_filtered_base, annee_next)
                
                if kpis_current:
                    def générer_html_delta(val_curr, kpis_ref, cle, label_annee, unite="€", is_pct=False):
                        if not kpis_ref:
                            return f"<span style='color: #94a3b8;'>vs {label_annee} : -</span>"
                        val_ref = kpis_ref[cle]
                        diff = val_curr - val_ref
                        if diff > 0.001:
                            color = "#22c55e"; arrow = "↗"; sign = "+"
                        elif diff < -0.001:
                            color = "#ef4444"; arrow = "↘"; sign = ""
                        else:
                            color = "#64748b"; arrow = "→"; sign = ""
                        if is_pct:
                            val_txt = f"{sign}{diff:.2f} pts".replace('.', ',')
                        else:
                            val_txt = f"{sign}{diff:,.0f} {unite}".replace(',', ' ')
                        return f"<span style='color: {color}; font-weight: 500;'>{arrow} vs {label_annee} : {val_txt}</span>"

                    c1, c2, c3 = st.columns(3)
                    with c1:
                        st.metric(label=f"CA Signé {annee_selectionnee}", value=f"{kpis_current['ca_signe']:,.0f} €".replace(',', ' '))
                        st.markdown(f"{générer_html_delta(kpis_current['ca_signe'], kpis_prev, 'ca_signe', annee_prev)} &nbsp;|&nbsp; {générer_html_delta(kpis_current['ca_signe'], kpis_next, 'ca_signe', annee_next)}", unsafe_allow_html=True)
                    with c2:
                        st.metric(label=f"CA Devisé {annee_selectionnee}", value=f"{kpis_current['ca_devis']:,.0f} €".replace(',', ' '))
                        st.markdown(f"{générer_html_delta(kpis_current['ca_devis'], kpis_prev, 'ca_devis', annee_prev)} &nbsp;|&nbsp; {générer_html_delta(kpis_current['ca_devis'], kpis_next, 'ca_devis', annee_next)}", unsafe_allow_html=True)
                    with c3:
                        st.metric(label="% Succès (€)", value=f"{kpis_current['tx_ca']:.2f} %".replace('.', ','))
                        st.markdown(f"{générer_html_delta(kpis_current['tx_ca'], kpis_prev, 'tx_ca', annee_prev, is_pct=True)} &nbsp;|&nbsp; {générer_html_delta(kpis_current['tx_ca'], kpis_next, 'tx_ca', annee_next, is_pct=True)}", unsafe_allow_html=True)
                    
                    st.write("") 
                    c4, c5, c6 = st.columns(3)
                    with c4:
                        st.metric(label="Nombre de devis signés", value=f"{kpis_current['vol_signe']}")
                        st.markdown(f"{générer_html_delta(kpis_current['vol_signe'], kpis_prev, 'vol_signe', annee_prev, 'signés')} &nbsp;|&nbsp; {générer_html_delta(kpis_current['vol_signe'], kpis_next, 'vol_signe', annee_next, 'signés')}", unsafe_allow_html=True)
                    with c5:
                        st.metric(label="Nombre de devis émis", value=f"{kpis_current['vol_devis']}")
                        st.markdown(f"{générer_html_delta(kpis_current['vol_devis'], kpis_prev, 'vol_devis', annee_prev, 'devis')} &nbsp;|&nbsp; {générer_html_delta(kpis_current['vol_devis'], kpis_next, 'vol_devis', annee_next, 'devis')}", unsafe_allow_html=True)
                    with c6:
                        st.metric(label="% Succès (Volume)", value=f"{kpis_current['tx_vol']:.2f} %".replace('.', ','))
                        st.markdown(f"{générer_html_delta(kpis_current['tx_vol'], kpis_prev, 'tx_vol', annee_prev, is_pct=True)} &nbsp;|&nbsp; {générer_html_delta(kpis_current['tx_vol'], kpis_next, 'tx_vol', annee_next, is_pct=True)}", unsafe_allow_html=True)
                
                st.divider()
                
                df_annee = df_filtered_base[df_filtered_base['Année Devis'] == annee_selectionnee]
                
                if not df_annee.empty:
                    col_gauche, col_droite = st.columns([6.5, 3.5])
                    
                    with col_gauche:
                        st.markdown(f"**Performance Mensuelle Financement ({annee_selectionnee})**")
                        
                        base_ca = df_annee.groupby('Mois_Nom')['Prix total'].sum().reindex(mois_selectionnes_bornes, fill_value=0).reset_index(name='Chiffre d\'affaire')
                        df_annee_signe = df_annee[df_annee["Signé?"].astype(str).str.upper().str.strip() == "O"]
                        signe_ca = df_annee_signe.groupby('Mois_Nom')['Prix total'].sum().reindex(mois_selectionnes_bornes, fill_value=0).reset_index(name='CA signé')
                        
                        df_graph = pd.merge(base_ca, signe_ca, on='Mois_Nom', how='left').fillna(0)
                        df_graph['Taux succès (€)'] = df_graph.apply(lambda row: (row['CA signé'] / row['Chiffre d\'affaire']) if row['Chiffre d\'affaire'] > 0 else 0, axis=1)
                        
                        fig = make_subplots(specs=[[{"secondary_y": True}]])
                        fig.add_trace(go.Bar(x=df_graph['Mois_Nom'], y=df_graph['Chiffre d\'affaire'], name="Chiffre d'affaire", marker_color='#00a4bd'), secondary_y=False)
                        fig.add_trace(go.Bar(x=df_graph['Mois_Nom'], y=df_graph['CA signé'], name="CA signé", marker_color='#4ed2e6'), secondary_y=False)
                        fig.add_trace(go.Scatter(x=df_graph['Mois_Nom'], y=df_graph['Taux succès (€)'], name="Taux succès (€)", mode='lines+markers', line=dict(color='#0b5ca3', width=2)), secondary_y=True)
                        
                        fig.update_layout(barmode='group', xaxis=dict(type='category'), yaxis=dict(title="Montant (€)", showgrid=True), yaxis2=dict(title="Taux succès (€)", tickformat=".2%", range=[0, 1], showgrid=False), legend=dict(orientation="h", y=1.1, x=0), hovermode="x unified", height=380)
                        st.plotly_chart(fig, use_container_width=True)
                        
                        st.write("---") 
                        st.markdown(f"**Analyse Mensuelle des Volumes & Taux de Succès ({annee_selectionnee})**")
                        
                        col_id_devis = "DEVIS N°" if "DEVIS N°" in df_annee.columns else df_annee.columns[0]
                        base_vol = df_annee.groupby('Mois_Nom')[col_id_devis].nunique().reindex(mois_selectionnes_bornes, fill_value=0).reset_index(name='Nb devis')
                        signe_vol = df_annee_signe.groupby('Mois_Nom')[col_id_devis].nunique().reindex(mois_selectionnes_bornes, fill_value=0).reset_index(name='Nb devis signés')
                        
                        df_graph_vol = pd.merge(base_vol, signe_vol, on='Mois_Nom', how='left').fillna(0)
                        df_graph_vol['Taux succès (vol)'] = df_graph_vol.apply(lambda row: (row['Nb devis signés'] / row['Nb devis']) if row['Nb devis'] > 0 else 0, axis=1)
                        
                        fig_vol = make_subplots(specs=[[{"secondary_y": True}]])
                        fig_vol.add_trace(go.Bar(x=df_graph_vol['Mois_Nom'], y=df_graph_vol['Nb devis'], name="Nb devis", marker_color='#ff9f00'), secondary_y=False)
                        fig_vol.add_trace(go.Bar(x=df_graph_vol['Mois_Nom'], y=df_graph_vol['Nb devis signés'], name="Nb devis signés", marker_color='#ffcc66'), secondary_y=False)
                        fig_vol.add_trace(go.Scatter(x=df_graph_vol['Mois_Nom'], y=df_graph_vol['Taux succès (vol)'], name="Taux succès (vol)", mode='lines+markers', line=dict(color='#e65c00', width=2)), secondary_y=True)
                        
                        fig_vol.update_layout(barmode='group', xaxis=dict(type='category'), yaxis=dict(title="Nombre de devis", showgrid=True), yaxis2=dict(title="Taux succès (vol)", tickformat=".2%", range=[0, 1], showgrid=False), legend=dict(orientation="h", y=1.1, x=0), hovermode="x unified", height=380)
                        st.plotly_chart(fig_vol, use_container_width=True)
                        
                    with col_droite:
                        st.markdown(f"**Part du CA par Catégorie ABC ({annee_selectionnee})**")
                        df_pie = df_annee.groupby('Catégorie Client ABC')['Prix total'].sum().reset_index(name='Total CA')
                        colors_map = {"GRAND COMPTE": "#00a4bd", "CLIENT INTERMEDIAIRE": "#4ed2e6", "PETITS CLIENTS": "#ff9f00", "NOUVEAU CLIENT": "#94a3b8"}
                        
                        fig_pie = go.Figure(data=[go.Pie(
                            labels=df_pie['Catégorie Client ABC'], values=df_pie['Total CA'], hole=.4,
                            marker=dict(colors=[colors_map.get(x, "#94a3b8") for x in df_pie['Catégorie Client ABC']]),
                            textinfo='percent', hovertemplate="<b>%{label}</b><br>CA : %{value:,.0f} €<br>Part : %{percent}<extra></extra>"
                        )])
                        fig_pie.update_layout(legend=dict(orientation="h", y=-0.1, x=0), margin=dict(t=10, b=10, l=10, r=10), height=320)
                        st.plotly_chart(fig_pie, use_container_width=True)
                        
                        st.markdown(f"**📋 Liste des dossiers de devis ({annee_selectionnee})**")
                        df_table = pd.DataFrame()
                        df_table["Nom Client"] = df_annee["Nom Client"]
                        df_table["DEVIS N°"] = df_annee["DEVIS N°"] if "DEVIS N°" in df_annee.columns else "N/A"
                        
                        df_table["Statut Devis"] = "Refusé"
                        mask_o = df_annee["Signé?"].astype(str).str.upper().str.strip() == "O"
                        df_table.loc[mask_o, "Statut Devis"] = "Signé"
                        
                        df_table["CA"] = df_annee["Prix total"]
                        df_table = df_table.sort_values(by="CA", ascending=False)
                        
                        df_table_formatted = df_table.copy()
                        df_table_formatted["CA"] = df_table_formatted["CA"].map(lambda x: f"{x:,.0f} €".replace(',', ' '))
                        
                        st.dataframe(df_table_formatted, use_container_width=True, hide_index=True, height=440)
                else:
                    st.info(f"Aucune ligne de détail pour l'année {annee_selectionnee}.")
    else:
        st.info("💡 Sélectionnez au moins une année dans les filtres en haut de page.")


# ========================================================
# RENDER : VUE MARGE & TARIFICATION
# ========================================================
elif page == "💰 Analyse Marge & Prix":
    st.subheader("Analyse de la Marge & Règles de Prix")
    
    if not df_filtered_base.empty:
        df_filtered_base['Prix Unitaire'] = pd.to_numeric(df_filtered_base['Prix Unitaire'], errors='coerce').fillna(0)
        df_filtered_base['Taux Marge'] = pd.to_numeric(df_filtered_base['Taux Marge'], errors='coerce').fillna(0)
        
        min_prix = float(df_filtered_base['Prix Unitaire'].min())
        max_prix = float(df_filtered_base['Prix Unitaire'].max())
        min_marge = float(df_filtered_base['Taux Marge'].min())
        max_marge_val = float(df_filtered_base['Taux Marge'].max())
        
        if min_prix == max_prix: max_prix = min_prix + 1.0
        if min_marge == max_marge_val: max_marge_val = min_marge + 1.0

        st.write("---")
        st.markdown("**🎛️ Curseurs d'analyse de la Marge**")
        
        col_prix, col_marge = st.columns(2)
        with col_prix:
            prix_range = st.slider("Prix unité (€)", min_value=min_prix, max_value=max_prix, value=(min_prix, max_prix), step=1.0)
        with col_marge:
            marge_range = st.slider("Coef marge", min_value=min_marge, max_value=max_marge_val, value=(min_marge, max_marge_val), step=0.01)
            
        df_marge = df_filtered_base[
            (df_filtered_base['Prix Unitaire'] >= prix_range[0]) & (df_filtered_base['Prix Unitaire'] <= prix_range[1]) &
            (df_filtered_base['Taux Marge'] >= marge_range[0]) & (df_filtered_base['Taux Marge'] <= marge_range[1])
        ].copy()
        
        if not df_marge.empty:
            df_marge['Prix total'] = pd.to_numeric(df_marge['Prix total'], errors='coerce').fillna(0)
            df_marge['Nb exemplaires'] = pd.to_numeric(df_marge['Nb exemplaires'], errors='coerce').fillna(0)
            if 'Nombre de références totales' in df_marge.columns:
                df_marge['Nombre de références totales'] = pd.to_numeric(df_marge['Nombre de références totales'], errors='coerce').fillna(0)
            
            def get_statut(x):
                val = str(x).upper().strip()
                if val == 'O': return 'Signé'
                elif val == 'N': return 'Refusé'
                else: return 'En attente'
            df_marge['Statut Devis'] = df_marge['Signé?'].apply(get_statut)
            
            st.write("---")
            col_table, col_pie = st.columns([6, 4])
            
            with col_table:
                st.markdown("**📋 Récapitulatif des Devis**")
                colonnes_mapping = {
                    'Nom Client': 'Nom Client', 'DEVIS N°': 'DEVIS N°', 'Statut Devis': 'Statut Devis',
                    'Nb exemplaires': 'Quantité', 'Prix Unitaire': 'Prix unité', 'Taux Marge': 'Coef marge', 'Prix total': 'CA'
                }
                cols_presentes = [c for c in colonnes_mapping.keys() if c in df_marge.columns]
                df_table = df_marge[cols_presentes].rename(columns=colonnes_mapping)
                
                if 'CA' in df_table.columns:
                    df_table = df_table.sort_values(by='CA', ascending=False)
                    
                df_table_display = df_table.copy()
                if 'CA' in df_table_display.columns:
                    df_table_display['CA'] = df_table_display['CA'].apply(lambda x: f"{x:,.0f} €".replace(',', ' '))
                if 'Prix unité' in df_table_display.columns:
                    df_table_display['Prix unité'] = df_table_display['Prix unité'].apply(lambda x: f"{x:,.2f}".replace('.', ','))
                if 'Coef marge' in df_table_display.columns:
                    df_table_display['Coef marge'] = df_table_display['Coef marge'].apply(lambda x: f"{x:,.2f}".replace('.', ','))
                    
                st.dataframe(df_table_display, use_container_width=True, hide_index=True, height=350)
                
            with col_pie:
                st.markdown("**Statut Devis par Prix total (CA)**")
                df_pie = df_marge.groupby('Statut Devis')['Prix total'].sum().reset_index()
                color_map = {"Signé": "#f59e42", "Refusé": "#4287f5", "En attente": "#94a3b8"}
                
                fig_pie = go.Figure(data=[go.Pie(
                    labels=df_pie['Statut Devis'], values=df_pie['Prix total'], hole=.5, 
                    marker=dict(colors=[color_map.get(x, "#94a3b8") for x in df_pie['Statut Devis']]),
                    textinfo='percent', hovertemplate="<b>%{label}</b><br>CA : %{value:,.0f} €<extra></extra>"
                )])
                fig_pie.update_layout(margin=dict(t=10, b=10, l=10, r=10), height=350, showlegend=True)
                st.plotly_chart(fig_pie, use_container_width=True)

            st.write("---")
            st.markdown("**Évolution du coef de marge**")
            
            if len(annees_selectionnees) > 1:
                group_cols = ['Année Devis']
                x_axis_col = 'Année Devis'
                ca_devise = df_marge.groupby(group_cols)['Prix total'].sum().reset_index(name='CA devisé')
                df_signe = df_marge[df_marge['Statut Devis'] == 'Signé'].copy()
                
                if not df_signe.empty and 'Taux Marge' in df_signe.columns:
                    df_signe['Marge_Ponderee_Temp'] = df_signe['Taux Marge'] * df_signe['Prix total']
                    agg_signe = df_signe.groupby(group_cols).agg(CA_signe=('Prix total', 'sum'), Somme_Marge_Pond=('Marge_Ponderee_Temp', 'sum')).reset_index()
                    agg_signe['Coef marge'] = agg_signe['Somme_Marge_Pond'] / agg_signe['CA_signe']
                else:
                    agg_signe = pd.DataFrame(columns=group_cols + ['CA_signe', 'Coef marge'])
                df_combo = pd.merge(ca_devise, agg_signe, on=group_cols, how='left').fillna(0)
                df_combo = df_combo.sort_values('Année Devis')
            else:
                x_axis_col = 'Mois_Nom'
                ca_devise = df_marge.groupby('Mois_Nom')['Prix total'].sum().reindex(mois_selectionnes_bornes, fill_value=0).reset_index(name='CA devisé')
                df_signe = df_marge[df_marge['Statut Devis'] == 'Signé'].copy()
                
                if not df_signe.empty and 'Taux Marge' in df_signe.columns:
                    df_signe['Marge_Ponderee_Temp'] = df_signe['Taux Marge'] * df_signe['Prix total']
                    agg_signe = df_signe.groupby('Mois_Nom').agg(CA_signe=('Prix total', 'sum'), Somme_Marge_Pond=('Marge_Ponderee_Temp', 'sum')).reindex(mois_selectionnes_bornes, fill_value=0).reset_index()
                    agg_signe['Coef marge'] = agg_signe.apply(lambda row: row['Somme_Marge_Pond'] / row['CA_signe'] if row['CA_signe'] > 0 else 0, axis=1)
                else:
                    agg_signe = pd.DataFrame(columns=['Mois_Nom', 'CA_signe', 'Coef marge']).reindex(mois_selectionnes_bornes, fill_value=0).reset_index()
                df_combo = pd.merge(ca_devise, agg_signe[['Mois_Nom', 'CA_signe', 'Coef marge']], on='Mois_Nom', how='left').fillna(0)
            
            df_combo[x_axis_col] = df_combo[x_axis_col].astype(str)
            fig_combo = make_subplots(specs=[[{"secondary_y": True}]])
            
            fig_combo.add_trace(go.Bar(x=df_combo[x_axis_col], y=df_combo['CA devisé'], name="CA devisé", marker_color='#cccccc'), secondary_y=False)
            fig_combo.add_trace(go.Bar(x=df_combo[x_axis_col], y=df_combo['CA_signe'], name="CA signé", marker_color='#4287f5'), secondary_y=False)
            
            max_marge_graph = df_combo['Coef marge'].max()
            echelle_marge_max = max_marge_graph * 1.2 if max_marge_graph > 0 else 2 
            
            fig_combo.add_trace(go.Scatter(x=df_combo[x_axis_col], y=df_combo['Coef marge'], name="Coef marge", mode='lines+markers', line=dict(color='#b370d6', width=2), marker=dict(size=8)), secondary_y=True)
            fig_combo.update_layout(barmode='group', xaxis=dict(type='category'), yaxis=dict(title="CA (€)", showgrid=True), yaxis2=dict(title="Coef marge", range=[0, echelle_marge_max], showgrid=False), legend=dict(orientation="h", y=1.1, x=0), hovermode="x unified", height=400)
            st.plotly_chart(fig_combo, use_container_width=True)

            st.write("---")
            st.markdown("**🔬 Modèles de tarification & Corrélations (Taille des bulles = CA global)**")
            color_map_bubble = {"Signé": "#f59e42", "Refusé": "#4287f5", "En attente": "#94a3b8"}
            df_marge['Taille_Bulle_CA'] = df_marge['Prix total'].apply(lambda x: max(x, 1))

            col_g1, col_g2 = st.columns(2)
            with col_g1:
                st.markdown("##### 1. Coef Marge par Nombre de Références")
                if 'Nombre de références totales' in df_marge.columns:
                    fig_g1 = px.scatter(df_marge, x="Nombre de références totales", y="Taux Marge", size="Taille_Bulle_CA", color="Statut Devis", color_discrete_map=color_map_bubble, size_max=30, hover_name="Nom Client", hover_data=["DEVIS N°", "Prix total"])
                    fig_g1.update_layout(margin=dict(t=10, b=10), height=350, xaxis_title="Nombre total de références", yaxis_title="Coef marge")
                    st.plotly_chart(fig_g1, use_container_width=True)
                else:
                    st.info("La colonne 'Nombre de références totales' est absente.")

            with col_g2:
                st.markdown("##### 2. Coef Marge par Quantité (Nb exemplaires)")
                fig_g2 = px.scatter(df_marge, x="Nb exemplaires", y="Taux Marge", size="Taille_Bulle_CA", color="Statut Devis", color_discrete_map=color_map_bubble, size_max=30, hover_name="Nom Client", hover_data=["DEVIS N°", "Prix total"])
                fig_g2.update_layout(margin=dict(t=10, b=10), height=350, xaxis_title="Quantité devisée", yaxis_title="Coef marge")
                st.plotly_chart(fig_g2, use_container_width=True)

            col_g3, col_g4 = st.columns(2)
            with col_g3:
                st.markdown("##### 3. Prix Unitaire par Nombre de Références")
                if 'Nombre de références totales' in df_marge.columns:
                    fig_g3 = px.scatter(df_marge, x="Nombre de références totales", y="Prix Unitaire", size="Taille_Bulle_CA", color="Statut Devis", color_discrete_map=color_map_bubble, size_max=30, hover_name="Nom Client", hover_data=["DEVIS N°", "Prix total"])
                    fig_g3.update_layout(margin=dict(t=10, b=10), height=350, xaxis_title="Nombre total de références", yaxis_title="Prix unitaire (€)")
                    st.plotly_chart(fig_g3, use_container_width=True)
                else:
                    st.info("La colonne 'Nombre de références totales' est absente.")

            with col_g4:
                st.markdown("##### 4. Prix Unitaire par Quantité (Nb exemplaires)")
                fig_g4 = px.scatter(df_marge, x="Nb exemplaires", y="Prix Unitaire", size="Taille_Bulle_CA", color="Statut Devis", color_discrete_map=color_map_bubble, size_max=30, hover_name="Nom Client", hover_data=["DEVIS N°", "Prix total"])
                fig_g4.update_layout(margin=dict(t=10, b=10), height=350, xaxis_title="Quantité devisée", yaxis_title="Prix unitaire (€)")
                st.plotly_chart(fig_g4, use_container_width=True)
        else:
            st.warning("⚠️ Les filtres sélectionnés (Prix / Marge) masquent l'intégralité des données.")
    else:
        st.info("Aucune donnée disponible pour les critères temporels sélectionnés en haut.")