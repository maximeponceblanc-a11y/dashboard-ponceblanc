import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import datetime
import io
import requests

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

# --- CHARGEMENT DE LA FEUILLE DE SYNTHÈSE (MÉTHODE LOCALE) ---
@st.cache_data
# --- CHARGEMENT DE LA FEUILLE DE SYNTHÈSE DEPUIS GITHUB ---
@st.cache_data
def load_synthese_data():
    # URL du fichier brut (Raw) sur votre dépôt GitHub
    url_github = "https://raw.githubusercontent.com/maximeponceblanc-a11y/commercial/69a330105abe9d3395ebc621d0a64444f17adf28/Query_tableau_devis.xlsx"
    
    try:
        # Téléchargement du fichier Excel depuis GitHub
        response = requests.get(url_github)
        response.raise_for_status()  # Bloque le script et lève une erreur si le lien est cassé
        
        # Lecture du fichier Excel directement depuis la mémoire (flux binaire)
        df = pd.read_excel(io.BytesIO(response.content), sheet_name="PONCEBLANC")
    except Exception as e:
        st.error(f"Erreur lors de la récupération du fichier sur GitHub : {e}")
        st.stop()

    # Nettoyage des espaces cachés dans les noms de colonnes
    df.columns = df.columns.str.strip()

    # --- Gestion et conversion des Dates (Format Excel / Standard) ---
    if 'Dates' in df.columns:
        s_numeric = pd.to_numeric(df['Dates'], errors='coerce')
        df['Dates_Propres'] = pd.to_datetime(s_numeric, unit='D', origin='1899-12-30', errors='coerce')
        df['Dates_Propres'] = df['Dates_Propres'].fillna(pd.to_datetime(df['Dates'], errors='coerce'))
        df = df.dropna(subset=['Dates_Propres'])

    # --- Extraction et normalisation du Numéro de Mois ---
    if 'Mois devis' in df.columns:
        df['Mois_Num'] = pd.to_numeric(df['Mois devis'], errors='coerce').fillna(1).astype(int)
    elif 'Dates_Propres' in df.columns:
        df['Mois_Num'] = df['Dates_Probes'].dt.month.fillna(1).astype(int)
    else:
        df['Mois_Num'] = 1
        
    # --- Extraction et normalisation de l'Année ---
    if 'Année Devis' in df.columns:
        df['Année Devis'] = pd.to_numeric(df['Année Devis'], errors='coerce').fillna(2026).astype(int)
    elif 'Dates_Propres' in df.columns:
        df['Année Devis'] = df['Dates_Propres'].dt.year.fillna(2026).astype(int)
    else:
        df['Année Devis'] = 2026

    # --- Traduction des numéros de mois en texte en Français ---
    mois_fr = {
        1: "Janvier", 2: "Février", 3: "Mars", 4: "Avril", 5: "Mai", 6: "Juin", 
        7: "Juillet", 8: "Août", 9: "Septembre", 10: "Octobre", 11: "Novembre", 12: "Décembre"
    }
    df['Mois_Nom'] = df['Mois_Num'].map(mois_fr).fillna("Janvier")
    
    # --- Sécurisation de la colonne de Prix ---
    df['Prix total'] = pd.to_numeric(df['Prix total'], errors='coerce').fillna(0)
    
    return df
try:
    df = load_synthese_data()
except Exception as e:
    st.stop()


# ========================================================
# LOGIQUE DE CLASSIFICATION ABC PAR ANNÉE CALENDAIRE
# ========================================================
# Règles :
#   Pour chaque année N, on regarde si le client a AU MOINS UNE commande signée
#   dans les années STRICTEMENT antérieures à N (tout l'historique disponible).
#
#   → SI le client n'a jamais commandé avant l'année N : NOUVEAU CLIENT
#     (peu importe ce qu'il fait cette année N)
#
#   → SI le client a déjà commandé avant l'année N :
#       Classement ABC basé sur le CA signé de l'année N uniquement :
#       - GRAND COMPTE          : top 80 % du CA signé de l'année
#       - CLIENT INTERMEDIAIRE  : tranche 80–90 %
#       - PETITS CLIENTS        : tranche 90–100 % + clients connus sans commande cette année

def calculer_classification_abc_par_annee(dataframe):
    if dataframe.empty or 'Année Devis' not in dataframe.columns:
        dataframe['Catégorie Client ABC'] = "NOUVEAU CLIENT"
        return dataframe

    mask_signe_global = dataframe["Signé?"].astype(str).str.upper().str.strip() == "O"
    annees_triees = sorted(dataframe['Année Devis'].unique())

    # Accumulation des clients ayant au moins une commande signée,
    # année après année (on ne regarde que les années AVANT l'année courante)
    clients_connus_avant = {}       # annee -> frozenset de clients ayant déjà commandé
    clients_cumul = set()
    for annee in annees_triees:
        clients_connus_avant[annee] = frozenset(clients_cumul)
        # Clients qui ont commandé cette année → connus pour les années suivantes
        nouveaux = set(
            dataframe.loc[
                (dataframe['Année Devis'] == annee) & mask_signe_global,
                'Nom Client'
            ].dropna().unique()
        )
        clients_cumul |= nouveaux

    # Classement ABC par année
    categories = pd.Series("NOUVEAU CLIENT", index=dataframe.index)

    for annee in annees_triees:
        mask_annee = dataframe['Année Devis'] == annee
        df_annee = dataframe[mask_annee]
        connus = clients_connus_avant[annee]  # clients ayant commandé AVANT cette année

        # Clients connus cette année (ont commandé avant) → on les classe en ABC
        df_annee_signe = df_annee[mask_signe_global.reindex(df_annee.index, fill_value=False)]
        ca_par_client = (
            df_annee_signe.groupby('Nom Client')['Prix total']
            .sum()
            .sort_values(ascending=False)
            .reset_index()
        )
        # On ne conserve pour le classement ABC que les clients déjà connus
        ca_par_client_connus = ca_par_client[ca_par_client['Nom Client'].isin(connus)]

        total_ca = ca_par_client_connus['Prix total'].sum()
        dict_abc = {}

        if total_ca > 0:
            ca_par_client_connus = ca_par_client_connus.copy()
            ca_par_client_connus['CA_Cumule_Pct'] = (
                ca_par_client_connus['Prix total'].cumsum() / total_ca * 100
            )
            for _, row in ca_par_client_connus.iterrows():
                client = row['Nom Client']
                pct = row['CA_Cumule_Pct']
                if pct <= 80.001:
                    dict_abc[client] = "GRAND COMPTE"
                elif pct <= 90.001:
                    dict_abc[client] = "CLIENT INTERMEDIAIRE"
                else:
                    dict_abc[client] = "PETITS CLIENTS"

        # Attribution pour chaque ligne de cette année
        for idx, row in df_annee.iterrows():
            client = row['Nom Client']
            if client not in connus:
                # Jamais commandé avant cette année → NOUVEAU CLIENT
                categories.at[idx] = "NOUVEAU CLIENT"
            elif client in dict_abc:
                # Client connu + commande cette année → classement ABC
                categories.at[idx] = dict_abc[client]
            else:
                # Client connu mais sans commande cette année → PETITS CLIENTS
                categories.at[idx] = "PETITS CLIENTS"

    dataframe = dataframe.copy()
    dataframe['Catégorie Client ABC'] = categories
    return dataframe

df = calculer_classification_abc_par_annee(df)


# ========================================================
# ENTÊTE DU TABLEAU DE BORD
# ========================================================
st.title("📊 Tableau de bord commercial")


# ========================================================
# BARRE LATÉRALE : SYSTÈME DE FILTRES GLOBAUX
# ========================================================
st.sidebar.header("🎛️ Filtres & Paramètres")

# Initialisation du masque global
mask_metier = pd.Series(True, index=df.index)

# --- 1. FILTRES TEMPORELS GLOBAUX ---
st.sidebar.markdown("### 📅 Filtres Temporels Globaux")

vue_mode = st.sidebar.radio(
    "Maille d'analyse :",
    options=["📅 Mensuelle", "📆 Annuelle"],
    horizontal=False,
    help="Mensuelle : évolution mois par mois sur une année. Annuelle : comparaison inter-années."
)
vue_annuelle = (vue_mode == "📆 Annuelle")

if 'Année Devis' in df.columns:
    annees_possibles = sorted([int(a) for a in df['Année Devis'].unique() if a > 2000], reverse=True)
else:
    annees_possibles = [2026, 2025]

annees_selectionnees = st.sidebar.multiselect(
    "Années d'analyse :",
    options=annees_possibles,
    default=annees_possibles
)

if not vue_annuelle:
    start_month, end_month = st.sidebar.select_slider(
        "Intervalle des mois inclus :",
        options=liste_mois_noms,
        value=("Janvier", "Décembre")
    )
else:
    st.sidebar.markdown("<span style='color:#64748b;font-size:0.9em;'>En vue annuelle, tous les mois sont inclus.</span>", unsafe_allow_html=True)
    start_month, end_month = "Janvier", "Décembre"

mois_debut_num = liste_mois_noms.index(start_month) + 1
mois_fin_num = liste_mois_noms.index(end_month) + 1
mois_selectionnes_bornes = liste_mois_noms[mois_debut_num - 1:mois_fin_num]

mask_temporel = (df['Année Devis'].isin(annees_selectionnees)) & (df['Mois_Num'] >= mois_debut_num) & (df['Mois_Num'] <= mois_fin_num)


# --- 2. COMMERCIAL ---
st.sidebar.markdown("### 👤 Commercial")

if "Commercial" in df.columns:
    options_comm = sorted([str(x) for x in df["Commercial"].dropna().unique()])
    selected_commercial = st.sidebar.multiselect("Sélection du commercial", options=options_comm)
    if selected_commercial:
        mask_metier &= df["Commercial"].astype(str).isin(selected_commercial)
else:
    selected_commercial = []


# --- 3. DOSSIER ---
st.sidebar.markdown("### 📁 Dossier")

selected_abc = st.sidebar.multiselect(
    "Catégorie Client (ABC)",
    options=["GRAND COMPTE", "CLIENT INTERMEDIAIRE", "PETITS CLIENTS", "NOUVEAU CLIENT"],
    default=["GRAND COMPTE", "CLIENT INTERMEDIAIRE", "PETITS CLIENTS", "NOUVEAU CLIENT"]
)

search_client = st.sidebar.text_input("Rechercher un Nom Client")
search_devis = st.sidebar.text_input("Rechercher un N° de Devis")

if selected_abc and "Catégorie Client ABC" in df.columns:
    mask_metier &= df["Catégorie Client ABC"].isin(selected_abc)
if search_client and "Nom Client" in df.columns:
    mask_metier &= df["Nom Client"].astype(str).str.contains(search_client, case=False, na=False)
if search_devis and "DEVIS N°" in df.columns:
    mask_metier &= df["DEVIS N°"].astype(str).str.contains(search_devis, case=False, na=False)

for col_df, label in [("Type de produit", "Type de produit"), ("Matière", "Matière")]:
    if col_df in df.columns:
        options = sorted([str(x) for x in df[col_df].dropna().unique()])
        selected = st.sidebar.multiselect(label, options=options)
        if selected:
            mask_metier &= df[col_df].astype(str).isin(selected)


# Application des filtres sur le dataframe global
df_filtered_base = df[mask_metier & mask_temporel].copy()
st.sidebar.success(f"Données filtrées : {len(df_filtered_base)} lignes.")


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
# RENDER : VUE UNIQUE PERFORMANCE COMMERCIALE
# ========================================================
st.subheader("Performance Commerciale")

# -------------------------------------------------------
# VUE ANNUELLE : comparaison inter-années
# -------------------------------------------------------
if vue_annuelle and len(annees_selectionnees) > 0:
    annees_triees = sorted(annees_selectionnees)

    # KPIs synthétiques par année (pour les graphiques)
    rows = []
    for a in annees_triees:
        k = extraire_kpis_annee(df_filtered_base, a)
        if k:
            rows.append({"Année": str(a), **k})
    df_kpis_all = pd.DataFrame(rows)

    # --- Année de référence = année la plus récente sélectionnée ---
    annee_courante = max(annees_triees)
    annee_precedente = annee_courante - 1
    annee_en_cours_label = " — Année en cours" if annee_courante == datetime.datetime.now().year else ""

    kpis_courant = extraire_kpis_annee(df_filtered_base, annee_courante)
    kpis_precedent = extraire_kpis_annee(df[mask_metier], annee_precedente)

    def delta_str(val, ref, unite="€", is_pct=False):
        if ref is None:
            return None
        d = val - ref
        if is_pct:
            return f"{'+' if d >= 0 else ''}{d:.2f} pts".replace('.', ',')
        return f"{'+' if d >= 0 else ''}{d:,.0f} {unite}".replace(',', ' ')

    ref = kpis_precedent  # None si pas de données N-1

    st.markdown(
        f"##### 📌 Indicateurs Clés **{annee_courante}**"
        f"<span style='font-size:0.85em; font-weight:400; color:#64748b;'>"
        f"{annee_en_cours_label}"
        f"{'  ·  Δ vs ' + str(annee_precedente) if ref else ''}</span>",
        unsafe_allow_html=True
    )

    if kpis_courant:
        # --- Ligne 1 : indicateurs financiers (€) ---
        c1, c2, c3 = st.columns(3)
        with c1:
            st.metric(
                label=f"CA Commandes {annee_courante}",
                value=f"{kpis_courant['ca_signe']:,.0f} €".replace(',', ' '),
                delta=delta_str(kpis_courant['ca_signe'], ref['ca_signe'] if ref else None)
            )
        with c2:
            st.metric(
                label=f"CA Devisé {annee_courante}",
                value=f"{kpis_courant['ca_devis']:,.0f} €".replace(',', ' '),
                delta=delta_str(kpis_courant['ca_devis'], ref['ca_devis'] if ref else None)
            )
        with c3:
            st.metric(
                label="% Succès (€)",
                value=f"{kpis_courant['tx_ca']:.2f} %".replace('.', ','),
                delta=delta_str(kpis_courant['tx_ca'], ref['tx_ca'] if ref else None, is_pct=True)
            )

        st.write("")

        # --- Ligne 2 : indicateurs volumes ---
        c4, c5, c6 = st.columns(3)
        with c4:
            st.metric(
                label="Nombre de commandes",
                value=f"{int(kpis_courant['vol_signe'])}",
                delta=delta_str(kpis_courant['vol_signe'], ref['vol_signe'] if ref else None, unite="")
            )
        with c5:
            st.metric(
                label="Nombre de devis émis",
                value=f"{int(kpis_courant['vol_devis'])}",
                delta=delta_str(kpis_courant['vol_devis'], ref['vol_devis'] if ref else None, unite="")
            )
        with c6:
            st.metric(
                label="% Succès (Volume)",
                value=f"{kpis_courant['tx_vol']:.2f} %".replace('.', ','),
                delta=delta_str(kpis_courant['tx_vol'], ref['tx_vol'] if ref else None, is_pct=True)
            )

        if ref:
            st.caption(f"↕ Deltas calculés par rapport à {annee_precedente}")

    st.divider()

    col_g, col_d = st.columns([6, 4])

    with col_g:
        # Graphique CA Devisé / Commandes par année + taux succès
        st.markdown("**Performance Annuelle — Financement**")
        if not df_kpis_all.empty:
            fig_yr = make_subplots(specs=[[{"secondary_y": True}]])
            fig_yr.add_trace(go.Bar(x=df_kpis_all['Année'], y=df_kpis_all['ca_devis'], name="CA Devisé", marker_color='#00a4bd'), secondary_y=False)
            fig_yr.add_trace(go.Bar(x=df_kpis_all['Année'], y=df_kpis_all['ca_signe'], name="CA Commandes", marker_color='#4ed2e6'), secondary_y=False)
            fig_yr.add_trace(go.Scatter(x=df_kpis_all['Année'], y=df_kpis_all['tx_ca'] / 100, name="Taux succès (€)", mode='lines+markers+text',
                text=[f"{v:.1f}%".replace('.', ',') for v in df_kpis_all['tx_ca']], textposition='top center',
                line=dict(color='#0b5ca3', width=2)), secondary_y=True)
            fig_yr.update_layout(barmode='group', xaxis=dict(type='category'), yaxis=dict(title="Montant (€)"), yaxis2=dict(title="Taux succès", tickformat=".0%", range=[0, 1.2], showgrid=False),
                legend=dict(orientation="h", y=1.1), hovermode="x unified", height=380)
            st.plotly_chart(fig_yr, use_container_width=True)

        st.write("---")
        st.markdown("**Analyse Annuelle — Volumes & Taux de Succès**")
        if not df_kpis_all.empty:
            fig_vol_yr = make_subplots(specs=[[{"secondary_y": True}]])
            fig_vol_yr.add_trace(go.Bar(x=df_kpis_all['Année'], y=df_kpis_all['vol_devis'], name="Nb devis émis", marker_color='#ff9f00'), secondary_y=False)
            fig_vol_yr.add_trace(go.Bar(x=df_kpis_all['Année'], y=df_kpis_all['vol_signe'], name="Nb commandes", marker_color='#ffcc66'), secondary_y=False)
            fig_vol_yr.add_trace(go.Scatter(x=df_kpis_all['Année'], y=df_kpis_all['tx_vol'] / 100, name="Taux succès (vol)", mode='lines+markers+text',
                text=[f"{v:.1f}%".replace('.', ',') for v in df_kpis_all['tx_vol']], textposition='top center',
                line=dict(color='#e65c00', width=2)), secondary_y=True)
            fig_vol_yr.update_layout(barmode='group', xaxis=dict(type='category'), yaxis=dict(title="Nombre de devis"), yaxis2=dict(title="Taux succès", tickformat=".0%", range=[0, 1.2], showgrid=False),
                legend=dict(orientation="h", y=1.1), hovermode="x unified", height=380)
            st.plotly_chart(fig_vol_yr, use_container_width=True)

    with col_d:
        # Donut ABC — CA SIGNÉ uniquement
        st.markdown("**Part du CA signé par Catégorie ABC (toutes années)**")
        colors_map_abc = {"GRAND COMPTE": "#00a4bd", "CLIENT INTERMEDIAIRE": "#4ed2e6", "PETITS CLIENTS": "#ff9f00", "NOUVEAU CLIENT": "#94a3b8"}
        df_pie_all = df_filtered_base[df_filtered_base["Signé?"].astype(str).str.upper().str.strip() == "O"].groupby('Catégorie Client ABC')['Prix total'].sum().reset_index(name='Total CA')
        fig_pie_all = go.Figure(data=[go.Pie(
            labels=df_pie_all['Catégorie Client ABC'], values=df_pie_all['Total CA'], hole=.4,
            marker=dict(colors=[colors_map_abc.get(x, "#94a3b8") for x in df_pie_all['Catégorie Client ABC']]),
            textinfo='percent', hovertemplate="<b>%{label}</b><br>CA : %{value:,.0f} €<br>Part : %{percent}<extra></extra>"
        )])
        fig_pie_all.update_layout(legend=dict(orientation="h", y=-0.1, x=0), margin=dict(t=10, b=10, l=10, r=10), height=320)
        st.plotly_chart(fig_pie_all, use_container_width=True)

        # Tableau récap KPIs par année
        st.markdown("**📋 Récapitulatif KPIs par Année**")
        if not df_kpis_all.empty:
            df_recap = df_kpis_all.copy()
            df_recap['CA Commandes'] = df_recap['ca_signe'].map(lambda x: f"{x:,.0f} €".replace(',', ' '))
            df_recap['CA Devisé'] = df_recap['ca_devis'].map(lambda x: f"{x:,.0f} €".replace(',', ' '))
            df_recap['Tx Succès (€)'] = df_recap['tx_ca'].map(lambda x: f"{x:.2f} %".replace('.', ','))
            df_recap['Commandes'] = df_recap['vol_signe'].astype(int)
            df_recap['Devis émis'] = df_recap['vol_devis'].astype(int)
            df_recap['Tx Succès (vol)'] = df_recap['tx_vol'].map(lambda x: f"{x:.2f} %".replace('.', ','))
            st.dataframe(df_recap[['Année', 'CA Commandes', 'CA Devisé', 'Tx Succès (€)', 'Commandes', 'Devis émis', 'Tx Succès (vol)']], use_container_width=True, hide_index=True)

# -------------------------------------------------------
# VUE MENSUELLE (comportement original)
# -------------------------------------------------------
elif not vue_annuelle and len(annees_selectionnees) > 0:
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
                    st.metric(label=f"CA Commandes {annee_selectionnee}", value=f"{kpis_current['ca_signe']:,.0f} €".replace(',', ' '))
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
                    st.metric(label="Nombre de commandes", value=f"{kpis_current['vol_signe']}")
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
                    st.markdown(f"**Performance Mensuelle CA ({annee_selectionnee})**")
                    
                    base_ca = df_annee.groupby('Mois_Nom')['Prix total'].sum().reindex(mois_selectionnes_bornes, fill_value=0).reset_index(name='Chiffre d\'affaire')
                    df_annee_signe = df_annee[df_annee["Signé?"].astype(str).str.upper().str.strip() == "O"]
                    signe_ca = df_annee_signe.groupby('Mois_Nom')['Prix total'].sum().reindex(mois_selectionnes_bornes, fill_value=0).reset_index(name='CA commandes')
                    
                    df_graph = pd.merge(base_ca, signe_ca, on='Mois_Nom', how='left').fillna(0)
                    df_graph['Taux succès (€)'] = df_graph.apply(lambda row: (row['CA commandes'] / row['Chiffre d\'affaire']) if row['Chiffre d\'affaire'] > 0 else 0, axis=1)
                    
                    fig = make_subplots(specs=[[{"secondary_y": True}]])
                    fig.add_trace(go.Bar(x=df_graph['Mois_Nom'], y=df_graph['Chiffre d\'affaire'], name="Chiffre d'affaire", marker_color='#00a4bd'), secondary_y=False)
                    fig.add_trace(go.Bar(x=df_graph['Mois_Nom'], y=df_graph['CA commandes'], name="CA commandes", marker_color='#4ed2e6'), secondary_y=False)
                    fig.add_trace(go.Scatter(x=df_graph['Mois_Nom'], y=df_graph['Taux succès (€)'], name="Taux succès (€)", mode='lines+markers', line=dict(color='#0b5ca3', width=2)), secondary_y=True)
                    
                    fig.update_layout(barmode='group', xaxis=dict(type='category'), yaxis=dict(title="Montant (€)", showgrid=True), yaxis2=dict(title="Taux succès (€)", tickformat=".2%", range=[0, 1], showgrid=False), legend=dict(orientation="h", y=1.1, x=0), hovermode="x unified", height=380)
                    st.plotly_chart(fig, use_container_width=True)
                    
                    st.write("---") 
                    st.markdown(f"**Analyse Mensuelle des Volumes & Taux de Succès ({annee_selectionnee})**")
                    
                    col_id_devis = "DEVIS N°" if "DEVIS N°" in df_annee.columns else df_annee.columns[0]
                    base_vol = df_annee.groupby('Mois_Nom')[col_id_devis].nunique().reindex(mois_selectionnes_bornes, fill_value=0).reset_index(name='Nb devis')
                    signe_vol = df_annee_signe.groupby('Mois_Nom')[col_id_devis].nunique().reindex(mois_selectionnes_bornes, fill_value=0).reset_index(name='Nb commandes')
                    
                    df_graph_vol = pd.merge(base_vol, signe_vol, on='Mois_Nom', how='left').fillna(0)
                    df_graph_vol['Taux succès (vol)'] = df_graph_vol.apply(lambda row: (row['Nb commandes'] / row['Nb devis']) if row['Nb devis'] > 0 else 0, axis=1)
                    
                    fig_vol = make_subplots(specs=[[{"secondary_y": True}]])
                    fig_vol.add_trace(go.Bar(x=df_graph_vol['Mois_Nom'], y=df_graph_vol['Nb devis'], name="Nb devis", marker_color='#ff9f00'), secondary_y=False)
                    fig_vol.add_trace(go.Bar(x=df_graph_vol['Mois_Nom'], y=df_graph_vol['Nb commandes'], name="Nb commandes", marker_color='#ffcc66'), secondary_y=False)
                    fig_vol.add_trace(go.Scatter(x=df_graph_vol['Mois_Nom'], y=df_graph_vol['Taux succès (vol)'], name="Taux succès (vol)", mode='lines+markers', line=dict(color='#e65c00', width=2)), secondary_y=True)
                    
                    fig_vol.update_layout(barmode='group', xaxis=dict(type='category'), yaxis=dict(title="Nombre de devis", showgrid=True), yaxis2=dict(title="Taux succès (vol)", tickformat=".2%", range=[0, 1], showgrid=False), legend=dict(orientation="h", y=1.1, x=0), hovermode="x unified", height=380)
                    st.plotly_chart(fig_vol, use_container_width=True)
                    
                with col_droite:
                    st.markdown(f"**Part du CA signé par Catégorie ABC ({annee_selectionnee})**")
                    colors_map = {"GRAND COMPTE": "#00a4bd", "CLIENT INTERMEDIAIRE": "#4ed2e6", "PETITS CLIENTS": "#ff9f00", "NOUVEAU CLIENT": "#94a3b8"}
                    df_pie = df_annee[df_annee["Signé?"].astype(str).str.upper().str.strip() == "O"].groupby('Catégorie Client ABC')['Prix total'].sum().reset_index(name='Total CA')
                    
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
                    df_table.loc[mask_o, "Statut Devis" ] = "Commande"
                    
                    df_table["CA"] = df_annee["Prix total"]
                    df_table = df_table.sort_values(by="CA", ascending=False)
                    
                    df_table_formatted = df_table.copy()
                    df_table_formatted["CA"] = df_table_formatted["CA"].map(lambda x: f"{x:,.0f} €".replace(',', ' '))
                    
                    st.dataframe(df_table_formatted, use_container_width=True, hide_index=True, height=440)
            else:
                st.info(f"Aucune ligne de détail pour l'année {annee_selectionnee}.")

if len(annees_selectionnees) == 0:
    st.info("💡 Sélectionnez au moins une année dans les filtres de la barre latérale.")


# ========================================================
# SECTION : 3 CAMEMBERTS COMPLÉMENTAIRES (CA SIGNÉ)
# ========================================================
st.divider()
st.subheader("Répartition du CA Commandes")

mask_signe_global = df_filtered_base["Signé?"].astype(str).str.upper().str.strip() == "O"
df_signe = df_filtered_base[mask_signe_global].copy()

# Valeurs manquantes → "Non indiqué"
for col_ni in ["Type de produit", "Matière"]:
    if col_ni in df_signe.columns:
        df_signe[col_ni] = df_signe[col_ni].fillna("Non indiqué").replace("", "Non indiqué")

col_p1, col_p2, col_p3 = st.columns(3)

# --- Camembert 1 : CA commandes par client (top 6 en couleur dans légende) ---
with col_p1:
    if "Nom Client" in df_signe.columns and not df_signe.empty:
        df_c = df_signe.groupby("Nom Client")["Prix total"].sum().sort_values(ascending=False).reset_index()
        df_c.columns = ["Label", "Valeur"]
        top6 = df_c.head(6)["Label"].tolist()
        n = len(df_c)

        # Labels visibles dans la légende : vrais noms pour top 6, chaîne vide pour les autres
        legend_labels = [lbl if lbl in top6 else "" for lbl in df_c["Label"]]
        couleurs = (
            px.colors.qualitative.Set2[:min(6, n)]
            + ["#e2e8f0"] * max(0, n - 6)
        )

        fig_c1 = go.Figure(data=[go.Pie(
            labels=legend_labels,
            values=df_c["Valeur"],
            hole=0.4,
            textinfo="percent",
            customdata=df_c["Label"],
            hovertemplate="<b>%{customdata}</b><br>CA : %{value:,.0f} €<br>Part : %{percent}<extra></extra>",
            marker=dict(colors=couleurs),
        )])
        fig_c1.update_layout(
            title=dict(text="CA commandes par Client <span style='font-size:11px;color:#94a3b8'>(top 6 nommés, survol pour détail)</span>", font=dict(size=13), x=0),
            legend=dict(orientation="h", y=-0.15, x=0),
            margin=dict(t=40, b=10, l=10, r=10),
            height=340,
        )
        st.plotly_chart(fig_c1, use_container_width=True)
    else:
        st.info("Aucune donnée client disponible.")

# --- Camembert 2 : CA commandes par type de produit ---
with col_p2:
    if "Type de produit" in df_signe.columns and not df_signe.empty:
        df_p2 = df_signe.groupby("Type de produit")["Prix total"].sum().sort_values(ascending=False).reset_index()
        df_p2.columns = ["Label", "Valeur"]
        fig_c2 = go.Figure(data=[go.Pie(
            labels=df_p2["Label"],
            values=df_p2["Valeur"],
            hole=0.4,
            textinfo="percent",
            hovertemplate="<b>%{label}</b><br>CA : %{value:,.0f} €<br>Part : %{percent}<extra></extra>",
        )])
        fig_c2.update_layout(
            title=dict(text="CA commandes par Type de produit", font=dict(size=13), x=0),
            legend=dict(orientation="h", y=-0.15, x=0),
            margin=dict(t=40, b=10, l=10, r=10),
            height=340,
        )
        st.plotly_chart(fig_c2, use_container_width=True)
    else:
        st.info("Colonne 'Type de produit' non disponible.")

# --- Camembert 3 : CA commandes par matière ---
with col_p3:
    if "Matière" in df_signe.columns and not df_signe.empty:
        df_p3 = df_signe.groupby("Matière")["Prix total"].sum().sort_values(ascending=False).reset_index()
        df_p3.columns = ["Label", "Valeur"]
        fig_c3 = go.Figure(data=[go.Pie(
            labels=df_p3["Label"],
            values=df_p3["Valeur"],
            hole=0.4,
            textinfo="percent",
            hovertemplate="<b>%{label}</b><br>CA : %{value:,.0f} €<br>Part : %{percent}<extra></extra>",
        )])
        fig_c3.update_layout(
            title=dict(text="CA commandes par Matière", font=dict(size=13), x=0),
            legend=dict(orientation="h", y=-0.15, x=0),
            margin=dict(t=40, b=10, l=10, r=10),
            height=340,
        )
        st.plotly_chart(fig_c3, use_container_width=True)
    else:
        st.info("Colonne 'Matière' non disponible.")


# ========================================================
# SECTION : DÉLAI MOYEN D'OUVERTURE PAR MAILLE TEMPORELLE
# ========================================================
st.divider()
st.subheader("⏱️ Délai Moyen d'Ouverture des Devis")

if 'Délai devis ouverture' in df_filtered_base.columns:
    df_delai = df_filtered_base.copy()
    df_delai['Délai devis ouverture'] = pd.to_numeric(df_delai['Délai devis ouverture'], errors='coerce')
    df_delai = df_delai.dropna(subset=['Délai devis ouverture'])

    if not df_delai.empty:
        if vue_annuelle:
            # Vue annuelle : regroupement par année
            df_delai_groupe = (
                df_delai.groupby('Année Devis')['Délai devis ouverture']
                .mean()
                .reset_index()
            )
            df_delai_groupe.columns = ['Période', 'Délai moyen']
            df_delai_groupe['Période'] = df_delai_groupe['Période'].astype(str)
            x_label = "Année"
        else:
            # Vue mensuelle : regroupement par mois (dans l'intervalle sélectionné, toutes années confondues)
            df_delai_groupe = (
                df_delai.groupby('Mois_Nom')['Délai devis ouverture']
                .mean()
                .reindex(mois_selectionnes_bornes, fill_value=None)
                .reset_index()
            )
            df_delai_groupe.columns = ['Période', 'Délai moyen']
            x_label = "Mois"

        col_delai_graph, col_delai_table = st.columns([6, 4])

        with col_delai_graph:
            fig_delai = go.Figure()
            fig_delai.add_trace(go.Bar(
                x=df_delai_groupe['Période'],
                y=df_delai_groupe['Délai moyen'].round(1),
                marker_color='#6366f1',
                text=df_delai_groupe['Délai moyen'].round(1).apply(
                    lambda v: f"{v:.1f} j".replace('.', ',') if pd.notna(v) else ""
                ),
                textposition='outside',
                hovertemplate="<b>%{x}</b><br>Délai moyen : %{y:.1f} jours<extra></extra>",
            ))
            fig_delai.update_layout(
                xaxis=dict(title=x_label, type='category'),
                yaxis=dict(title="Jours", showgrid=True),
                hovermode="x unified",
                height=360,
                margin=dict(t=20, b=40),
            )
            st.plotly_chart(fig_delai, use_container_width=True)

        with col_delai_table:
            st.markdown("**📋 Détail des délais d'ouverture**")
            cols_detail = ["DEVIS N°", "Nom Client", "Année Devis", "Mois_Nom", "Délai devis ouverture"]
            cols_present = [c for c in cols_detail if c in df_delai.columns]
            df_detail_delai = df_delai[cols_present].copy()
            df_detail_delai = df_detail_delai.rename(columns={
                "DEVIS N°": "N° Devis",
                "Année Devis": "Année",
                "Mois_Nom": "Mois",
                "Délai devis ouverture": "Délai (j)"
            })
            df_detail_delai["Délai (j)"] = df_detail_delai["Délai (j)"].round(1)
            df_detail_delai = df_detail_delai.sort_values("Délai (j)", ascending=False).reset_index(drop=True)
            st.dataframe(df_detail_delai, use_container_width=True, hide_index=True, height=350)
    else:
        st.info("Aucune donnée de délai disponible pour les filtres sélectionnés.")
else:
    st.info("Colonne 'Délai devis ouverture' non trouvée dans les données.")


# ========================================================
# PIED DE PAGE : RÉSUMÉ DES FILTRES ACTIFS
# ========================================================
def construire_resume_filtres():
    parties = []
    maille = "Annuelle" if vue_annuelle else "Mensuelle"
    parties.append(f"Maille : {maille}")
    if annees_selectionnees:
        parties.append(f"Années : {', '.join(str(a) for a in sorted(annees_selectionnees))}")
    if not vue_annuelle:
        if start_month == end_month:
            parties.append(f"Mois : {start_month}")
        else:
            parties.append(f"Mois : {start_month} → {end_month}")
    if selected_commercial:
        parties.append(f"Commercial : {', '.join(selected_commercial)}")
    else:
        parties.append("Commercial : Tous")
    if selected_abc and len(selected_abc) < 4:
        parties.append(f"Catégorie ABC : {', '.join(selected_abc)}")
    else:
        parties.append("Catégorie ABC : Toutes")
    if search_client:
        parties.append(f"Client : « {search_client} »")
    if search_devis:
        parties.append(f"Devis : « {search_devis} »")
    parties.append(f"{len(df_filtered_base)} ligne(s) affichée(s)")
    now = datetime.datetime.now().strftime("%d/%m/%Y à %H:%M")
    parties.append(f"Édité le {now}")
    return "  ·  ".join(parties)

st.markdown("---")
st.markdown(
    f"<p style='font-size:0.75em; color:#94a3b8; text-align:center; margin-top:4px;'>"
    f"🖨️ {construire_resume_filtres()}"
    f"</p>",
    unsafe_allow_html=True
)