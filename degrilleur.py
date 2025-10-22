import streamlit as st
import pandas as pd
from io import BytesIO
import datetime
import numpy as np

def create_excel_file(data, fournisseurs):
    """Crée un fichier Excel basé sur les données du formulaire"""
    # Création d'un DataFrame avec la structure du fichier original
    excel_data = []
    
    # En-tête
    excel_data.extend([
        ["", "NOM ET N° DU PROJET :", data.get('nom_projet', ''), "", ""],
        ["", "", "SPECIFICATION TECHNIQUE PARTICULIERE N° :", data.get('stp_numero', ''), ""],
        ["UNITE FONCTIONNELLE", "", "DEGRILLEUR", "", ""],
        ["MATERIEL", "", "Tamis escalier", "", ""],
        ["REPERE PID", "", data.get('repere_pid', 'DGR D 1001 / DGR D 2001'), "", ""],
        ["Rédacteur :", data.get('redacteur', ''), "STATUT :", data.get('statut', ''), ""],
        ["Vérificateur / Approbateur :", data.get('verificateur', ''), "INDICE :", data.get('indice', ''), ""],
        ["", "", "", "", ""],
        ["", "", "demande STEREAU"] + [f"Réponse {fournisseur}" for fournisseur in fournisseurs] + ["Ind."]
    ])
    
    # Section 1 - Spécifications générales
    excel_data.append(["1. SPECIFICATIONS TECHNIQUES GENERALES JOINTES", "", "1 STE SPG ENS 0001"] + 
                     [data.get(f'spec_techniques_{fournisseur}', '') for fournisseur in fournisseurs] + [""])
    
    # Section 2 - Conditions de fonctionnement
    excel_data.extend([
        ["2. CONDITIONS DE FONCTIONNEMENT", "", "", "", ""] + [""] * len(fournisseurs),
        ["Nombre d'équipements", "", data.get('nb_equipements', '')] + 
        [data.get(f'nb_equipements_{fournisseur}', '') for fournisseur in fournisseurs] + [""],
        ["Situation", "", data.get('situation', '')] + 
        [data.get(f'situation_{fournisseur}', '') for fournisseur in fournisseurs] + [""],
        ["Effluent à traiter", "", data.get('effluent', '')] + 
        [data.get(f'effluent_{fournisseur}', '') for fournisseur in fournisseurs] + [""],
        ["Zone ATEX", "", data.get('zone_atex', '')] + 
        [data.get(f'zone_atex_{fournisseur}', '') for fournisseur in fournisseurs] + [""],
        ["Implantation dans l'ouvrage", "", data.get('implantation', '')] + 
        [data.get(f'implantation_{fournisseur}', '') for fournisseur in fournisseurs] + [""]
    ])
    
    # Section 3 - Performances et dimensionnement
    sections_3 = [
        ["3. PERFORMANCES ET DIMENSIONNEMENT REQUIS", "", "", "", ""] + [""] * len(fournisseurs),
        ["Tamis escalier", "", "", "", ""] + [""] * len(fournisseurs),
        ["Maille", "mm", data.get('maille', '')] + 
        [data.get(f'maille_{fournisseur}', '') for fournisseur in fournisseurs] + [""],
        ["Type d'alimentation", "", data.get('type_alimentation', '')] + 
        [data.get(f'type_alimentation_{fournisseur}', '') for fournisseur in fournisseurs] + [""],
        ["Refus de tamis à relever (attendu)", "l/h", data.get('refus_tamis', '')] + 
        [data.get(f'refus_tamis_{fournisseur}', '') for fournisseur in fournisseurs] + [""],
        ["Largeur canal", "m", data.get('largeur_canal', '')] + 
        [data.get(f'largeur_canal_{fournisseur}', '') for fournisseur in fournisseurs] + [""],
        ["Radier canal", "NGF", data.get('radier_canal', '')] + 
        [data.get(f'radier_canal_{fournisseur}', '') for fournisseur in fournisseurs] + [""],
        ["Niveau de la zone de circulation dessus canal", "NGF", data.get('niveau_circulation', '')] + 
        [data.get(f'niveau_circulation_{fournisseur}', '') for fournisseur in fournisseurs] + [""],
        ["NL (AMONT TAMIS) maxi à 1800 m3/h", "NGF", data.get('nl_amont', '')] + 
        [data.get(f'nl_amont_{fournisseur}', '') for fournisseur in fournisseurs] + [""],
        ["Débit traversier de pointe", "m³/h", data.get('debit_pointe', '')] + 
        [data.get(f'debit_pointe_{fournisseur}', '') for fournisseur in fournisseurs] + [""],
        ["Débit traversier maxi acceptable en régime nominal", "m³/h", data.get('debit_max', '')] + 
        [data.get(f'debit_max_{fournisseur}', '') for fournisseur in fournisseurs] + [""],
        ["Hauteur minimum d'évacuation des déchets", "mm", data.get('hauteur_evacuation', '')] + 
        [data.get(f'hauteur_evacuation_{fournisseur}', '') for fournisseur in fournisseurs] + [""],
        ["Perte de charge grille encrassée à 30% à débit maxi", "mCE", data.get('perte_charge_encrassee', '')] + 
        [data.get(f'perte_charge_encrassee_{fournisseur}', '') for fournisseur in fournisseurs] + [""],
        ["Perte de charge maximale en fonctionnement normal", "mCE", data.get('perte_charge_max', '')] + 
        [data.get(f'perte_charge_max_{fournisseur}', '') for fournisseur in fournisseurs] + [""],
        ["Angle du tamis", "°", data.get('angle_tamis', '')] + 
        [data.get(f'angle_tamis_{fournisseur}', '') for fournisseur in fournisseurs] + [""]
    ]
    
    excel_data.extend(sections_3)
    
    # Section 4 - Accessoires et matériaux
    sections_4 = [
        ["4. ACCESSOIRES, MATERIAUX, PROTECTIONS ET SECURITES REQUIS", "", "", "", ""] + [""] * len(fournisseurs),
        ["Limiteur de couple / type fourni", "", data.get('limiteur_couple', '')] + 
        [data.get(f'limiteur_couple_{fournisseur}', '') for fournisseur in fournisseurs] + [""],
        ["Capotage de sécurité", "", data.get('capotage_securite', '')] + 
        [data.get(f'capotage_securite_{fournisseur}', '') for fournisseur in fournisseurs] + [""],
        ["Sonde de niveau différentiel", "", data.get('sonde_niveau', '')] + 
        [data.get(f'sonde_niveau_{fournisseur}', '') for fournisseur in fournisseurs] + [""],
        ["Supportage inclus", "", data.get('supportage', '')] + 
        [data.get(f'supportage_{fournisseur}', '') for fournisseur in fournisseurs] + [""],
        ["Adaptation au canal", "", data.get('adaptation_canal', '')] + 
        [data.get(f'adaptation_canal_{fournisseur}', '') for fournisseur in fournisseurs] + [""],
        ["Raccord pour gaine d'air frais", "", data.get('raccord_gaine', '')] + 
        [data.get(f'raccord_gaine_{fournisseur}', '') for fournisseur in fournisseurs] + [""],
        ["Rampe de lavage", "", data.get('rampe_lavage', '')] + 
        [data.get(f'rampe_lavage_{fournisseur}', '') for fournisseur in fournisseurs] + [""],
        ["Brosse de nettoyage", "", data.get('brosse_nettoyage', '')] + 
        [data.get(f'brosse_nettoyage_{fournisseur}', '') for fournisseur in fournisseurs] + [""],
        ["Trémie de liaison", "", data.get('tremie_liaison', '')] + 
        [data.get(f'tremie_liaison_{fournisseur}', '') for fournisseur in fournisseurs] + [""],
        ["Capteurs et moteurs précablés", "", data.get('capteurs_moteurs', '')] + 
        [data.get(f'capteurs_moteurs_{fournisseur}', '') for fournisseur in fournisseurs] + [""],
        ["Coffret de commande", "", data.get('coffret_commande', '')] + 
        [data.get(f'coffret_commande_{fournisseur}', '') for fournisseur in fournisseurs] + [""],
        ["Electrovanne eau de lavage", "", data.get('electrovanne', '')] + 
        [data.get(f'electrovanne_{fournisseur}', '') for fournisseur in fournisseurs] + [""],
        ["Matériau grille", "", data.get('materiau_grille', '')] + 
        [data.get(f'materiau_grille_{fournisseur}', '') for fournisseur in fournisseurs] + [""],
        ["Matériau lamelles émergées", "", data.get('materiau_lamelles_emergées', '')] + 
        [data.get(f'materiau_lamelles_emergées_{fournisseur}', '') for fournisseur in fournisseurs] + [""],
        ["Matériau lamelles immergées", "", data.get('materiau_lamelles_immergées', '')] + 
        [data.get(f'materiau_lamelles_immergées_{fournisseur}', '') for fournisseur in fournisseurs] + [""],
        ["Matériau chassis", "", data.get('materiau_chassis', '')] + 
        [data.get(f'materiau_chassis_{fournisseur}', '') for fournisseur in fournisseurs] + [""]
    ]
    
    excel_data.extend(sections_4)
    
    # Section 1 - Référence des fournitures
    sections_ref = [
        ["1. REFERENCE DES FOURNITURES", "", "", "", ""] + [""] * len(fournisseurs),
        ["Fournisseur", "", data.get('fournisseur', '')] + 
        [data.get(f'fournisseur_{fournisseur}', '') for fournisseur in fournisseurs] + [""],
        ["Modèle - Type", "", data.get('modele_type', '')] + 
        [data.get(f'modele_type_{fournisseur}', '') for fournisseur in fournisseurs] + [""]
    ]
    
    excel_data.extend(sections_ref)
    
    # Calcul du nombre de colonnes
    num_cols = 5 + len(fournisseurs)  # 5 colonnes de base + fournisseurs
    
    # Création du DataFrame avec le bon nombre de colonnes
    columns = ['A', 'B', 'C', 'D', 'E'] + [f'F{i+1}' for i in range(len(fournisseurs))]
    df = pd.DataFrame(excel_data, columns=columns[:num_cols])
    
    # Création du fichier Excel en mémoire
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Tamis', index=False, header=False)
    
    return output.getvalue()

def create_comparison_table(fournisseurs_data):
    """Crée un tableau de comparaison des fournisseurs"""
    if not fournisseurs_data:
        return pd.DataFrame()
    
    data = []
    for fournisseur, details in fournisseurs_data.items():
        data.append({
            'Fournisseur': fournisseur,
            'Prix (€)': details.get('prix', 0),
            'Délai (jours)': details.get('delai', 0),
            'Options': details.get('options', ''),
            'Remise (%)': details.get('remise', 0),
            'Commentaires': details.get('commentaires', '')
        })
    
    df = pd.DataFrame(data)
    
    # Trier par prix
    if not df.empty and 'Prix (€)' in df.columns:
        df = df.sort_values('Prix (€)')
    
    return df

def style_comparison_table(df):
    """Applique un style coloré au tableau de comparaison"""
    if df.empty:
        return df
    
    styled_df = df.copy()
    
    # Colorer les prix
    if 'Prix (€)' in df.columns:
        min_price = df['Prix (€)'].min()
        max_price = df['Prix (€)'].max()
        
        def color_price(val):
            if val == min_price:
                return 'background-color: #90EE90'  # Vert clair pour le moins cher
            elif val == max_price:
                return 'background-color: #FFB6C1'  # Rouge clair pour le plus cher
            else:
                return 'background-color: #FFFACD'  # Jaune clair pour les intermédiaires
        
        styled_df = styled_df.style.applymap(color_price, subset=['Prix (€)'])
    
    return styled_df

def import_fournisseur_data(uploaded_file):
    """Importe les données d'un fournisseur depuis un fichier Excel"""
    try:
        df = pd.read_excel(uploaded_file, sheet_name='Tamis', header=None)
        return df
    except Exception as e:
        st.error(f"Erreur lors de l'import du fichier: {str(e)}")
        return None

def main():
    st.set_page_config(page_title="Formulaire Degrilleur", page_icon="📊", layout="wide")
    
    st.title("📊 Formulaire Technique - Degrilleur Tamis Escalier")
    st.markdown("Renseignez les informations techniques pour générer le fichier Excel")
    
    # Initialisation de la session state
    if 'fournisseurs' not in st.session_state:
        st.session_state.fournisseurs = []
    if 'fournisseurs_data' not in st.session_state:
        st.session_state.fournisseurs_data = {}
    
    # Configuration des fournisseurs
    st.sidebar.header("🔧 Configuration des fournisseurs")
    
    nb_fournisseurs = st.sidebar.number_input("Nombre de fournisseurs", min_value=1, max_value=10, value=1)
    
    # Mise à jour de la liste des fournisseurs
    if len(st.session_state.fournisseurs) != nb_fournisseurs:
        st.session_state.fournisseurs = [f"Fournisseur {i+1}" for i in range(nb_fournisseurs)]
        # Réinitialiser les données des fournisseurs
        st.session_state.fournisseurs_data = {}
    
    # Saisie des noms des fournisseurs
    st.sidebar.subheader("Noms des fournisseurs")
    for i in range(nb_fournisseurs):
        nouveau_nom = st.sidebar.text_input(f"Nom du fournisseur {i+1}", 
                                          value=st.session_state.fournisseurs[i],
                                          key=f"nom_fournisseur_{i}")
        if nouveau_nom != st.session_state.fournisseurs[i]:
            st.session_state.fournisseurs[i] = nouveau_nom
    
    # Import des données fournisseurs
    st.sidebar.subheader("📤 Import données fournisseurs")
    for fournisseur in st.session_state.fournisseurs:
        uploaded_file = st.sidebar.file_uploader(f"Import {fournisseur}", 
                                               type=['xlsx'],
                                               key=f"upload_{fournisseur}")
        if uploaded_file is not None:
            data = import_fournisseur_data(uploaded_file)
            if data is not None:
                st.sidebar.success(f"Données de {fournisseur} importées avec succès!")
    
    # Formulaire principal
    with st.form("degrilleur_form"):
        form_data = create_dynamic_form(st.session_state.fournisseurs)
        submitted = st.form_submit_button("Générer le fichier Excel")
    
    if submitted:
        # Génération du fichier Excel
        excel_file = create_excel_file(form_data, st.session_state.fournisseurs)
        filename = f"degrilleur_tamis_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        
        # Téléchargement
        st.success("Fichier Excel généré avec succès!")
        st.download_button(
            label="📥 Télécharger le fichier Excel",
            data=excel_file,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    # Section comparaison des fournisseurs
    st.header("📊 Tableau de comparaison des fournisseurs")
    
    # Saisie des données de comparaison
    st.subheader("Saisie des offres fournisseurs")
    
    cols = st.columns(min(3, nb_fournisseurs))
    
    for idx, fournisseur in enumerate(st.session_state.fournisseurs):
        col_idx = idx % 3
        with cols[col_idx]:
            st.markdown(f"**{fournisseur}**")
            
            if fournisseur not in st.session_state.fournisseurs_data:
                st.session_state.fournisseurs_data[fournisseur] = {}
            
            prix = st.number_input(f"Prix (€)", 
                                 value=st.session_state.fournisseurs_data[fournisseur].get('prix', 0),
                                 key=f"prix_{fournisseur}")
            delai = st.number_input(f"Délai (jours)", 
                                  value=st.session_state.fournisseurs_data[fournisseur].get('delai', 0),
                                  key=f"delai_{fournisseur}")
            options = st.text_input(f"Options", 
                                  value=st.session_state.fournisseurs_data[fournisseur].get('options', ''),
                                  key=f"options_{fournisseur}")
            remise = st.number_input(f"Remise (%)", 
                                   value=st.session_state.fournisseurs_data[fournisseur].get('remise', 0),
                                   key=f"remise_{fournisseur}")
            commentaires = st.text_area(f"Commentaires", 
                                      value=st.session_state.fournisseurs_data[fournisseur].get('commentaires', ''),
                                      key=f"commentaires_{fournisseur}")
            
            # Mise à jour des données
            st.session_state.fournisseurs_data[fournisseur] = {
                'prix': prix,
                'delai': delai,
                'options': options,
                'remise': remise,
                'commentaires': commentaires
            }
    
    # Affichage du tableau de comparaison
    if st.session_state.fournisseurs_data:
        comparison_df = create_comparison_table(st.session_state.fournisseurs_data)
        if not comparison_df.empty:
            st.subheader("Comparaison des offres")
            styled_df = style_comparison_table(comparison_df)
            st.dataframe(styled_df, use_container_width=True)
            
            # Synthèse
            st.subheader("📝 Synthèse et recommandations")
            synthese = st.text_area("Rédigez votre synthèse et recommandations ici...", 
                                  height=200,
                                  placeholder="Analysez les offres et rédigez vos recommandations...")
            
            if st.button("Sauvegarder la synthèse"):
                st.success("Synthèse sauvegardée!")

def create_dynamic_form(fournisseurs):
    """Crée un formulaire dynamique avec le nombre de fournisseurs configuré"""
    form_data = {}
    
    st.header("📋 Informations générales")
    col1, col2 = st.columns(2)
    
    with col1:
        form_data['nom_projet'] = st.text_input("Nom et n° du projet", value="")
        form_data['redacteur'] = st.text_input("Rédacteur", value="")
        form_data['verificateur'] = st.text_input("Vérificateur/Approbateur", value="")
        form_data['repere_pid'] = st.text_input("Repère PID", value="DGR D 1001 / DGR D 2001")
        
    with col2:
        form_data['stp_numero'] = st.text_input("Spécification Technique Particulière N°", value="")
        form_data['statut'] = st.text_input("Statut", value="")
        form_data['indice'] = st.text_input("Indice", value="")
    
    # Sections techniques avec colonnes dynamiques pour les fournisseurs
    sections = [
        ("🔧 Conditions de fonctionnement", [
            ('nb_equipements', 'Nombre d\'équipements', '2 (1 + 1 secours intégral installé)'),
            ('situation', 'Situation', 'Intérieur'),
            ('effluent', 'Effluent à traiter', 'Effluent brut'),
            ('zone_atex', 'Zone ATEX', 'Non'),
            ('implantation', 'Implantation dans l\'ouvrage', 'Voir plan joint')
        ]),
        ("📊 Performances et dimensionnement", [
            ('maille', 'Maille (mm)', '6'),
            ('type_alimentation', 'Type d\'alimentation', 'canal'),
            ('refus_tamis', 'Refus de tamis à relever (l/h)', '1500'),
            ('largeur_canal', 'Largeur canal (m)', '0.8'),
            ('radier_canal', 'Radier canal (NGF)', '191.2'),
            ('niveau_circulation', 'Niveau zone circulation (NGF)', '192.67'),
            ('nl_amont', 'NL AMONT TAMIS maxi (NGF)', '192.39'),
            ('debit_pointe', 'Débit traversier de pointe (m³/h)', '1800'),
            ('debit_max', 'Débit traversier maxi (m³/h)', '1800'),
            ('hauteur_evacuation', 'Hauteur min évacuation déchets', 'à préciser par fournisseur'),
            ('perte_charge_encrassee', 'Perte de charge grille encrassée (mCE)', 'à préciser par fournisseur'),
            ('perte_charge_max', 'Perte de charge maximale (mCE)', 'à préciser par fournisseur'),
            ('angle_tamis', 'Angle du tamis (°)', 'à préciser par fournisseur')
        ])
    ]
    
    for section_title, fields in sections:
        st.header(section_title)
        
        for field_key, field_label, default_value in fields:
            cols = st.columns([2] + [1] * len(fournisseurs))
            
            with cols[0]:
                form_data[field_key] = st.text_input(field_label, value=default_value, key=f"demande_{field_key}")
            
            for idx, fournisseur in enumerate(fournisseurs):
                with cols[idx + 1]:
                    form_data[f"{field_key}_{fournisseur}"] = st.text_input(
                        f"{fournisseur}", 
                        value="", 
                        key=f"{field_key}_{fournisseur}"
                    )
    
    return form_data

if __name__ == "__main__":
    main()