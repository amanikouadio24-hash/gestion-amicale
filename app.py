import streamlit as st
import pandas as pd
import os

DB_FILE = "base_amicale_python.xlsx"

def initialiser_base():
    if not os.path.exists(DB_FILE):
        with pd.ExcelWriter(DB_FILE, engine="openpyxl") as writer:
            df_membres = pd.DataFrame(columns=["ID Membre", "Nom & Prénoms", "Contact", "Date d'adhésion", "Statut"])
            df_membres.to_excel(writer, sheet_name="Membres", index=False)
            
            df_mensuel = pd.DataFrame(columns=["ID Membre", "Nom & Prénoms", "Total Général Versé"])
            df_mensuel.to_excel(writer, sheet_name="Cotisations_Mensuelles", index=False)
            
            df_departs = pd.DataFrame(columns=["ID Membre", "Nom & Prénoms", "Motif", "Total Versé", "A déjà bénéficié ?", "Taux", "Montant Remboursable"])
            df_departs.to_excel(writer, sheet_name="Retraites_Mutations", index=False)
            
            df_decis = pd.DataFrame(columns=["Date", "Type", "Motif", "Caisse Imputée", "Montant"])
            df_decis.to_excel(writer, sheet_name="Decaissements", index=False)

initialiser_base()

st.set_page_config(page_title="Gestion Financière - Amicale", layout="wide")

st.title("🤝 Logiciel de Gestion Financière - Amicale")
st.sidebar.header("Navigation")

menu = st.sidebar.selectbox("Choisir une section", ["Tableau de Bord", "Gestion des Membres", "Retraites & Mutations", "Décaissements"])

if menu == "Tableau de Bord":
    st.subheader("📊 Vue d'ensemble de la Trésorerie")
    
    xls = pd.ExcelFile(DB_FILE)
    df_mens = pd.read_excel(xls, "Cotisations_Mensuelles")
    df_decis = pd.read_excel(xls, "Decaissements")
    df_departs = pd.read_excel(xls, "Retraites_Mutations")
    
    total_cotis = df_mens["Total Général Versé"].sum() if not df_mens.empty else 0
    total_rembours = df_departs["Montant Remboursable"].sum() if not df_departs.empty else 0
    
    st.metric("Total Cotisations Mensuelles Perçues", f"{total_cotis:,.0f} FCFA")
    st.metric("Total Remboursements Versés (Retraites/Mutations)", f"{total_rembours:,.0f} FCFA")

elif menu == "Gestion des Membres":
    st.subheader("👥 Gestion des Membres (Ajout, Modification, Suppression)")
    
    df_membres = pd.read_excel(DB_FILE, sheet_name="Membres")
    
    # Onglets pour séparer l'ajout, la modification et la suppression
    tab_liste, tab_ajout, tab_modif, tab_suppr = st.tabs(["Liste des Membres", "Ajouter un Membre", "Modifier un Membre", "Supprimer un Membre"])
    
    with tab_liste:
        st.dataframe(df_membres, use_container_width=True)
        
    with tab_ajout:
        with st.form("ajout_membre"):
            st.write("Ajouter un nouveau membre")
            id_m = st.text_input("ID Membre (ex: M005)")
            nom_m = st.text_input("Nom & Prénoms")
            contact_m = st.text_input("Contact")
            date_adh = st.date_input("Date d'adhésion")
            submit = st.form_submit_button("Enregistrer")
            
            if submit and id_m and nom_m:
                new_row = pd.DataFrame([[id_m, nom_m, contact_m, str(date_adh), "Actif"]], 
                                       columns=["ID Membre", "Nom & Prénoms", "Contact", "Date d'adhésion", "Statut"])
                df_membres = pd.concat([df_membres, new_row], ignore_index=True)
                with pd.ExcelWriter(DB_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                    df_membres.to_excel(writer, sheet_name="Membres", index=False)
                st.success("Membre ajouté avec succès ! (Rafraîchissez pour voir)")

    with tab_modif:
        if df_membres.empty:
            st.info("Aucun membre enregistré pour le moment.")
        else:
            id_a_modifier = st.selectbox("Sélectionner le membre à modifier", df_membres["ID Membre"].astype(str) + " - " + df_membres["Nom & Prénoms"])
            if id_a_modifier:
                selected_id = id_a_modifier.split(" - ")[0]
                membre_data = df_membres[df_membres["ID Membre"].astype(str) == selected_id].iloc[0]
                
                with st.form("form_modif"):
                    nouveau_nom = st.text_input("Nom & Prénoms", value=str(membre_data["Nom & Prénoms"]))
                    nouveau_contact = st.text_input("Contact", value=str(membre_data["Contact"]))
                    nouveau_statut = st.selectbox("Statut", ["Actif", "Retraité", "Muté", "Suspendu"], index=0)
                    
                    btn_sauver_modif = st.form_submit_button("Enregistrer les modifications")
                    
                    if btn_sauver_modif:
                        df_membres.loc[df_membres["ID Membre"].astype(str) == selected_id, "Nom & Prénoms"] = nouveau_nom
                        df_membres.loc[df_membres["ID Membre"].astype(str) == selected_id, "Contact"] = nouveau_contact
                        df_membres.loc[df_membres["ID Membre"].astype(str) == selected_id, "Statut"] = nouveau_statut
                        
                        with pd.ExcelWriter(DB_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                            df_membres.to_excel(writer, sheet_name="Membres", index=False)
                        st.success("Modifications enregistrées avec succès !")

    with tab_suppr:
        if df_membres.empty:
            st.info("Aucun membre enregistré pour le moment.")
        else:
            id_a_supprimer = st.selectbox("Sélectionner le membre à supprimer", df_membres["ID Membre"].astype(str) + " - " + df_membres["Nom & Prénoms"], key="suppr_select")
            if st.button("Supprimer définitivement ce membre", type="primary"):
                selected_id = id_a_supprimer.split(" - ")[0]
                df_membres = df_membres[df_membres["ID Membre"].astype(str) != selected_id]
                
                with pd.ExcelWriter(DB_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                    df_membres.to_excel(writer, sheet_name="Membres", index=False)
                st.success("Membre supprimé avec succès !")

elif menu == "Retraites & Mutations":
    st.subheader("🚪 Gestion des Départs (80% / 60%)")
    st.info("Règle : 80% des cotisations si aucune assistance reçue, 60% sinon.")
    
    df_departs = pd.read_excel(DB_FILE, sheet_name="Retraites_Mutations")
    st.dataframe(df_departs, use_container_width=True)
    
    with st.form("calcul_depart"):
        id_Membre_dep = st.text_input("ID du Membre concerné")
        nom_dep = st.text_input("Nom & Prénoms")
        motif_dep = st.selectbox("Motif", ["Retraite", "Mutation"])
        total_verse_dep = st.number_input("Total des cotisations versées par le membre (FCFA)", min_value=0, step=1000)
        assistance = st.selectbox("A déjà bénéficié d'une assistance (Mariage/Décès/Naissance) ?", ["Non", "Oui"])
        
        btn_calc = st.form_submit_button("Calculer et Enregistrer le Départ")
        
        if btn_calc and id_Membre_dep:
            taux = 0.60 if assistance == "Oui" else 0.80
            montant_rembours = total_verse_dep * taux
            
            new_dep = pd.DataFrame([[id_Membre_dep, nom_dep, motif_dep, total_verse_dep, assistance, taux, montant_rembours]],
                                   columns=["ID Membre", "Nom & Prénoms", "Motif", "Total Versé", "A déjà bénéficié ?", "Taux", "Montant Remboursable"])
            df_departs = pd.concat([df_departs, new_dep], ignore_index=True)
            
            with pd.ExcelWriter(DB_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                df_departs.to_excel(writer, sheet_name="Retraites_Mutations", index=False)
            st.success(f"Calcul effectué ! Montant à verser : {montant_rembours:,.0f} FCFA (Taux appliqué : {int(taux*100)}%)")

elif menu == "Décaissements":
    st.subheader("💸 Registre des Dépenses et Sorties de Caisse")
    
    df_decis = pd.read_excel(DB_FILE, sheet_name="Decaissements")
    st.dataframe(df_decis, use_container_width=True)
    
    with st.form("ajout_decis"):
        st.write("Enregistrer une nouvelle sortie d'argent")
        date_decis = st.date_input("Date du décaissement")
        type_decis = st.selectbox("Type de Sortie", ["Dépense Générale", "Aide Événement", "Remboursement Retraite/Mutation"])
        motif_decis = st.text_input("Motif / Bénéficiaire (ex: Fonctionnement bureau, Mariage de X...)")
        caisse_imput = st.selectbox("Caisse Imputée", ["Caisse Mensuelle", "Caisse Événements (Collecte dédiée)"])
        montant_decis = st.number_input("Montant (FCFA)", min_value=0, step=1000)
        
        submit_decis = st.form_submit_button("Enregistrer le Décaissement")
        
        if submit_decis and motif_decis:
            new_d = pd.DataFrame([[str(date_decis), type_decis, motif_decis, caisse_imput, montant_decis]],
                                 columns=["Date", "Type", "Motif", "Caisse Imputée", "Montant"])
            df_decis = pd.concat([df_decis, new_d], ignore_index=True)
            
            with pd.ExcelWriter(DB_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                df_decis.to_excel(writer, sheet_name="Decaissements", index=False)
            st.success("Décaissement enregistré avec succès !")