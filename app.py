import os
import pandas as pd
import streamlit as st

# Configuration de la page Streamlit
st.set_page_config(
    page_title="AMADIS - Gestion Financière", page_icon="🏥", layout="wide"
)

# Nom du fichier de base de données Excel persistant
DB_FILE = "base_amicale_python.xlsx"


def init_db():
  """Initialise le fichier Excel de la base de données s'il n'existe pas."""
  if not os.path.exists(DB_FILE):
    df_membres = pd.DataFrame(
        columns=[
            "ID Membre",
            "Nom et Prénoms",
            "Contact",
            "Date Adhesion",
            "Statut",
        ]
    )
    df_cotisations = pd.DataFrame(
        columns=[
            "ID Membre",
            "Nom et Prénoms",
            "Mois",
            "Année",
            "Montant",
            "Date Paiement",
        ]
    )
    df_evenements = pd.DataFrame(
        columns=[
            "ID Membre",
            "Nom et Prénoms",
            "Type Evenement",
            "Montant Verse",
            "Date",
            "Assistance Anterieure",
        ]
    )
    df_depenses = pd.DataFrame(
        columns=["Libelle", "Montant", "Date", "Categorie"]
    )

    with pd.ExcelWriter(DB_FILE, engine="openpyxl") as writer:
      df_membres.to_excel(writer, sheet_name="Membres", index=False)
      df_cotisations.to_excel(writer, sheet_name="Cotisations", index=False)
      df_evenements.to_excel(writer, sheet_name="Evenements", index=False)
      df_depenses.to_excel(writer, sheet_name="Depenses", index=False)


init_db()


def load_data():
  """Charge toutes les tables et nettoie les doublons de colonnes."""
  try:
    xls = pd.ExcelFile(DB_FILE)
    sheets = xls.sheet_names

    # Membres
    if "Membres" in sheets:
      df_membres = pd.read_excel(DB_FILE, sheet_name="Membres", dtype=str)
      df_membres.columns = df_membres.columns.str.strip()
      df_membres = df_membres.loc[:, ~df_membres.columns.duplicated()]

      colonnes_officielles = [
          "ID Membre",
          "Nom et Prénoms",
          "Contact",
          "Date Adhesion",
          "Statut",
      ]
      for col in colonnes_officielles:
        if col not in df_membres.columns:
          df_membres[col] = ""
      df_membres = df_membres[colonnes_officielles]
    else:
      df_membres = pd.DataFrame(
          columns=[
              "ID Membre",
              "Nom et Prénoms",
              "Contact",
              "Date Adhesion",
              "Statut",
          ]
      )

    # Cotisations
    df_cotisations = (
        pd.read_excel(DB_FILE, sheet_name="Cotisations", dtype=str)
        if "Cotisations" in sheets
        else pd.DataFrame(
            columns=[
                "ID Membre",
                "Nom et Prénoms",
                "Mois",
                "Année",
                "Montant",
                "Date Paiement",
            ]
        )
    )
    df_cotisations = df_cotisations.loc[:, ~df_cotisations.columns.duplicated()]

    # Evenements
    df_evenements = (
        pd.read_excel(DB_FILE, sheet_name="Evenements", dtype=str)
        if "Evenements" in sheets
        else pd.DataFrame(
            columns=[
                "ID Membre",
                "Nom et Prénoms",
                "Type Evenement",
                "Montant Verse",
                "Date",
                "Assistance Anterieure",
            ]
        )
    )
    df_evenements = df_evenements.loc[:, ~df_evenements.columns.duplicated()]

    # Depenses
    df_depenses = (
        pd.read_excel(DB_FILE, sheet_name="Depenses", dtype=str)
        if "Depenses" in sheets
        else pd.DataFrame(columns=["Libelle", "Montant", "Date", "Categorie"])
    )
    df_depenses = df_depenses.loc[:, ~df_depenses.columns.duplicated()]

    return df_membres, df_cotisations, df_evenements, df_depenses
  except Exception as e:
    df_membres = pd.DataFrame(
        columns=[
            "ID Membre",
            "Nom et Prénoms",
            "Contact",
            "Date Adhesion",
            "Statut",
        ]
    )
    df_cotisations = pd.DataFrame(
        columns=[
            "ID Membre",
            "Nom et Prénoms",
            "Mois",
            "Année",
            "Montant",
            "Date Paiement",
        ]
    )
    df_evenements = pd.DataFrame(
        columns=[
            "ID Membre",
            "Nom et Prénoms",
            "Type Evenement",
            "Montant Verse",
            "Date",
            "Assistance Anterieure",
        ]
    )
    df_depenses = pd.DataFrame(
        columns=["Libelle", "Montant", "Date", "Categorie"]
    )
    save_data(df_membres, df_cotisations, df_evenements, df_depenses)
    return df_membres, df_cotisations, df_evenements, df_depenses


def save_data(df_membres, df_cotisations, df_evenements, df_depenses):
  """Sauvegarde toutes les tables dans le fichier Excel."""
  with pd.ExcelWriter(DB_FILE, engine="openpyxl") as writer:
    df_membres.to_excel(writer, sheet_name="Membres", index=False)
    df_cotisations.to_excel(writer, sheet_name="Cotisations", index=False)
    df_evenements.to_excel(writer, sheet_name="Evenements", index=False)
    df_depenses.to_excel(writer, sheet_name="Depenses", index=False)


def valider_telephone_ivoirien(tel):
  """Vérifie et formate un numéro de téléphone selon les standards ivoiriens (10 chiffres)."""
  if not tel or pd.isna(tel):
    return ""
  tel_clean = "".join(filter(str.isdigit, str(tel)))
  if tel_clean.startswith("225") and len(tel_clean) == 13:
    tel_clean = tel_clean[3:]
  if len(tel_clean) == 10:
    return f"{tel_clean[0:2]} {tel_clean[2:4]} {tel_clean[4:6]} {tel_clean[6:8]} {tel_clean[8:10]}"
  return str(tel).strip()


# Chargement initial
df_membres, df_cotisations, df_evenements, df_depenses = load_data()

# Barre latérale de navigation
st.sidebar.title("🏥 Navigation AMADIS")
section = st.sidebar.selectbox(
    "Choisir une section",
    [
        "Tableau de Bord",
        "Gestion des Membres",
        "Cotisations Mensuelles",
        "Calculateur Prêts / Secours",
        "Journal des Dépenses",
    ],
)

# Option d'impression globale
st.sidebar.markdown("---")
st.sidebar.subheader("🖨️ Impression & Export")
if st.sidebar.button("Imprimer / Exporter cette page"):
  st.markdown("<script>window.print();</script>", unsafe_allow_html=True)
  st.sidebar.success("Fenêtre d'impression ouverte !")

# ---------------------------------------------------------
# 1. TABLEAU DE BORD
# ---------------------------------------------------------
if section == "Tableau de Bord":
  st.title("🏥 AMICALE DES AGENTS DE SANTÉ DE GUÉYO (AMADIS)")
  st.markdown("### 📊 Tableau de Bord & Trésorerie")

  total_cotiz = (
      pd.to_numeric(df_cotisations["Montant"], errors="coerce").sum()
      if not df_cotisations.empty and "Montant" in df_cotisations.columns
      else 0
  )
  total_dep = (
      pd.to_numeric(df_depenses["Montant"], errors="coerce").sum()
      if not df_depenses.empty and "Montant" in df_depenses.columns
      else 0
  )
  nb_membres = len(df_membres) if not df_membres.empty else 0

  col1, col2, col3 = st.columns(3)
  col1.metric("Membres Inscrits", f"{nb_membres}")
  col2.metric("Total Cotisations Perçues", f"{total_cotiz:,.0f} FCFA")
  col3.metric("Total Dépenses / Sorties", f"{total_dep:,.0f} FCFA")

  st.markdown("---")
  st.info(
      "💡 Bienvenue sur le logiciel de gestion financière de l'AMADIS. Utilisez"
      " le menu latéral pour naviguer."
  )

# ---------------------------------------------------------
# 2. GESTION DES MEMBRES
# ---------------------------------------------------------
elif section == "Gestion des Membres":
  st.title("👥 Gestion des Membres - AMADIS")

  tab1, tab2, tab3 = st.tabs(
      ["Liste des Membres", "Ajouter un Membre", "Modifier / Supprimer"]
  )

  with tab1:
    st.subheader("Liste officielle des membres")
    if not df_membres.empty:
      st.dataframe(df_membres, use_container_width=True)
    else:
      st.warning("Aucun membre enregistré pour le moment.")

  with tab2:
    st.subheader("Enregistrer un nouveau membre")
  with st.form("form_ajout_membre", clear_on_submit=True):
      id_membre = st.text_input(
          "ID Membre (ex: AMA-001)",
          value=f"AMA-{len(df_membres)+1:03d}",
      )
      nom = st.text_input("Nom et Prénoms")
      contact = st.text_input("Contact Téléphonique (ex: 0701020304)")
      date_adhesion = st.date_input("Date d'adhésion")
      statut = st.selectbox("Statut", ["Actif", "Suspendu", "Retraité"])

      submit_ajout = st.form_submit_button("Ajouter le membre")

      if submit_ajout:
        if nom.strip() == "":
          st.error("Le nom et prénoms sont obligatoires.")
        else:
          contact_forme = valider_telephone_ivoirien(contact)
          nouveau_ligne = pd.DataFrame(
              [
                  {
                      "ID Membre": str(id_membre).strip(),
                      "Nom et Prénoms": str(nom).strip().upper(),
                      "Contact": str(contact_forme),
                      "Date Adhesion": str(date_adhesion),
                      "Statut": str(statut),
                  }
              ]
          )
          df_membres = pd.concat([df_membres, nouveau_ligne], ignore_index=True)
          save_data(df_membres, df_cotisations, df_evenements, df_depenses)
          st.success(f"Membre {nom} ajouté avec succès !")
          st.rerun()

  with tab3:
    st.subheader("Modifier ou Supprimer un membre")
    if not df_membres.empty and "ID Membre" in df_membres.columns:
      liste_ids = df_membres["ID Membre"].dropna().tolist()
      if liste_ids:
        selected_id = st.selectbox(
            "Sélectionner l'ID du membre à modifier", liste_ids
        )

        membre_actuel = df_membres[
            df_membres["ID Membre"] == selected_id
        ].iloc[0]

        with st.form("form_modif_membre"):
          nouveau_nom = st.text_input(
              "Nom et Prénoms",
              value=str(membre_actuel.get("Nom et Prénoms", "")),
          )
          nouveau_contact = st.text_input(
              "Contact Téléphonique",
              value=str(membre_actuel.get("Contact", "")),
          )
          statut_actuel = str(membre_actuel.get("Statut", "Actif"))
          statuts_possibles = ["Actif", "Suspendu", "Retraité"]
          idx_statut = (
              statuts_possibles.index(statut_actuel)
              if statut_actuel in statuts_possibles
              else 0
          )
          nouveau_statut = st.selectbox(
              "Statut", statuts_possibles, index=idx_statut
          )

          col_m1, col_m2 = st.columns(2)
          submit_mod = col_m1.form_submit_button("Enregistrer les modifications")
          submit_supp = col_m2.form_submit_button("Supprimer ce membre")

          if submit_mod:
            contact_forme = valider_telephone_ivoirien(nouveau_contact)
            df_membres["ID Membre"] = df_membres["ID Membre"].astype(str)
            df_membres.loc[
                df_membres["ID Membre"] == str(selected_id), "Nom et Prénoms"
            ] = str(nouveau_nom).upper()
            df_membres.loc[
                df_membres["ID Membre"] == str(selected_id), "Contact"
            ] = str(contact_forme)
            df_membres.loc[
                df_membres["ID Membre"] == str(selected_id), "Statut"
            ] = str(nouveau_statut)

            save_data(df_membres, df_cotisations, df_evenements, df_depenses)
            st.success("Modifications enregistrées avec succès !")
            st.rerun()

          if submit_supp:
            df_membres = df_membres[df_membres["ID Membre"] != selected_id]
            save_data(df_membres, df_cotisations, df_evenements, df_depenses)
            st.warning("Membre supprimé.")
            st.rerun()
      else:
        st.info("Aucun ID membre valide trouvé.")
    else:
      st.info("Aucun membre à modifier.")

# ---------------------------------------------------------
# 3. COTISATIONS MENSUELLES
# ---------------------------------------------------------
elif section == "Cotisations Mensuelles":
  st.title("💰 Gestion des Cotisations - AMADIS")
  st.markdown("Montant standard par membre : **1 000 FCFA / mois**")

  if df_membres.empty:
    st.warning("Veuillez d'abord enregistrer des membres.")
  else:
    with st.form("form_cotisation"):
      options_membres = (
          df_membres["ID Membre"].astype(str)
          + " - "
          + df_membres["Nom et Prénoms"].astype(str)
      ).tolist()
      membre_choisi = st.selectbox("Sélectionner le membre", options_membres)
      mois = st.selectbox(
          "Mois",
          [
              "Janvier",
              "Février",
              "Mars",
              "Avril",
              "Mai",
              "Juin",
              "Juillet",
              "Août",
              "Septembre",
              "Octobre",
              "Novembre",
              "Décembre",
          ],
      )
      annee = st.selectbox("Année", ["2025", "2026", "2027"])
      montant = st.number_input("Montant (FCFA)", value=1000, step=500)

      submit_cotiz = st.form_submit_button("Enregistrer la cotisation")

      if submit_cotiz:
        id_m = membre_choisi.split(" - ")[0]
        nom_m = membre_choisi.split(" - ")[1]
        nouvelle_cotiz = pd.DataFrame(
            [
                {
                    "ID Membre": str(id_m),
                    "Nom et Prénoms": str(nom_m),
                    "Mois": str(mois),
                    "Année": str(annee),
                    "Montant": str(montant),
                    "Date Paiement": str(pd.Timestamp.now().strftime("%Y-%m-%d")),
                }
            ]
        )
        df_cotisations = pd.concat(
            [df_cotisations, nouvelle_cotiz], ignore_index=True
        )
        save_data(df_membres, df_cotisations, df_evenements, df_depenses)
        st.success(
            f"Cotisation de {montant} FCFA enregistrée pour {nom_m} ({mois}"
            f" {annee})."
        )
        st.rerun()

    st.subheader("Historique des Cotisations")
    if not df_cotisations.empty:
      st.dataframe(df_cotisations, use_container_width=True)
    else:
      st.info("Aucune cotisation enregistrée.")

# ---------------------------------------------------------
# 4. CALCULATEUR PRÊTS / SECOURS (Règles 60% / 80%)
# ---------------------------------------------------------
elif section == "Calculateur Prêts / Secours":
  st.title("📐 Calculateur des Versements - AMADIS")
  st.markdown(
      "Règles financières : **80%** du total si aucune assistance antérieure;"
      " **60%** si déjà perçu."
  )

  type_secours = st.selectbox(
      "Type d'événement",
      [
          "Naissance (2 000 FCFA par membre)",
          "Décès Parent (5 000 FCFA par membre)",
          "Décès Membre (10 000 FCFA par membre)",
      ],
  )
  deja_recu = st.radio(
      "Le membre a-t-il déjà reçu une assistance antérieure ?",
      ["Non (Taux de 80%)", "Oui (Taux de 60%)"],
  )

  nb_actifs = (
      len(df_membres[df_membres["Statut"] == "Actif"])
      if not df_membres.empty and "Statut" in df_membres.columns
      else len(df_membres)
  )
  if nb_actifs == 0:
    nb_actifs = len(df_membres)

  cotisation_unitaire = 2000
  if "Décès Parent" in type_secours:
    cotisation_unitaire = 5000
  elif "Décès Membre" in type_secours:
    cotisation_unitaire = 10000

  total_attendu = nb_actifs * cotisation_unitaire
  taux = 0.80 if "Non" in deja_recu else 0.60
  montant_versement = total_attendu * taux

  st.markdown("---")
  col_c1, col_c2, col_c3 = st.columns(3)
  col_c1.metric("Membres Cotisants pris en compte", f"{nb_actifs}")
  col_c2.metric("Total Théorique Collecté", f"{total_attendu:,.0f} FCFA")
  col_c3.metric(
      f"Montant du Versement ({int(taux*100)}%)",
      f"{montant_versement:,.0f} FCFA",
  )

# ---------------------------------------------------------
# 5. JOURNAL DES DÉPENSES
# ---------------------------------------------------------
elif section == "Journal des Dépenses":
  st.title("💸 Journal des Dépenses - AMADIS")

  with st.form("form_depense"):
    libelle = st.text_input("Libellé de la dépense")
    montant_dep = st.number_input("Montant (FCFA)", value=5000, step=1000)
    categorie = st.selectbox(
        "Catégorie", ["Fonctionnement", "Secours Versé", "Aide Sociale", "Autre"]
    )
    date_dep = st.date_input("Date de la dépense")

    submit_dep = st.form_submit_button("Enregistrer la dépense")

    if submit_dep:
      nouvelle_dep = pd.DataFrame(
          [
              {
                  "Libelle": str(libelle),
                  "Montant": str(montant_dep),
                  "Date": str(date_dep),
                  "Categorie": str(categorie),
              }
          ]
      )
      df_depenses = pd.concat([df_depenses, nouvelle_dep], ignore_index=True)
      save_data(df_membres, df_cotisations, df_evenements, df_depenses)
      st.success("Dépense enregistrée avec succès !")
      st.rerun()

  st.subheader("Historique des Dépenses")
  if not df_depenses.empty:
    st.dataframe(df_depenses, use_container_width=True)
  else:
    st.info("Aucune dépense enregistrée.")