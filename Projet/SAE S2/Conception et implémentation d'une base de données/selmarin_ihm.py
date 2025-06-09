# -*- coding: utf-8 -*-
"""
Created on Thu Apr 17 17:37:00 2025

@author: Utilisateur
"""

import pandas as pd
import mysql.connector
import os
import customtkinter as ctk
from tkinter import messagebox, filedialog
from PIL import Image
import csv
import uuid

# Configuration de la base de données
CONFIG_BDD = {
    'host': 'localhost',
    'user': 'root',
    'password': '',
    'database': 'selmarin_create'
}

# Constantes globales pour les tables et requêtes prédéfinies
LISTE_TABLES = ['SAUNIER', 'CLIENT', 'PRODUIT', 'SORTIE', 'ENTREE', 'COÛTE', 'CONCERNER', 'ANNEEPRIX']
REQUETES_PREDEFINIES = {
    "Les 3 meilleurs contributeurs": """
        SELECT S.numSau, S.nomSau, S.prenomSau, SUM(E.qteEnt) AS Quantite_Totale_Fournie
        FROM SAUNIER S
        JOIN ENTREE E ON S.numSau = E.numSau
        GROUP BY S.numSau, S.nomSau, S.prenomSau
        ORDER BY Quantite_Totale_Fournie DESC
        LIMIT 3;
    """,
    "Top 5 des clients par chiffre d'affaires": """
        SELECT C.numCli, C.nomCli, SUM(CO.qteSort * C2.Prix_Vente) AS Chiffre_Affaires
        FROM CLIENT C
        JOIN SORTIE S ON C.numCli = S.numCli
        JOIN CONCERNER CO ON S.numSort = CO.numSort
        JOIN PRODUIT P ON CO.numPdt = P.numPdt
        JOIN COÛTE C2 ON P.numPdt = C2.numPdt
        GROUP BY C.numCli, C.nomCli, C.villeCli
        ORDER BY Chiffre_Affaires DESC
        LIMIT 5;
    """,
    "Insertion d'un saunier (GARNIER François)": """
        INSERT INTO SAUNIER VALUES (3, "GARNIER", "François", "La Flotte");
    """,
    "État des stocks": """
        SELECT numPdt, libPdt, stockPdt_t_ 
        FROM PRODUIT;
    """,
    "État des stocks par saunier": """
        SELECT S.nomSau, S.prenomSau, E.numPdt, P.libPdt, P.stockPdt_t_, SUM(E.qteEnt) AS Quantite_Totale_Entree
        FROM SAUNIER S
        JOIN ENTREE E ON S.numSau = E.numSau
        JOIN PRODUIT P ON E.numPdt = P.numPdt
        GROUP BY S.nomSau, S.prenomSau, E.numPdt, P.libPdt, P.stockPdt_t_;
    """,
    "Évolution du prix de vente par année": """
        SELECT A.Annee, C.Prix_Vente, P.libPdt 
        FROM PRODUIT P
        JOIN COÛTE C ON P.numPdt = C.numPdt
        JOIN ANNEEPRIX A ON C.Annee = A.Annee
        ORDER BY P.libPdt, A.Annee;
    """,
    "Créer utilisateur 'garnier'": """
        CREATE USER 'garnier'@'localhost' IDENTIFIED BY '1234';
    """,
    "Créer vue CA_Annuel": """
        CREATE VIEW CA_Annuel AS
        SELECT YEAR(S.datSort) AS Annee, SUM(CO.qteSort * C2.Prix_Vente) AS Chiffre_Affaires
        FROM SORTIE S
        JOIN CONCERNER CO ON S.numSort = CO.numSort
        JOIN COÛTE C2 ON CO.numPdt = C2.numPdt
        GROUP BY Annee
        ORDER BY Annee DESC;
    """,
    "Droits SELECT sur CA_Annuel pour garnier": """
        GRANT SELECT ON selmarin_create.CA_Annuel TO 'garnier'@'localhost';
    """,
    "Afficher le CA_Annuel (vue)": """
        SELECT * FROM CA_Annuel;
    """,
    "Marge moyenne par produit": """
        SELECT P.numPdt, P.libPdt, AVG(C.Prix_Vente - C.Prix_Achat) AS Marge_Moyenne
        FROM PRODUIT P
        JOIN COÛTE C ON P.numPdt = C.numPdt
        GROUP BY P.numPdt, P.libPdt
        ORDER BY Marge_Moyenne DESC;
    """,
    "Produits jamais vendus": """
        SELECT DISTINCT P.numPdt, P.libPdt
        FROM PRODUIT P
        WHERE P.numPdt NOT IN (SELECT DISTINCT C.numPdt FROM CONCERNER C);
    """
}

# Variables globales pour l'interface
table_actuelle = None
ligne_selectionnee = None
ligne_valeurs = None
ligne_frame_selectionne = None
historique_requetes_sql = []

# Initialisation des images pour les boutons
def init_images():
    """Charge les icônes pour les boutons de l'interface."""
    return {
        'rechercher': ctk.CTkImage(Image.open("Icons/loupe.png").convert("RGBA"), size=(24, 24)),
        'reset': ctk.CTkImage(Image.open("Icons/reset.png").convert("RGBA"), size=(24, 24)),
        'ajouter': ctk.CTkImage(Image.open("Icons/ajouter.png").convert("RGBA"), size=(24, 24)),
        'modifier': ctk.CTkImage(Image.open("Icons/modifier.png").convert("RGBA"), size=(24, 24)),
        'supprimer': ctk.CTkImage(Image.open("Icons/poubelle.png").convert("RGBA"), size=(24, 24)),
        'importer': ctk.CTkImage(Image.open("Icons/importer.png").convert("RGBA"), size=(24, 24))
    }

# Fonction de connexion à la base de données
def connecter_bdd():
    """Établit une connexion à la base de données MySQL."""
    try:
        return mysql.connector.connect(**CONFIG_BDD)
    except mysql.connector.Error as err:
        messagebox.showerror("Connexion échouée", str(err))
        return None

# Fonctions d'affichage des données
def afficher_table(nom_table):
    """Affiche les données d'une table sélectionnée dans l'interface."""
    global table_actuelle
    table_actuelle = nom_table
    recherche_var.set("")
    charger_donnees_table(nom_table)

def charger_donnees_table(nom_table, mot_cle=None):
    """Charge les données d'une table dans le cadre de l'interface, avec option de recherche."""
    global ligne_selectionnee, ligne_valeurs
    ligne_selectionnee = None
    ligne_valeurs = None
    for widget in table_frame.winfo_children():
        widget.destroy()

    conn = connecter_bdd()
    if not conn:
        return

    try:
        curseur = conn.cursor()
        requete = f"SELECT * FROM {nom_table}"

        if mot_cle:
            curseur.execute(f"SHOW COLUMNS FROM {nom_table}")
            colonnes = [col[0] for col in curseur.fetchall()]
            clauses = [f"{col} LIKE '%{mot_cle}%'" for col in colonnes]
            requete += " WHERE " + " OR ".join(clauses)

        curseur.execute(requete)
        lignes = curseur.fetchall()
        colonnes = [desc[0] for desc in curseur.description]

        tableau = ctk.CTkFrame(table_frame)
        tableau.pack(fill="both", expand=True)

        entetes = ctk.CTkFrame(tableau)
        entetes.pack(fill="x")
        for col in colonnes:
            ctk.CTkLabel(entetes, text=col, width=100, anchor="w").pack(side="left")

        for ligne in lignes:
            ligne_frame = ctk.CTkFrame(tableau)
            ligne_frame.pack(fill="x")
            ligne_frame.bind("<Button-1>", lambda e, l=ligne, f=ligne_frame: selectionner_ligne(l, f))

            for valeur in ligne:
                lbl = ctk.CTkLabel(ligne_frame, text=str(valeur), width=100, anchor="w")
                lbl.pack(side="left")
                lbl.bind("<Button-1>", lambda e, l=ligne, f=ligne_frame: selectionner_ligne(l, f))

    except mysql.connector.Error as err:
        messagebox.showerror("Erreur SQL", str(err))
    finally:
        curseur.close()
        conn.close()

def afficher_resultats_console(colonnes, lignes):
    """Affiche les résultats des requêtes SQL dans la console de l'interface."""
    for widget in resultat_console_frame.winfo_children():
        widget.destroy()

    entete_frame = ctk.CTkFrame(resultat_console_frame)
    entete_frame.pack(fill="x")
    for col in colonnes:
        ctk.CTkLabel(entete_frame, text=col, width=100, anchor="w").pack(side="left")

    for ligne in lignes:
        ligne_frame = ctk.CTkFrame(resultat_console_frame)
        ligne_frame.pack(fill="x")
        for valeur in ligne:
            ctk.CTkLabel(ligne_frame, text=str(valeur), width=100, anchor="w").pack(side="left")

# Gestion des lignes dans l'interface
def selectionner_ligne(ligne, frame):
    """Sélectionne une ligne dans la table affichée et met à jour l'interface."""
    global ligne_selectionnee, ligne_valeurs, ligne_frame_selectionne
    ligne_valeurs = ligne
    ligne_selectionnee = ligne[0]

    if ligne_frame_selectionne:
        ligne_frame_selectionne.configure(fg_color="transparent")
    ligne_frame_selectionne = frame
    ligne_frame_selectionne.configure(fg_color="#2a5a77")
    messagebox.showinfo("Sélection", f"Ligne sélectionnée : {ligne}")

def rechercher():
    """Effectue une recherche dans la table actuelle avec le mot-clé saisi."""
    if table_actuelle:
        charger_donnees_table(table_actuelle, recherche_var.get())

def reinitialiser_recherche():
    """Réinitialise la recherche et recharge la table actuelle."""
    if table_actuelle:
        recherche_var.set("")
        charger_donnees_table(table_actuelle)

# Fonctions de modification des données
def ouvrir_popup_ajout():
    """Ouvre une fenêtre pour ajouter un nouveau tuple à la table actuelle."""
    if not table_actuelle:
        return

    conn = connecter_bdd()
    curseur = conn.cursor()
    curseur.execute(f"SHOW COLUMNS FROM {table_actuelle}")
    colonnes = [col[0] for col in curseur.fetchall()]
    curseur.close()
    conn.close()

    popup = ctk.CTkToplevel(app)
    popup.title("Ajouter un tuple")
    popup.geometry("400x600")
    popup.lift()
    popup.focus_force()
    popup.grab_set()
    champs = {}

    for col in colonnes:
        ctk.CTkLabel(popup, text=col).pack()
        champ = ctk.CTkEntry(popup)
        champ.pack(pady=5)
        champs[col] = champ

    def valider():
        valeurs = [champ.get() for champ in champs.values()]
        placeholders = ', '.join(['%s'] * len(valeurs))
        requete = f"INSERT INTO {table_actuelle} VALUES ({placeholders})"

        conn = connecter_bdd()
        try:
            curseur = conn.cursor()
            curseur.execute(requete, valeurs)
            conn.commit()
            popup.destroy()
            charger_donnees_table(table_actuelle)
            messagebox.showinfo("Succès", "Tuple ajouté.")
        except mysql.connector.Error as err:
            messagebox.showerror("Erreur", str(err))
        finally:
            curseur.close()
            conn.close()

    ctk.CTkButton(popup, text="Valider", command=valider).pack(pady=20)

def ouvrir_popup_modification():
    """Ouvre une fenêtre pour modifier un tuple sélectionné."""
    global ligne_valeurs
    if not ligne_valeurs:
        messagebox.showwarning("Aucune sélection", "Veuillez sélectionner une ligne à modifier.")
        return

    conn = connecter_bdd()
    curseur = conn.cursor()
    curseur.execute(f"SHOW COLUMNS FROM {table_actuelle}")
    colonnes = [col[0] for col in curseur.fetchall()]
    curseur.close()
    conn.close()

    popup = ctk.CTkToplevel(app)
    popup.title("Modifier un tuple")
    popup.geometry("400x600")
    popup.lift()
    popup.focus_force()
    popup.grab_set()
    champs = {}

    for i, col in enumerate(colonnes):
        ctk.CTkLabel(popup, text=col).pack()
        champ = ctk.CTkEntry(popup)
        champ.insert(0, ligne_valeurs[i])
        champ.pack(pady=5)
        champs[col] = champ

    def valider():
        valeurs = [champ.get() for champ in champs.values()]
        requete = f"UPDATE {table_actuelle} SET " + ", ".join([f"{col} = %s" for col in colonnes]) + f" WHERE {colonnes[0]} = %s"
        valeurs.append(ligne_valeurs[0])

        conn = connecter_bdd()
        try:
            curseur = conn.cursor()
            curseur.execute(requete, valeurs)
            conn.commit()
            popup.destroy()
            charger_donnees_table(table_actuelle)
            messagebox.showinfo("Succès", "Tuple modifié.")
        except mysql.connector.Error as err:
            messagebox.showerror("Erreur", str(err))
        finally:
            curseur.close()
            conn.close()

    ctk.CTkButton(popup, text="Valider", command=valider).pack(pady=20)

def supprimer_ligne_selectionnee():
    """Supprime la ligne sélectionnée après confirmation."""
    global ligne_selectionnee
    if not ligne_selectionnee:
        messagebox.showwarning("Aucune sélection", "Veuillez sélectionner une ligne à supprimer.")
        return

    reponse = messagebox.askyesno("Confirmation", "Voulez-vous vraiment supprimer cette ligne ?")
    if not reponse:
        return

    conn = connecter_bdd()
    try:
        curseur = conn.cursor()
        curseur.execute(f"SHOW COLUMNS FROM {table_actuelle}")
        colonne_cle = curseur.fetchone()[0]
        requete = f"DELETE FROM {table_actuelle} WHERE {colonne_cle} = %s"
        curseur.execute(requete, (ligne_selectionnee,))
        conn.commit()
        charger_donnees_table(table_actuelle)
        messagebox.showinfo("Succès", "Tuple supprimé.")
    except mysql.connector.Error as err:
        messagebox.showerror("Erreur", str(err))
    finally:
        curseur.close()
        conn.close()

# Fonctions SQL
def inserer_requete(selection):
    """Insère une requête prédéfinie dans la zone de texte SQL."""
    requete_sql_var.set(REQUETES_PREDEFINIES[selection])

def maj_menu_deroulant():
    """Met à jour le menu déroulant des requêtes prédéfinies."""
    menu_requetes.configure(values=list(REQUETES_PREDEFINIES.keys()))

def executer_requete_sql():
    """Exécute la requête SQL saisie ou sélectionnée."""
    requete = requete_sql_var.get()
    if not requete.strip():
        return

    conn = connecter_bdd()
    if not conn:
        return

    try:
        curseur = conn.cursor()
        curseur.execute(requete)

        if requete.strip().lower().startswith("select"):
            resultats = curseur.fetchall()
            colonnes = [desc[0] for desc in curseur.description]
            afficher_resultats_console(colonnes, resultats)
        else:
            conn.commit()
            afficher_resultats_console(["Message"], [["Requête exécutée avec succès."]])

    except mysql.connector.Error as err:
        afficher_resultats_console(["Erreur"], [[str(err)]])
    finally:
        curseur.close()
        conn.close()

def importer_csv():
    """Importe un fichier CSV dans la table actuelle, avec gestion des dates."""
    if not table_actuelle:
        messagebox.showwarning("Aucune table sélectionnée", "Veuillez sélectionner une table avant d'importer un fichier CSV.")
        return

    fichier = filedialog.askopenfilename(
        title="Sélectionner un fichier CSV",
        filetypes=[("Fichiers CSV", "*.csv")]
    )
    if not fichier:
        return

    conn = connecter_bdd()
    if not conn:
        return

    curseur = None  # Initialize curseur to None to avoid UnboundLocalError
    try:
        # Lire le fichier CSV avec pandas
        df = pd.read_csv(fichier, sep=";")

        # Conversion des colonnes contenant des dates
        for col in df.columns:
            if "date" in col.lower() or "dat" in col.lower():
                try:
                    df[col] = pd.to_datetime(df[col], format="%d/%m/%Y", errors="coerce")
                    df[col] = df[col].dt.strftime("%Y-%m-%d %H:%M:%S")
                except Exception as e:
                    messagebox.showwarning("Erreur de date", f"Erreur lors de la conversion de la colonne {col} : {e}")

        # Nettoyage des noms de colonnes
        df.columns = [col.strip().replace(" ", "_") for col in df.columns]

        # Création du curseur
        curseur = conn.cursor()

        # Vérification des colonnes de la table
        curseur.execute(f"SHOW COLUMNS FROM {table_actuelle}")
        colonnes_table = [col[0] for col in curseur.fetchall()]
        if list(df.columns) != colonnes_table:
            messagebox.showerror(
                "Erreur de format",
                f"Les colonnes du CSV ({list(df.columns)}) ne correspondent pas à celles de la table {table_actuelle} ({colonnes_table})."
            )
            return

        # Préparation de la requête d'insertion
        colonnes = ", ".join(f"`{col}`" for col in df.columns)
        placeholders = ", ".join(["%s"] * len(df.columns))
        sql = f"INSERT INTO `{table_actuelle}` ({colonnes}) VALUES ({placeholders})"

        # Insertion des données
        for ligne in df.itertuples(index=False, name=None):
            try:
                curseur.execute(sql, ligne)
            except mysql.connector.Error as err:
                messagebox.showerror("Erreur d'insertion", f"Erreur avec {ligne} : {err}")

        conn.commit()
        charger_donnees_table(table_actuelle)
        messagebox.showinfo("Succès", f"Données importées depuis {fichier} dans la table {table_actuelle}.")

    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur lors de l'importation : {str(e)}")
    finally:
        if curseur is not None:  # Only close cursor if it was created
            curseur.close()
        if conn is not None:  # Ensure connection is closed
            conn.close()

# Initialisation de l'interface
def init_ui():
    """Configure et lance l'interface graphique."""
    global app, table_frame, console_frame, resultat_console_frame
    global recherche_var, requete_sql_var, menu_requetes, images

    ctk.set_appearance_mode("system")
    ctk.set_default_color_theme("dark-blue")

    app = ctk.CTk()
    app.title("Gestion de la base")
    app.attributes("-fullscreen", True)

    images = init_images()

    cadre_boutons = ctk.CTkFrame(app)
    cadre_boutons.pack(fill="x", padx=10, pady=10)
    for table in LISTE_TABLES:
        ctk.CTkButton(cadre_boutons, text=table, width=120, command=lambda t=table: afficher_table(t)).pack(side="left", padx=5)

    cadre_recherche = ctk.CTkFrame(app)
    cadre_recherche.pack(fill="x", padx=10)
    recherche_var = ctk.StringVar()
    ctk.CTkEntry(cadre_recherche, textvariable=recherche_var, width=300).pack(side="left", padx=5)

    boutons_recherche = [
        ("Rechercher", images['rechercher'], rechercher),
        ("Réinitialiser", images['reset'], reinitialiser_recherche),
        ("Ajouter", images['ajouter'], ouvrir_popup_ajout),
        ("Modifier", images['modifier'], ouvrir_popup_modification),
        ("Supprimer", images['supprimer'], supprimer_ligne_selectionnee),
        ("Importer CSV", images['importer'], importer_csv)
    ]
    for text, image, command in boutons_recherche:
        ctk.CTkButton(cadre_recherche, image=image, text=text, command=command).pack(side="left", padx=5)

    table_frame = ctk.CTkFrame(app)
    table_frame.pack(fill="both", expand=True, padx=10, pady=10)

    console_frame = ctk.CTkFrame(app)
    console_frame.pack(fill="x", padx=10, pady=(0, 10))
    ctk.CTkLabel(console_frame, text="Console SQL :").pack(anchor="w", padx=5)

    requete_sql_var = ctk.StringVar()
    zone_requete = ctk.CTkEntry(console_frame, textvariable=requete_sql_var, width=600)
    zone_requete.pack(side="left", padx=5, pady=5)

    menu_requetes = ctk.CTkOptionMenu(console_frame, values=list(REQUETES_PREDEFINIES.keys()), width=250, command=inserer_requete)
    menu_requetes.pack(side="left", padx=5)

    ctk.CTkButton(console_frame, text="Exécuter", command=executer_requete_sql).pack(side="left", padx=5)

    resultat_console_frame = ctk.CTkFrame(app)
    resultat_console_frame.pack(fill="both", expand=False, padx=10, pady=(0, 10))

    ctk.CTkButton(app, text="❌", command=app.destroy, width=40, height=40).pack(side="bottom", pady=10)

# Lancement de l'application
if __name__ == "__main__":
    init_ui()
    app.mainloop()