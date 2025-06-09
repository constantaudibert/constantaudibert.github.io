# -*- coding: utf-8 -*-
"""
Created on Sun Jun  8 14:16:36 2025

@author: Utilisateur
"""

# -*- coding: utf-8 -*-
"""
Created on Fri Jun  6 11:44:15 2025

@author: Utilisateur
"""

# Importation des bibliothèques nécessaires
import subprocess  # Pour exécuter des commandes système (comme l'installation de modules)
import sys  # Pour accéder aux informations du système (par exemple, l'exécutable Python)
import customtkinter as ctk  # Bibliothèque pour créer une interface graphique moderne
import pandas as pd  # Pour manipuler et analyser les données (DataFrames)
import os  # Pour interagir avec le système de fichiers
import json  # Pour parser les chaînes JSON
import re  # Pour les expressions régulières (nettoyage des chaînes)
import pyperclip  # Pour copier du texte dans le presse-papiers
import platform  # Pour identifier le système d'exploitation
import threading  # Pour exécuter des tâches en parallèle (threading)
from tkinter import filedialog  # Pour ouvrir des boîtes de dialogue de sélection de fichiers
from scipy.stats import chi2_contingency, fisher_exact  # Pour les tests statistiques Chi² et Fisher

# Installation automatique des dépendances nécessaires via pip
subprocess.check_call([sys.executable, "-m", "pip", "install", "customtkinter"])  # Installe customtkinter
subprocess.check_call([sys.executable, "-m", "pip", "install", "pyperclip"])  # Installe pyperclip
subprocess.check_call([sys.executable, "-m", "pip", "install", "scipy"])  # Installe scipy

# Configuration du thème de l'interface graphique avec customtkinter
ctk.set_appearance_mode("dark")  # Définit le mode sombre pour l'interface
ctk.set_default_color_theme("dark-blue")  # Applique un thème bleu sombre

# Dictionnaire de mappage pour renommer les colonnes de la feuille "BD_Quest" dans le fichier Excel
mapping_bd_quest = {
    "VOLONTAIRE N°": "Id_volontaire",
    "ANNEE DE NAISSANCE": "Annee_naissance",
    "GENRE": "Genre",
    "DATE DE REMPLISSAGE": "Date_remplissage",
    "Où êtes-vous né(e) ?": "Lieu_naissance",
    "Merci de préciser votre pays de naissance": "Pays_naissance",
    "À quel âge êtes-vous arrivé(e) en France ?": "Age_arrivee_france",
    "Quelle est votre situation actuelle par rapport à l'emploi ?": "Situation_emploi",
    "Quelle est votre profession actuelle ou la dernière profession que vous avez exercée ?": "Profession_actuelle",
    "Quel est le diplôme le plus élevé que vous avez obtenu ?": "Diplome",
    "Cochez toutes les activités que vous pratiquez.": "Activites_pratiquees",
    "Avez-vous pratiqué une activité physique ou sportive au cours des 12 derniers mois ?": "Activite_physique_12mois",
    "Quelles sont les activités sportives que vous pratiquez ?": "Sports_pratiques",
    "Vous êtes :": "Situation_familiale",
    "Quel est votre poids actuel en kg ?": "Poids_kg",
    "Quelle est votre taille actuelle en cm ?": "Taille_cm",
    "Sur cette échelle de 1 à 10, en moyenne au cours de la semaine passée, comment vous êtes-vous senti sur le plan physique ?": "Score_physique",
    "Sur cette échelle de 1 à 10, en moyenne au cours de la semaine passée, comment vous êtes-vous senti sur le plan mental ?": "Score_mental",
    "Au cours des 12 derniers mois, avez-vous eu un ou des accidents ?": "Accident_12mois",
    "Quels sont les accidents que vous avez eu ?": "Types_accidents",
    "Souffrez-vous d'une déficience ou d'un handicap ?": "Handicap",
    "Vous êtes :_6": "Situation_sociale",
    "Combien fumez-vous ou fumiez-vous de cigarettes, cigarillos, cigares ou pipes par jour ?": "Consommation_tabac",
    "Avez-vous consommé du cannabis (haschisch, marijuana, herbe, joint, shit) au cours des 30 derniers jours ?": "Cannabis_30jours",
    "Combien de fois, au cours des 30 derniers jours, en avez-vous consommé ?": "Frequence_cannabis",
    "A quelle fréquence consommez-vous de l'alcool (Vin, bière, cidre,apéritif, digestif, …) ?": "Frequence_alcool",
    "Combien de verres contenant de l'alcool consommez-vous un jour typique où vous buvez ?": "Quantite_alcool",
    "Avec quelle fréquence buvez-vous 6 verres ou davantage lors d'une occasion particulière ?": "Alcool_binge",
    "DATE DE REMPLISSAGE.FOYER": "Date_remplissage_foyer",
    "Combien de personnes vivent avec vous dans votre foyer ?": "Personnes_foyer",
    "Comment vivez-vous ?": "Type_foyer",
    "Qui sont les personnes vivant avec vous dans votre foyer ?": "Composition_foyer",
    "Parmi les tranches suivantes, dans laquelle se situe le revenu mensuel net de vo": "Revenu_mensuel",
    "Avez-vous des animaux domestiques ?": "Animaux_domestiques",
    "Quelle est la race de votre(vos) chien(s) ?": "Race_chien",
    "Quel est le code postal de votre commune d'habitation ?": "Code_postal",
    "Sélectionnez votre commune": "Commune",
    "Votre lieu de résidence se trouve en :": "Type_zone",
    "Quel est le type d'habitat de votre voisinage ?": "Type_voisinage",
    "Vous habitez dans :": "Type_logement",
    "Vous occupez le logement en tant que :": "Statut_occupation",
    "A quel étage habitez-vous ?": "Etage_logement",
    "Dans votre logement, combien y a-t-il d'étages  ?": "Nb_etages_logement",
    "Quelle est la surface de votre logement ?": "Surface_logement",
    "Merci de préciser la surface exacte si vous la connaissez": "Surface_logement_exacte",
    "Combien avez-vous de pièces dans votre logement ?": "Nb_pieces_logement",
    "Y a-t-il des escaliers à l'intérieur de votre logement ?": "Escaliers_interieur",
    "Y a-t-il des escaliers à l'extérieur de votre logement ?": "Escaliers_exterieur",
    "Votre logement dispose-t-il d'un grenier ?": "Presence_grenier",
    "Votre logement dispose-t-il d'une cave ?": "Presence_cave",
    "Votre logement dispose-t-il d'un ou plusieurs balcons ?": "Presence_balcon",
    "De quel type est le chauffage principal de votre logement ?": "Chauffage_principal",
    "Quelles sont les sources d'énergie du chauffage de votre logement ?": "Sources_energie",
    "De quel(s) appareils de chauffage disposez-vous ?": "Appareils_chauffage",
    "Êtes-vous équipés de détecteur de fumée ?": "Detecteur_fumee",
    "Êtes-vous équipés de détecteur de monoxyde de carbone ?": "Detecteur_monoxyde",
    "Êtes-vous équipés d'extincteur dans votre logement ?": "Extincteur",
    "Disposez-vous d'un box ou d'un garage/box ?": "Garage",
    "Disposez-vous d’un espace extérieur (jardin, terrain, cour…)  ?": "Espace_exterieur",
    "Quelle est la surface de cet espace extérieur ?": "Surface_exterieur",
    "Merci de préciser la surface exacte si vous la connaissez_1": "Surface_exterieur_exacte",
    "Avez-vous un abri ou une cabane de jardin ?": "Abri_jardin",
    "Y a -t-il un plan d'eau et/ou une piscine dans votre espace extérieur ?": "Plan_eau_ou_piscine"
}

# Dictionnaire de mappage pour renommer les colonnes de la feuille "Accident" dans le fichier Excel
mapping_accident = {
    "VOLONTAIRE N°": "Id_volontaire",
    "ANNEE DE NAISSANCE": "Annee_naissance",
    "GENRE": "Genre",
    "DATE DE REMPLISSAGE": "Date_remplissage",
    "Dans les deux dernier mois, avez-vous été victime d'un accident de la vie courante ?": "Accident_2mois",
    "À quelle date a eu lieu l'accident de la vie courante ?": "Date_accident",
    "À quelle heure ?": "Heure_accident",
    "Un tiers est-il partiellement ou entièrement responsable de l'accident ?": "Tiers_responsable",
    "Dans quel état de fatigue vous sentiez-vous au moment de l'accident ?": "Fatigue_avant_accident",
    "Quel est le code postal du lieu de l'accident ?": "Code_postal_accident",
    "Sélectionnez la commune de l'accident": "Commune_accident",
    "Où a eu lieu l'accident ?": "Lieu_accident",
    "Précisez le lieu de l'accident :": "Zone_de_transport",
    "Précisez le lieu de l'accident :_1": "Habitat_dans_mon_propre_foyer",
    "Précisez le lieu de l'accident :_2": "Habitat_chez_une_autre_personne",
    "Précisez le lieu de l'accident :_6": "Aire_de_sport",
    "Précisez le lieu de l'accident :_7": "Equipement_de_loisirs_et_de_divertissements_parc",
    "Précisez le lieu de l'accident :_8": "Pleine_nature",
    "Précisez le lieu de l'accident :_9": "Mer_lac_et_riviere",
    "Que faisiez-vous au moment de l'accident ?": "Activite_avant_accident",
    "Précisez l'activité pratiquée :": "Activites_domestiques_menageres_jardinage",
    "Précisez l'activité pratiquée :_10": "Travaux_de_bricolage",
    "Précisez l'activité pratiquée :_11": "Jeu_et_activite_de_loisirs",
    "Précisez l'activité pratiquée :_12": "Activite_sportive",
    "Précisez l'activité pratiquée :_13": "Activite_vitale",
    "Précisez l'activité pratiquée :_14": "Marcher_se_deplacer",
    "Quel sport pratiquiez-vous au moment de l'accident ?": "Sport_lie_accident",
    "Précisez le sport pratiqué :": "Sports_d_escalade",
    "De quel type d'accident s'agissait-il ?": "Type_accident",
    "Dans quelle direction êtes-vous tombé(e) ?": "Direction_chute",
    "Précisez d'où vous êtes tombée(e) :": "Origine_chute",
    "Vous êtes-vous blessé(e) dans l'accident ?": "Blesse",
    "Quelle(s) blessure(s) l'accident a-t-il provoqué ?": "Blessures",
    "Avez-vous reçu des soins après l'accident ?": "Soins_post_accident",
    "Par qui avez-vous reçu ces soins ?": "Dispensateur_soins",
    "Combien de jours avez-vous été hospitalisé(e) ?": "Duree_hospitalisation",
    "Êtes-vous toujours hospitalisé(e) ?": "Hospitalisation_en_cours",
    "Avez-vous été en arrêt de travail suite à cet accident ?": "Arret_travail",
    "Combien de jours (consécutifs) avez-vous été en arrêt de travail ?": "Duree_arret_travail",
    "Êtes-vous toujours en arrêt de travail ?": "Arret_travail_en_cours",
    "Au cours des 48 heures qui ont suivi l'accident, avez-vous été limité(e) dans vos activités habituelles ?": "Limitation_activites_48h",
    "L'accident a-t-il entraîné un arrêt de la pratique de sport (entraînement ou compétition) ?": "Arret_sport",
    "Combien de temps a duré cet arrêt de la pratique de sport ?": "Duree_arret_sport",
    "Êtes-vous toujours en arrêt de la pratique de sport ?": "Arret_sport_en_cours",
    "Précisez par quoi (allergie, intoxication ou corrosion) :": "All_intox_corro",
    "Le surmenage est-il arrivé au cours d'un :": "surmenage",
    "Pouvez-vous décrire  en quelques mots le déroulement de l'accident et ses conséquences ?": "Description_accident"
}

def normaliser_nom_pays(country):
    """Uniformise la casse des noms de pays en style 'title case'."""
    # Vérifie si la valeur est nulle, non textuelle ou vide
    if pd.isna(country) or not isinstance(country, str) or country.strip() == "":
        return country
    # Supprime les espaces inutiles et met en majuscule la première lettre de chaque mot
    return country.strip().title()

def normaliser_taille_cm(height):
    """Normalise la taille en centimètres, convertissant les mètres si nécessaire."""
    # Si la valeur est nulle, retourne NA
    if pd.isna(height):
        return pd.NA
    try:
        # Convertit la valeur en chaîne, supprime les espaces
        height_str = str(height).replace(" ", "")
        # Remplace les virgules par des points pour gérer les formats décimaux
        if "," in height_str:
            height_str = height_str.replace(",", ".")
        height_value = float(height_str)
        # Convertit les tailles en mètres (ex. 1.75) en centimètres (175)
        if 0 < height_value < 3:
            height_cm = height_value * 100
        else:
            height_cm = height_value
        # Vérifie si la taille est dans une plage réaliste (50-250 cm)
        if 50 <= height_cm <= 250:
            return round(height_cm, 1)
        return pd.NA
    except ValueError:
        return pd.NA  # Retourne NA en cas d'erreur de conversion

def nettoyer_chaine_json(val):
    """Nettoie et parse une chaîne JSON-like."""
    # Vérifie si la valeur est nulle ou non textuelle
    if pd.isna(val) or not isinstance(val, str):
        return []
    # Nettoie les caractères spéciaux (accents) et ajuste le format JSON
    val = val.strip().replace("\\u00e9", "é").replace("\\u00e0", "à").replace("\\u00e8", "è")
    if not val.startswith('['):
        val = f'[{val}'  # Ajoute une ouverture de liste si absente
    if not val.endswith(']'):
        val = f'{val}]'  # Ajoute une fermeture de liste si absente
    try:
        return json.loads(val)  # Tente de parser la chaîne JSON
    except json.JSONDecodeError as e:
        print(f"Erreur de parsing JSON : {e} pour la valeur : {val}")
        # Remplace les guillemets simples par des doubles pour corriger le JSON
        val = re.sub(r"'", '"', val)
        try:
            return json.loads(val)
        except json.JSONDecodeError:
            print(f"Échec du parsing JSON après tentative de correction : {val}")
            return []  # Retourne une liste vide en cas d'échec

def extraire_info_activites(val):
    """Extrait les activités sportives d'une chaîne JSON."""
    # Nettoie la chaîne JSON et récupère les données
    data = nettoyer_chaine_json(val)
    if not data:
        return pd.Series({'Nombre_activites': 0})  # Retourne un compteur à 0 si aucune donnée
    result = {}
    # Parcourt chaque activité pour extraire les informations
    for i, activite in enumerate(data, 1):
        result[f'Type_activite_{i}'] = activite.get("Type d'activité")
        result[f'Activite_precise_{i}'] = activite.get("Activité précise")
    result['Nombre_activites'] = len(data)  # Compte le nombre d'activités
    return pd.Series(result)

def extraire_info_foyer(val):
    """Extrait la composition du foyer d'une chaîne JSON, sans gérer les animaux domestiques."""
    # Nettoie la chaîne JSON et récupère les données
    data = nettoyer_chaine_json(val)
    if not data or not isinstance(data, list):
        return pd.Series({'Nombre_personnes_foyer': 0})  # Retourne un compteur à 0 si aucune donnée
    result = {}
    # Parcourt chaque personne pour extraire les informations
    for i, person in enumerate(data, 1):
        result[f'Sexe_personne_{i}'] = person.get('Sexe')
        result[f'Type_personne_{i}'] = person.get('Type de personne')
        result[f'Annee_naissance_personne_{i}'] = person.get('Année de naissance')
        result[f'Occupation_personne_{i}'] = person.get('Occupation du logement')
        result[f'Inscription_MAVIE_personne_{i}'] = person.get('Est-elle inscrite à MAVIE ?')
    result['Nombre_personnes_foyer'] = len(data)  # Compte le nombre de personnes
    return pd.Series(result)

def extraire_info_loisirs(val):
    """Extrait les activités de loisirs d'une chaîne séparée par des virgules."""
    # Si la valeur est nulle ou non textuelle, retourne un dictionnaire par défaut
    if pd.isna(val) or not isinstance(val, str):
        return pd.Series({
            'Activite_Ecran': 'Non', 'Activite_Lecture': 'Non', 'Activite_Jeu_Interieur': 'Non',
            'Activite_Jeu_Exterieur': 'Non', 'Activite_Jardinage': 'Non', 'Activite_Bricolage': 'Non',
            'Activite_Menage': 'Non', 'Activite_Sorties': 'Non', 'Activite_Artistique': 'Non',
            'Activite_Manuelle': 'Non', 'Activite_Sportive': 'Non', 'Activite_Autre': 'Non',
            'Activite_Aucune': 'Non'
        })
    # Sépare la chaîne en liste d'activités
    activities = [act.strip() for act in val.split(',')]
    # Vérifie la présence de chaque type d'activité et attribue "Oui" ou "Non"
    result = {
        'Activite_Ecran': 'Oui' if any('Activités avec un écran' in act for act in activities) else 'Non',
        'Activite_Lecture': 'Oui' if any('Lecture' in act for act in activities) else 'Non',
        'Activite_Jeu_Interieur': 'Oui' if any('Jeu intérieur' in act for act in activities) else 'Non',
        'Activite_Jeu_Exterieur': 'Oui' if any('Jeu extérieur' in act for act in activities) else 'Non',
        'Activite_Jardinage': 'Oui' if any('Jardinage' in act for act in activities) else 'Non',
        'Activite_Bricolage': 'Oui' if any('Bricolage' in act for act in activities) else 'Non',
        'Activite_Menage': 'Oui' if any('Activités ménagères' in act for act in activities) else 'Non',
        'Activite_Sorties': 'Oui' if any('Sorties' in act for act in activities) else 'Non',
        'Activite_Artistique': 'Oui' if any('Activités artistiques' in act for act in activities) else 'Non',
        'Activite_Manuelle': 'Oui' if any('Activités manuelles' in act for act in activities) else 'Non',
        'Activite_Sportive': 'Oui' if any('Activité sportive' in act for act in activities) else 'Non',
        'Activite_Autre': 'Oui' if any('Autre activité' in act for act in activities) else 'Non',
        'Activite_Aucune': 'Oui' if any('Aucune activité' in act for act in activities) else 'Non'
    }
    return pd.Series(result)

def extraire_info_handicap(val):
    """Extrait les informations sur les handicaps d'une chaîne séparée par des virgules."""
    # Si la valeur est nulle ou non textuelle, retourne un dictionnaire par défaut
    if pd.isna(val) or not isinstance(val, str):
        return pd.Series({
            'Handicap_Aucun': 'Non', 'Handicap_Deplacement': 'Non', 'Handicap_Vision': 'Non',
            'Handicap_Daltonisme': 'Non', 'Handicap_Audition': 'Non', 'Handicap_Mental': 'Non',
            'Handicap_Autre': 'Non'
        })
    # Sépare la chaîne en liste de handicaps
    handicaps = [h.strip() for h in val.split(',')]
    # Vérifie la présence de chaque type de handicap et attribue "Oui" ou "Non"
    result = {
        'Handicap_Aucun': 'Oui' if any('Aucune' in h for h in handicaps) else 'Non',
        'Handicap_Deplacement': 'Oui' if any('déplacements' in h.lower() for h in handicaps) else 'Non',
        'Handicap_Vision': 'Oui' if any('vision' in h.lower() and 'couleurs' not in h.lower() for h in handicaps) else 'Non',
        'Handicap_Daltonisme': 'Oui' if any('daltonisme' in h.lower() or 'couleurs' in h.lower() for h in handicaps) else 'Non',
        'Handicap_Audition': 'Oui' if any('audition' in h.lower() for h in handicaps) else 'Non',
        'Handicap_Mental': 'Oui' if any('mentale' in h.lower() or 'intellectuelle' in h.lower() or 'psychologique' in h.lower() for h in handicaps) else 'Non',
        'Handicap_Autre': 'Oui' if any('autre' in h.lower() for h in handicaps) else 'Non'
    }
    return pd.Series(result)

def extraire_info_difficultes(val):
    """Extrait les difficultés de mouvement quotidiennes d'une chaîne séparée par des virgules."""
    # Si la valeur est nulle ou non textuelle, retourne un dictionnaire par défaut
    if pd.isna(val) or not isinstance(val, str):
        return pd.Series({
            'Difficulte_Aucun': 'Non', 'Difficulte_Toilette': 'Non', 'Difficulte_Habillage': 'Non',
            'Difficulte_Manger': 'Non', 'Difficulte_LeversAsseoir': 'Non', 'Difficulte_Deplacement_Interieur': 'Non',
            'Difficulte_Escalier': 'Non', 'Difficulte_Deplacement_Exterieur': 'Non', 'Difficulte_Marcher': 'Non',
            'Difficulte_Courses': 'Non', 'Difficulte_Porter': 'Non', 'Difficulte_Menage': 'Non',
            'Difficulte_Autre': 'Non'
        })
    # Sépare la chaîne en liste de difficultés
    difficulties = [d.strip() for d in val.split(',')]
    # Vérifie la présence de chaque type de difficulté et attribue "Oui" ou "Non"
    result = {
        'Difficulte_Aucun': 'Oui' if any('Aucun' in d for d in difficulties) else 'Non',
        'Difficulte_Toilette': 'Oui' if any('toilette' in d.lower() or 'hygiène' in d.lower() for d in difficulties) else 'Non',
        'Difficulte_Habillage': 'Oui' if any('habiller' in d.lower() or 'déshabiller' in d.lower() for d in difficulties) else 'Non',
        'Difficulte_Manger': 'Oui' if any('manger' in d.lower() or 'boire' in d.lower() for d in difficulties) else 'Non',
        'Difficulte_LeversAsseoir': 'Oui' if any('lever' in d.lower() or 'asseoir' in d.lower() for d in difficulties) else 'Non',
        'Difficulte_DeplacementInterieur': 'Oui' if any('déplacer à l\'intérieur' in d.lower() for d in difficulties) else 'Non',
        'Difficulte_Escalier': 'Oui' if any('escalier' in d.lower() for d in difficulties) else 'Non',
        'Difficulte_DeplacementExterieur': 'Oui' if any('déplacer à l\'extérieur' in d.lower() for d in difficulties) else 'Non',
        'Difficulte_Marcher': 'Oui' if any('marcher' in d.lower() or 'kilomètres' in d.lower() for d in difficulties) else 'Non',
        'Difficulte_Courses': 'Oui' if any('courses' in d.lower() for d in difficulties) else 'Non',
        'Difficulte_Porter': 'Oui' if any('porter' in d.lower() or 'objets lourds' in d.lower() for d in difficulties) else 'Non',
        'Difficulte_Menage': 'Oui' if any('ménagères' in d.lower() or 'cuisine' in d.lower() or 'vaisselle' in d.lower() or 'lessive' in d.lower() or 'ménage' in d.lower() for d in difficulties) else 'Non',
        'Difficulte_Autre': 'Oui' if any('autre' in d.lower() for d in difficulties) else 'Non'
    }
    return pd.Series(result)

def calculer_risque_relatif(tableau_contingence, category1, category2):
    """Calcule le risque relatif entre deux catégories."""
    try:
        # Extrait les valeurs du tableau de contingence pour les deux catégories
        a = tableau_contingence.loc[category1, "Oui"]
        b = tableau_contingence.loc[category1, "Non"]
        c = tableau_contingence.loc[category2, "Oui"]
        d = tableau_contingence.loc[category2, "Non"]
        # Vérifie si les dénominateurs sont non nuls
        if b == 0 or d == 0:
            return None
        # Calcule les risques pour chaque catégorie
        risk1 = a / (a + b)
        risk2 = c / (c + d)
        # Retourne le risque relatif si risk2 est non nul
        return risk1 / risk2 if risk2 != 0 else None
    except:
        return None  # Retourne None en cas d'erreur

def normaliser_nom_colonne(col):
    """Normalise les noms de colonnes."""
    # Supprime les espaces et caractères de nouvelle ligne pour uniformiser les noms
    return col.strip().replace("\n", " ").replace("\r", " ")

# Variable globale pour stocker le chemin du fichier nettoyé
chemin_fichier_nettoye = None

def nettoyer_fichier(filepath, progress_bar, label_status, root):
    """Nettoie le fichier Excel, effectue le test Chi² automatiquement et enregistre les fichiers de sortie."""
    global chemin_fichier_nettoye
    # Vérifie si un fichier a été sélectionné
    if not filepath:
        root.after(0, lambda: label_status.configure(text="Aucun fichier sélectionné."))
        return None
    try:
        # Affiche la barre de progression et le message de statut
        root.after(0, lambda: progress_bar.start())
        root.after(0, lambda: label_status.configure(text="Nettoyage en cours..."))

        # Charge le fichier Excel
        xls = pd.ExcelFile(filepath)
        required_sheets = ["BD_Quest", "Accident"]
        # Vérifie si les feuilles requises sont présentes
        if not all(sheet in xls.sheet_names for sheet in required_sheets):
            root.after(0, lambda: label_status.configure(text=f"Erreur : Les feuilles {required_sheets} sont requises."))
            return None

        # Charge les DataFrames et normalise les noms des colonnes
        df_quest = xls.parse("BD_Quest")
        df_quest.columns = [normaliser_nom_colonne(col) for col in df_quest.columns]

        df_accident = xls.parse("Accident")
        df_accident.columns = [normaliser_nom_colonne(col) for col in df_accident.columns]

        # Nettoyage de la feuille BD_Quest
        if "ANNEE DE NAISSANCE" in df_quest.columns:
            # Convertit les années en nombres, filtre les valeurs aberrantes (1900-2025)
            df_quest["ANNEE DE NAISSANCE"] = pd.to_numeric(df_quest["ANNEE DE NAISSANCE"], errors="coerce")
            df_quest = df_quest[df_quest["ANNEE DE NAISSANCE"].between(1900, 2025)]
            df_quest["ANNEE DE NAISSANCE"] = df_quest["ANNEE DE NAISSANCE"].astype("Int64")
        else:
            print("Colonne 'ANNEE DE NAISSANCE' manquante dans BD_Quest.")

        if "Merci de préciser votre pays de naissance" in df_quest.columns:
            # Normalise les noms de pays
            df_quest["Merci de préciser votre pays de naissance"] = df_quest["Merci de préciser votre pays de naissance"].apply(normaliser_nom_pays)

        if "Quelle est votre taille actuelle en cm ?" in df_quest.columns:
            # Normalise les tailles en centimètres
            df_quest["Quelle est votre taille actuelle en cm ?"] = df_quest["Quelle est votre taille actuelle en cm ?"].apply(normaliser_taille_cm)

        # Traitement des colonnes JSON et autres
        if "Cochez toutes les activités que vous pratiquez." in df_quest.columns:
            # Extrait les informations sur les loisirs
            df_loisirs = df_quest["Cochez toutes les activités que vous pratiquez."].apply(extraire_info_loisirs)
            df_quest = pd.concat([df_quest, df_loisirs], axis=1)
            df_quest = df_quest.drop(columns=["Cochez toutes les activités que vous pratiquez."], errors='ignore')

        if "Souffrez-vous d'une déficience ou d'un handicap ?" in df_quest.columns:
            # Extrait les informations sur les handicaps
            df_handicap = df_quest["Souffrez-vous d'une déficience ou d'un handicap ?"].apply(extraire_info_handicap)
            df_quest = pd.concat([df_quest, df_handicap], axis=1)
            df_quest = df_quest.drop(columns=["Souffrez-vous d'une déficience ou d'un handicap ?"], errors='ignore')

        # Combine les colonnes de difficultés si elles existent
        difficultes_cols = [col for col in df_quest.columns if col.startswith("Pour quel(s) mouvement(s) de la vie quotidienne présentez-vous des difficultés ?")]
        if difficultes_cols:
            combined_difficultes = df_quest[difficultes_cols].bfill(axis=1).iloc[:, 0]
            df_difficultes = combined_difficultes.apply(extraire_info_difficultes)
            df_quest = pd.concat([df_quest, df_difficultes], axis=1)
            df_quest = df_quest.drop(columns=difficultes_cols, errors='ignore')

        if "Quelles sont les activités sportives que vous pratiquez ?" in df_quest.columns:
            # Extrait les informations sur les activités sportives
            df_activites = df_quest["Quelles sont les activités sportives que vous pratiquez ?"].apply(extraire_info_activites)
            df_quest = pd.concat([df_quest, df_activites], axis=1)
            df_quest = df_quest.drop(columns=["Quelles sont les activités sportives que vous pratiquez ?"], errors='ignore')

        if "Qui sont les personnes vivant avec vous dans votre foyer ?" in df_quest.columns:
            # Extrait les informations sur la composition du foyer
            df_foyer = df_quest["Qui sont les personnes vivant avec vous dans votre foyer ?"].apply(extraire_info_foyer)
            df_quest = pd.concat([df_quest, df_foyer], axis=1)
            df_quest = df_quest.drop(columns=["Qui sont les personnes vivant avec vous dans votre foyer ?"], errors='ignore')

        if "Quels sont les accidents que vous avez eu ?" in df_quest.columns:
            # Nettoie et extrait les types d'accidents
            df_quest["accidents_cleaned"] = df_quest["Quels sont les accidents que vous avez eu ?"].apply(nettoyer_chaine_json)
            df_quest["type_accident"] = df_quest["accidents_cleaned"].apply(lambda x: x[0].get("Type d'accident") if x and isinstance(x, list) and len(x) > 0 else None)
            df_quest["accident_12mois"] = df_quest["type_accident"].apply(lambda x: "Oui" if pd.notna(x) else "Non")
            df_quest = df_quest.drop(columns=["Quels sont les accidents que vous avez eu ?", "accidents_cleaned"], errors='ignore')

        if "De quel type d'accident s'agissait-il ?" in df_accident.columns:
            # Nettoie et extrait les types d'accidents dans la feuille Accident
            df_accident["accident_type_cleaned"] = df_accident["De quel type d'accident s'agissait-il ?"].apply(nettoyer_chaine_json)
            df_accident["type_accident_derive"] = df_accident["accident_type_cleaned"].apply(lambda x: x[0].get("Type d'accident") if x and isinstance(x, list) and len(x) > 0 else None)
            df_accident = df_accident.drop(columns=["De quel type d'accident s'agissait-il ?", "accident_type_cleaned"], errors='ignore')

        if "Quelle(s) blessure(s) l'accident a-t-il provoqué ?" in df_accident.columns:
            # Nettoie et extrait les blessures dans la feuille Accident
            df_accident["blessures_cleaned"] = df_accident["Quelle(s) blessure(s) l'accident a-t-il provoqué ?"].apply(nettoyer_chaine_json)
            df_accident["blessure_derive"] = df_accident["blessures_cleaned"].apply(lambda x: x[0].get("Blessure") if x and isinstance(x, list) and len(x) > 0 else None)
            df_accident = df_accident.drop(columns=["Quelle(s) blessure(s) l'accident a-t-il provoqué ?", "blessures_cleaned"], errors='ignore')

        if "ANNEE DE NAISSANCE" in df_accident.columns:
            # Nettoie les années de naissance dans la feuille Accident
            df_accident["ANNEE DE NAISSANCE"] = pd.to_numeric(df_accident["ANNEE DE NAISSANCE"], errors="coerce")
            df_accident = df_accident[df_accident["ANNEE DE NAISSANCE"].between(1900, 2025)]
            df_accident["ANNEE DE NAISSANCE"] = df_accident["ANNEE DE NAISSANCE"].astype("Int64")

        # Renomme les colonnes selon les dictionnaires de mappage
        df_quest.rename(columns=mapping_bd_quest, inplace=True)
        df_accident.rename(columns=mapping_accident, inplace=True)

        # Définit le chemin du fichier de sortie
        dossier = os.path.dirname(filepath)
        nom_fichier_sortie = os.path.splitext(os.path.basename(filepath))[0] + "_nettoye.xlsx"
        chemin_fichier_nettoye = os.path.join(dossier, nom_fichier_sortie)

        # Sauvegarde les DataFrames nettoyés dans un nouveau fichier Excel
        try:
            with pd.ExcelWriter(chemin_fichier_nettoye, engine="openpyxl") as writer:
                df_quest.to_excel(writer, sheet_name="bd_quest", index=False)
                df_accident.to_excel(writer, sheet_name="accident", index=False)
        except PermissionError:
            root.after(0, lambda: label_status.configure(text="Erreur : Permission refusée pour écrire le fichier nettoyé."))
            return None
        except Exception as e:
            root.after(0, lambda: label_status.configure(text=f"Erreur lors de l'écriture du fichier nettoyé : {str(e)}"))
            return None

        # Effectue l'analyse statistique Chi²
        tester_chi2_animaux_powerbi()

        # Met à jour l'interface avec le chemin du fichier nettoyé
        root.after(0, lambda: label_status.configure(text=f"Fichier nettoyé sauvegardé :\n{chemin_fichier_nettoye}"))
        return chemin_fichier_nettoye

    except pd.errors.EmptyDataError:
        # Gère le cas où le fichier Excel est vide ou corrompu
        root.after(0, lambda: label_status.configure(text="Erreur : Le fichier Excel est vide ou corrompu."))
        return None
    except Exception as e:
        # Gère les autres erreurs potentielles
        root.after(0, lambda: label_status.configure(text=f"Erreur : {str(e)}"))
        return None
    finally:
        # Arrête la barre de progression
        root.after(0, lambda: progress_bar.stop())
        root.after(0, lambda: progress_bar.pack_forget())

def tester_chi2_animaux_powerbi():
    """Effectue un test Chi² et Fisher, et exporte les résultats pour Power BI."""
    global chemin_fichier_nettoye
    # Vérifie si un fichier nettoyé est disponible
    if not chemin_fichier_nettoye:
        root.after(0, lambda: label_status.configure(text="Aucun fichier nettoyé disponible pour l’analyse."))
        return
    
    try:
        # Charge la feuille bd_quest du fichier nettoyé
        df_quest = pd.read_excel(chemin_fichier_nettoye, sheet_name="bd_quest")
        
        # Vérifie la présence des colonnes nécessaires
        required_columns = ["accident_12mois", "Animaux_domestiques"]
        missing_cols = [col for col in required_columns if col not in df_quest.columns]
        if missing_cols:
            root.after(0, lambda: label_status.configure(text=f"Erreur : Colonnes manquantes {missing_cols}."))
            return
        
        def categoriser_animaux(value):
            # Simplifie la colonne des animaux domestiques en deux catégories
            if pd.isna(value) or value == "Aucun":
                return "Aucun"
            else:
                return "Avec animaux"
        
        # Applique la catégorisation
        df_quest["Animaux_Domestiques_Simplifie"] = df_quest["Animaux_domestiques"].apply(categoriser_animaux)
        df_clean = df_quest[["accident_12mois", "Animaux_Domestiques_Simplifie"]].dropna()
        
        # Crée un tableau de contingence
        tableau_contingence = pd.crosstab(df_clean["Animaux_Domestiques_Simplifie"], df_clean["accident_12mois"])
        # Effectue le test Chi²
        chi2_stat, p_valeur, dof, expected = chi2_contingency(tableau_contingence)
        
        # Calcule les proportions
        proportions = tableau_contingence.div(tableau_contingence.sum(axis=1), axis=0) * 100
        # Calcule le risque relatif
        rr = calculer_risque_relatif(tableau_contingence, "Avec animaux", "Aucun")
        # Effectue le test de Fisher
        odds_ratio, p_valeur_fisher = fisher_exact(tableau_contingence)
        
        # Définit le chemin pour le fichier de résultats
        dossier = os.path.dirname(chemin_fichier_nettoye)
        chemin_resultats = os.path.join(dossier, "chi2_animaux_results_powerbi.xlsx")
        # Sauvegarde les résultats dans un fichier Excel
        with pd.ExcelWriter(chemin_resultats, engine="openpyxl") as writer:
            tableau_contingence.to_excel(writer, sheet_name="Contingence")
            proportions.to_excel(writer, sheet_name="Proportions")
            stats_df = pd.DataFrame({
                "Test": ["Chi²", "Fisher"],
                "Statistique": [chi2_stat, odds_ratio],
                "P-valeur": [p_valeur, p_valeur_fisher],
                "Degrés de liberté": [dof, None],
                "Risque Relatif (Avec animaux vs Aucun)": [rr, None]
            })
            stats_df.to_excel(writer, sheet_name="Statistiques", index=False)
        
        # Met à jour l'interface avec les résultats
        root.after(0, lambda: label_status.configure(text=f"Fichier nettoyé sauvegardé :\n{chemin_fichier_nettoye}\nAnalyse Chi² terminée. Résultats exportés dans :\n{chemin_resultats}"))
        
    except Exception as e:
        # Gère les erreurs lors de l'analyse
        root.after(0, lambda: label_status.configure(text=f"Fichier nettoyé sauvegardé :\n{chemin_fichier_nettoye}\nErreur lors de l’analyse Chi² : {str(e)}"))

def parcourir_et_nettoyer():
    """Sélectionne et nettoie un fichier Excel dans un thread séparé."""
    global chemin_fichier_nettoye
    # Ouvre une boîte de dialogue pour sélectionner un fichier Excel
    filepath = filedialog.askopenfilename(filetypes=[("Fichiers Excel", "*.xlsx *.xls")])
    if filepath:
        # Affiche la barre de progression
        progress_bar.pack(anchor="w", pady=5)
        # Lance le nettoyage dans un thread séparé pour ne pas bloquer l'interface
        threading.Thread(target=lambda: nettoyer_fichier(filepath, progress_bar, label_status, root), daemon=True).start()
        # Vérifie régulièrement si le nettoyage est terminé
        root.after(100, verifier_completion_nettoyage)

def verifier_completion_nettoyage():
    """Vérifie si le nettoyage est terminé et met à jour l'interface."""
    global chemin_fichier_nettoye
    if chemin_fichier_nettoye:
        # Met à jour l'interface avec le chemin du dossier
        chemin_dossier = os.path.dirname(chemin_fichier_nettoye) + "/"
        file_path_var.set(chemin_dossier)
        copy_button.configure(state="normal")  # Active le bouton de copie
        launch_button.configure(state="normal")  # Active le bouton de lancement
    else:
        file_path_var.set("")
        copy_button.configure(state="disabled")  # Désactive le bouton de copie
        launch_button.configure(state="disabled")  # Désactive le bouton de lancement
    # Relance la vérification après 100 ms
    root.after(100, verifier_completion_nettoyage)

def copier_dans_presse_papiers():
    """Copie le chemin du dossier dans le presse-papiers."""
    chemin_dossier = os.path.dirname(file_path_var.get())
    if chemin_dossier:
        # Copie le chemin dans le presse-papiers
        pyperclip.copy(chemin_dossier + "/")
        label_status.configure(text="Chemin du dossier copié dans le presse-papiers.")
    else:
        label_status.configure(text="Erreur : Aucun chemin à copier.")

def lancer_visualisation():
    """Lance le rapport Power BI."""
    try:
        # Définit le chemin du fichier Power BI
        pbit_path = os.path.abspath("Rapport_AcVC.pbit")
        if not os.path.exists(pbit_path):
            label_status.configure(text="Erreur : rapport introuvable.")
            return
        # Ouvre le fichier en fonction du système d'exploitation
        if platform.system() == "Windows":
            os.startfile(pbit_path)
        elif platform.system() == "Darwin":
            subprocess.run(["open", pbit_path], check=True)
        else:
            subprocess.run(["xdg-open", pbit_path], check=True)
        label_status.configure(text="Lancement de Power BI...")
    except Exception as e:
        label_status.configure(text=f"Erreur lors du lancement de Power BI : {str(e)}")

# Configuration de l'interface graphique
root = ctk.CTk()  # Crée la fenêtre principale
root.title("Nettoyage et Visualisation Power BI")  # Définit le titre
root.geometry("600x400")  # Définit la taille de la fenêtre
root.resizable(False, False)  # Empêche le redimensionnement

file_path_var = ctk.StringVar()  # Variable pour stocker le chemin du dossier

# Crée un cadre pour organiser les éléments de l'interface
frame = ctk.CTkFrame(master=root, corner_radius=10)
frame.pack(pady=30, padx=30, fill="both", expand=True)

# Étiquette et bouton pour importer le fichier
label_status = ctk.CTkLabel(frame, text="1. Importez le fichier Excel à nettoyer :", font=("Segoe UI", 16))
label_status.pack(anchor="w", pady=(10, 5))

import_button = ctk.CTkButton(frame, text="📂 Importer", command=parcourir_et_nettoyer)
import_button.pack(anchor="w")

# Barre de progression (initialement cachée)
progress_bar = ctk.CTkProgressBar(frame, mode="indeterminate")

# Étiquette et champ pour afficher/copier le chemin
label2 = ctk.CTkLabel(frame, text="2. Copiez le chemin :", font=("Segoe UI", 16))
label2.pack(anchor="w", pady=(20, 5))

path_entry = ctk.CTkEntry(frame, textvariable=file_path_var, width=400)
path_entry.pack(anchor="w", pady=5)

copy_button = ctk.CTkButton(frame, text="📋 Copier le chemin", command=copier_dans_presse_papiers, state="disabled")
copy_button.pack(anchor="w", pady=5)

# Étiquette et bouton pour lancer Power BI
label3 = ctk.CTkLabel(frame, text="3. Visualisez vos données :", font=("Segoe UI", 16))
label3.pack(anchor="w", pady=(20, 5))

launch_button = ctk.CTkButton(frame, text="📊 Ouvrir le rapport", command=lancer_visualisation, state="disabled")
launch_button.pack(anchor="w", pady=(0, 10))

# Lancement de la boucle principale de l'interface graphique
root.mainloop()