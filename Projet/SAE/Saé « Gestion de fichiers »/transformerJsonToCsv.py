"""
Nom du fichier : convertir_json_en_csv.py
Description : Script pour convertir un fichier JSON contenant des informations sur des événements culturels de Paris en un fichier CSV.
Auteurs : Constant Audibert et Paul Babin
Date : 13/12/2024
"""


import json
import csv
import re
import os

# Chemin vers le fichier JSON et le CSV
json_file = 'que-faire-a-paris-.json'
csv_final = 'final.csv'

# On vérifie que le JSON est trouvable
if not os.path.isfile(json_file):
    print(f"\033[91mErreur : Le fichier {json_file} n'est pas dans le dossier courant.\033[0m")
    raise FileNotFoundError(f"Le fichier {json_file} est introuvable. Veuillez vérifier son emplacement.")


# Champs à inclure dans le CSV
champs = ["ID", "URL", "Titre", "Chapeau", "Description", "Date de début", "Heure de début", 
    "Date de fin", "Heure de fin", "Nom du lieu", "Adresse du lieu", "Code Postal", "Ville", 
    "Coordonnées géographiques", "Accès PMR", "Accès mal voyant", "Accès mal entendant", "Transport", "Téléphone de contact", "Email de contact", "Url de contact", "Type d’accès", 
    "Détail du prix", "URL de l’image de couverture"]

# On fait une fonction pour nettoyer les balises html du texte :
def nettoyer_texte(texte):
    if texte is None:
        return ""
    try:
        # On supprime toutes les balises HTML comme <p>, <b>, etc.
        texte = re.sub(r'<[^>]+>', '', texte)
        # On remplace les caractères non encodables par un caractère de remplacement
        texte = texte.encode('utf-8', errors='replace').decode('utf-8')
        # On supprime les balises &nbsp et &gt
        texte = re.sub(r'&nbsp|&gt;', '', texte)
        return texte
    except Exception as e:
        print(f"Erreur lors du nettoyage du texte : {e}")
        return ""

def nettoyer_transport(transport):
    if transport:
        # On Coupe le texte dès qu'une balise <a> est rencontrée
        transport = transport.split('<a')[0]
        # On Supprime toutes les balises HTML restantes
        transport = re.sub(r'<[^>]*>', '', transport)
        # On Nettoie les espaces non désirés
        transport = transport.strip()
        return transport
    return ""



# On fait une fonction qui permet de renvoyer Oui ou non en fonction du type d'accès renseigné   
def o_n(texte):
    if texte == 1:
        return "Oui"
    return "Non"
    
# On ajoute une fonction pour indiquer les directions des coordonées
def format_coordonnées(coordinates):
    if coordinates:
        lon = coordinates.get("lon")
        lat = coordinates.get("lat")
        if lon and lat:
            lon_dir = "E" if lon >= 0 else "O"
            lat_dir = "N" if lat >= 0 else "S"
            return f"{lon_dir} {lon}°, {lat_dir} {lat}°"
    return ""

# On charge le fichier JSON
with open(json_file, 'r', encoding='utf-8') as file:
    data = json.load(file)

# On vérifie si le JSON est une liste ou un dictionnaire
if isinstance(data, dict):  # Si c'est un dictionnaire unique, l'encapsuler dans une liste
    data = [data]

# On convertit les données en CSV
with open(csv_final, mode='w', newline='', encoding='utf-8') as file:
    # On définit l'en-tête avec notre liste
    writer = csv.DictWriter(file, fieldnames=champs, delimiter=';')

    # On ecrit l'en-tête dans le fichier CSV
    writer.writeheader()

    # Parcours et écriture de chaque élément du JSON
    for item in data:
        # Traitement de la date et de l'heure pour la "Date de début" et "Heure de début"
        date_debut = ""
        heure_debut = ""
        if item.get("date_start"):
            try:
                # Séparer la date et l'heure
                date_part, time_part = item["date_start"].split("T")
                # Vérifier si la partie date est correcte
                if len(date_part.split("-")) == 3:
                    jour, mois, annee = date_part.split("-")
                    # Vérifier si la partie heure est correcte
                    if len(time_part.split(":")) >= 2:
                        heure, minute = time_part.split(":")[:2]
                        date_debut = f"{jour}-{mois}-{annee}"
                        heure_debut = f"{heure}:{minute}"
                    else:
                        print(f"Problème avec l'heure dans : {item['date_start']}")
                else:
                    print(f"Problème avec la date dans : {item['date_start']}")
            except Exception as e:
                print(f"Erreur lors de l'extraction de la date/heure de début: {e}")
                pass  # Si une erreur survient, laisse les champs vides

        # On fait la même chose pour la "Date de fin" et "Heure de fin"
        date_fin = ""
        heure_fin = ""
        if item.get("date_end"):
            try:
                # Séparer la date et l'heure
                date_part, time_part = item["date_end"].split("T")
                # Vérifier si la partie date est correcte
                if len(date_part.split("-")) == 3:
                    jour, mois, annee = date_part.split("-")
                    # Vérifier si la partie heure est correcte
                    if len(time_part.split(":")) >= 2:
                        heure, minute = time_part.split(":")[:2]
                        date_fin = f"{jour}-{mois}-{annee}"
                        heure_fin = f"{heure}:{minute}"
                    else:
                        print(f"Problème avec l'heure dans : {item['date_end']}")
                else:
                    print(f"Problème avec la date dans : {item['date_end']}")
            except Exception as e:
                print(f"Erreur lors de l'extraction de la date/heure de fin: {e}")
                pass  # Si une erreur survient, laisse les champs vides

        # Écrire les données dans le fichier CSV en utilisant nos fonctions si besoin
        writer.writerow({
            "ID": item.get("id", ""),
            "URL": item.get("url", ""),
            "Titre": item.get("title", ""),
            "Chapeau": item.get("lead_text", ""),
            "Description": nettoyer_texte(item.get("description", "")),
            "Date de début": date_debut,
            "Heure de début": heure_debut,
            "Date de fin": date_fin,
            "Heure de fin": heure_fin,
            "Nom du lieu": item.get("address_name", ""),
            "Adresse du lieu": item.get("address_street", ""),
            "Code Postal": item.get("address_zipcode", ""),
            "Ville": item.get("address_city", ""),
            "Coordonnées géographiques": format_coordonnées(item.get("lat_lon", "")),
            "Accès PMR": o_n(item.get("pmr", "")),
            "Accès mal voyant": o_n(item.get("blind", "")),
            "Accès mal entendant": o_n(item.get("deaf", "")),
            "Transport": nettoyer_transport(item.get("transport", "")),
            "Téléphone de contact": item.get("contact_phone", ""),
            "Email de contact": item.get("contact_mail", ""),
            "Url de contact": item.get("contact_url", ""),
            "Type d’accès": item.get("access_type", ""),
            "Détail du prix": nettoyer_texte(item.get("price_detail", "")),
            "URL de l’image de couverture": item.get("cover_url", "")
        })

print(f"Les données ont été exportées dans {csv_final}. \nIl est préférable d'importer les données depuis un nouveau classeur Excel (power query) en CSV pour éviter les problèmes d'encodage.")
