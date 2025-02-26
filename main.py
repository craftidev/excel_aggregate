# Besoin de os pour utiliser le systeme de parcours de dossier/fichier
import os
# Pandas va pouvoir stocker toutes nos donnees et les convertir pour excel
import pandas as pd
# Utiliser datetime pour travailler avec des dates et heures
from datetime import datetime
# Importer openpyxl pour lire les fichiers Excel
from openpyxl import load_workbook



# Chemin du dossier 'resources'
resources_path = 'resources'
# Liste pour stocker les données extraites
data_list = []
# Variables pour le comptage des fichiers
files_processed = 0
# Des listes pour se rappeler des fichiers qui n'ont pas pu etre traites
files_skipped = []
files_empty = []



# Fonction pour nettoyer les valeurs extraites.
# On declare la fonction que l'on pourra utiliser plus tard.
# Une fonction peut avoir des arguments et des retours.
# Un argument est une variable lu par la fonction.
# Un retour est une variable retournee par la fonction.
# Ici, la fonction aura besoin de la valeur de la cellule.
# Argument `cell_value`
# Puis elle rendra cette valeure nettoyee.
# Ce que la fonction retourne est indique par le mot clef `return`
def sanitize_value(cell_value):
    if cell_value:
        # Convertir la valeur en chaîne de caractères
        value_str = str(cell_value)
        # Diviser la chaîne sur le premier ':'
        # Donc `Genre véhicule:SH 150 AD` devient:
        # `Genre véhicule:` + ` SH 150 AD`
        parts = value_str.split(':', 1)
        if len(parts) > 1: # La variable `parts` a bien 2 parties.
            # Prendre la partie après le ':', et supprimer les espaces
            # en debut/fin avec la methode .strip()
            return parts[1].strip()
        else: # Sinon, ca veut dire qu'il n'y avait pas de `:`
            # On retourne la valeur originale
            return value_str.strip()
    else:
        return ''



# Parcourir tous les fichiers et sous-dossiers dans 'resources'
for root, dirs, files in os.walk(resources_path):
    # `for` cree une boucle
    # pour chaque element dans la liste `files` nous allons
    # executer la meme procedure jusqu'a ce que tous les
    # elements soient traites.
    for file in files:
        # Vérifier si le fichier est un fichier Excel (.xlsx)
        if file.endswith('.xlsx'):
            # On combine les 2 adresses ensemble
            # Exemple:
            # root = `C:/``
            # file = `coucou_maman.xlx`
            # donc file_path = `C:/coucou_maman.xlx`
            file_path = os.path.join(root, file)
            # Ecrire un message dans la console
            print(f"Traitement du fichier: {file_path}")

            # Le mot clef `try` veut dire `essaye`
            # Tout ce qui est dans le bloc try doit fonctionner.
            # A la moindre erreur, rien est fait et le script execute
            # la ce qui est dans le bloc `except` plus bas.
            # C'est simplement de la gestion d'erreur sur une partie
            # sensible du script qui peut effectivement merder.
            try:
                # Charger le classeur excel avec data_only=False pour accéder aux formules
                # on en a besoin pour les dates qui merde quand la formule `=TODAY()` a ete utilisee
                workbook = load_workbook(filename=file_path, data_only=False)
                # On chope la feuille active du classeur vu qu'il y en a qu'une
                sheet = workbook.active

                # Extraire et nettoyer les informations des cellules spécifiques
                # voir plus haut quand on a declarer notre fonction perso `sanitize_value`
                path = file_path
                nom = sanitize_value(sheet['A9'].value)
                prenom = sanitize_value(sheet['A10'].value)
                adresse = sanitize_value(sheet['A11'].value)
                localite = sanitize_value(sheet['A12'].value)
                telephone = sanitize_value(sheet['A13'].value)
                genre_vehicule = sanitize_value(sheet['A14'].value)
                plaques = sanitize_value(sheet['A15'].value)
                annee_km = sanitize_value(sheet['A16'].value)

                # Gestion de la date de la facture
                # Vérifier D12 d'abord
                # Si D12 contient une formule ou une valeur on la recupere
                # Sinon on traite normalement avec C12
                # C'est parce que toutes les factures ne sont pas traitees pareilles
                # pour la date.
                cell_D12 = sheet['D12']
                cell_C12 = sheet['C12']
                date_facture = ''

                # Fonction pour obtenir la date de la cellule
                def get_date_from_cell(cell):
                    # Si il y a une quelconque valeur dans l'argument passe dans cette fonction (`cell`)
                    if cell.value:
                        # Si la cellule contient une formule
                        if cell.data_type == 'f':
                            if cell.value == '=TODAY()':
                                # Indiquer que la formule TODAY a ete utilisee, impossible de savoir
                                return 'formule TODAY()'
                            else:
                                # Formule non reconnue, indiquer que la date n'a pas pu être lue
                                return 'formule inconnue'
                        else:
                            # Sinon, faire comme d'habitude, nettoyer et renvoyer la valeure.
                            return sanitize_value(cell.value)
                    # Sinon, si la valeur de l'argument `cell` est vide
                    else:
                        # On retourne du text vide
                        return ''

                # Obtenir la date de D12 ou C12
                date_facture = get_date_from_cell(cell_D12)
                # Apres avoir tester D12, on verifie qu'il y a quelque chose dedans.
                # Sinon, on travaillera avec C12
                if not date_facture or date_facture.startswith('Date non lue'):
                    date_facture = get_date_from_cell(cell_C12)

                # On a finit de traiter toutes les donnees
                # On ajoute tout ca dans notre liste perso
                new_entry = {
                    'Path': path,
                    'Nom': nom,
                    'Prénom': prenom,
                    'Adresse': adresse,
                    'Localité': localite,
                    'Téléphone': telephone,
                    'Genre véhicule': genre_vehicule,
                    'Plaques': plaques,
                    'Année/KM': annee_km,
                    'Date de la facture': date_facture
                }

                # Test si la nouvelle entree est vide, si tout est vide sauf Path,
                # on considere l'erreure en ajoutant a la liste des fichiers vides
                if all(not str(value).strip() for key, value in new_entry.items() if key != 'Path'):
                    print(f"Erreur lors du traitement du fichier {file_path}, donnees vides")
                    files_empty.append(file_path)
                    continue

                data_list.append(new_entry)

                # Incrementation du nombre de fichiers xlsx traite
                files_processed += 1
            # Bloc qui s'active seulement si une erreur survient
            except Exception as e:
                # Un petit message
                print(f"Erreur lors du traitement du fichier {file_path}: {e}")
                # Et on met ce fichier dans notre liste qui se rappelera des fichiers non traite
                # Que l'on rapporte a l'utilisateur a la fin.
                files_skipped.append(file_path)
        else:
            # Fichier non traité (ne se termine pas par .xlsx)
            # Pareil, fichier non traite dans la liste.
            # L'utilisateur saura si un fichier non .xlsx traine par megarde.
            files_skipped.append(os.path.join(root, file))



# Créer un DataFrame pandas avec les données.
# On balance notre liste remplit de toutes les
# donnees lues. Et on va utiliser pandas pour
# refourger ca a Excel.
df = pd.DataFrame(data_list, columns=[
    'Path', 'Nom', 'Prénom', 'Adresse', 'Localité', 'Téléphone',
    'Genre véhicule', 'Plaques', 'Année/KM', 'Date de la facture'
])

# Enregistrer le DataFrame dans un nouveau fichier Excel
df.to_excel('agrégation_factures.xlsx', index=False)

print(f"\nAgrégation terminée. Le fichier 'agrégation_factures.xlsx' a été créé.")
print(f"Nombre de fichiers traités : {files_processed}")

if files_empty: # Si il y a quelque chose dans notre liste de fichier vide
    print("\nLes fichiers suivants sont vide :")
    for empty_file in files_empty:
        print(f"- {empty_file}")
if files_skipped: # Si il y a quelque chose dans notre liste de fichier non traite
    print("\nLes fichiers suivants n'ont pas été traités :")
    for skipped_file in files_skipped:
        print(f"- {skipped_file}")
