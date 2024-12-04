# Agrégation de Factures
Python project. Agregate multiples Excel files. Ignore comments in the code it was for helping someone.

# 🇫🇷 Guide perso pour Quiroc

## **Prérequis**

- **Python 3.13.1** installé sur ton Mac.
- **pip** (gestionnaire de paquets Python) installé.

## **Installation des Packages Nécessaires**

- `pandas` : pour manipuler les données et créer le fichier de sortie.
- `openpyxl` : pour lire les fichiers Excel.

### **Étapes d'Installation**

1. **Ouvre le Terminal**

   Terminal dans `Applications` > `Utilitaires` > `Terminal`.

2. **Mettre à Jour pip**

   pip est à jour :

   ```bash
   pip install --upgrade pip
   ```

3. **Installer pandas et openpyxl**

   Exécuter les commandes suivantes :

   ```bash
   pip install pandas
   pip install openpyxl
   ```

## **Structure du Projet**

- **main.py** : le script Python.
- **resources/** : le dossier contenant tous les fichiers Excel à traiter.
- **agrégation_factures.xlsx** : le fichier de sortie généré après l'exécution du script.

## **Comment Utiliser le Script**

1. **Placer les Fichiers Excel**

   Mettre tous les fichiers Excel (.xlsx) dans le dossier `resources`. Peu importe s'ils sont dans des sous-dossiers.

2. **Exécuter le Script**

   Dans le Terminal, naviguez jusqu'au dossier où se trouve `aggregate_bills.py`. Par exemple :

   ```bash
   cd /chemin/vers/votre/projet
   ```

   Puis exécutez le script :

   ```bash
   python main.py
   ```

3. **Vérifier le Résultat**

   Après l'exécution, un nouveau fichier `agrégation_factures.xlsx` sera créé dans le même dossier.

## **Comprendre le Script**

Le script effectue les actions suivantes :

- **Importe les modules nécessaires** :

  ```python
  import os
  import pandas as pd
  from openpyxl import load_workbook
  ```

- **Définit le chemin du dossier `resources`** :

  ```python
  resources_path = 'resources'
  ```

- **Parcourt tous les fichiers dans `resources`** :

  ```python
  for root, dirs, files in os.walk(resources_path):
      # ...
  ```

- **Vérifie si chaque fichier est un fichier Excel et le traite** :

  ```python
  if file.endswith('.xlsx'):
      # ...
  ```

- **Charge le fichier Excel et extrait les données des cellules spécifiques** :

  ```python
  workbook = load_workbook(filename=file_path, data_only=True)
  sheet = workbook.active
  nom = sheet['A9'].value or ''
  # ...
  ```

- **Stocke les données dans une liste de dictionnaires** :

  ```python
  data_list.append({
      'Nom': nom,
      'Prénom': prenom,
      # ...
  })
  ```

- **Crée un DataFrame pandas et l'enregistre dans un nouveau fichier Excel** :

  ```python
  df = pd.DataFrame(data_list)
  df.to_excel('agrégation_factures.xlsx', index=False)
  ```

---

J'ai mis des commentaires en francais dans le code.
