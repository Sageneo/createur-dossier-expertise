# Créateur de Dossier d'Expertise

Un outil pratique qui crée automatiquement une structure de dossiers pour les expertises à partir d'un fichier email (.eml), en extrayant les informations pertinentes et en organisant les pièces jointes.

## Fonctionnalités

- Extraction automatique des informations clés d'un email (.eml)
- Création d'une structure de dossiers organisée pour les expertises
- Tri intelligent des pièces jointes (images vers le dossier Photos)
- Interface graphique simple pour sélectionner les fichiers et dossiers
- Renommage automatique des dossiers selon les informations du client

## Installation

### Méthode 1 : Installation directe
1. Assurez-vous d'avoir Python installé (version 3.6 ou supérieure)
2. Téléchargez le fichier `creer_dossier_expertise.py`
3. Aucune bibliothèque externe n'est requise (utilise uniquement les modules standards de Python)

### Méthode 2 : Installation via pip
1. Assurez-vous d'avoir Python installé (version 3.6 ou supérieure)
2. Dans le dossier contenant `setup.py`, exécutez :
```bash
pip install -e .
