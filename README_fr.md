# MediaMetadataParser

![Version Python](https://img.shields.io/badge/python-3.8+-blue.svg)
![Licence](https://img.shields.io/badge/licence-GPLv3-green.svg)

## Motivation

Ce projet a été créé pour aider ma conjointe qui travaille dans la production vidéo et avait besoin d'un moyen plus rapide d'accéder aux métadonnées de grandes collections de rushes, de séquences vidéo et d'autres fichiers multimédias. Les outils existants étaient soit trop lents, ne fournissaient pas les bonnes informations, ou nécessitaient un traitement manuel de chaque fichier. MediaMetadataParser a été conçu pour :

- Analyser rapidement des répertoires entiers de fichiers multimédias
- Extraire toutes les métadonnées techniques pertinentes en une seule fois
- Fournir une sortie organisée et consultable
- Répondre aux besoins spécifiques des workflows de production vidéo

MediaMetadataParser est un outil puissant pour extraire et organiser les métadonnées des fichiers multimédias. Il prend en charge divers formats multimédias et fournit des informations détaillées aux formats Excel et JSON.

## Fonctionnalités

- Extrait des métadonnées complètes incluant :
  - Durée (en secondes et au format HH:MM:SS)
  - Résolution (largeur x hauteur)
  - FPS (images par seconde)
  - Informations sur le codec
  - Format de pixel
  - Profondeur de bits
  - Rotation
  - Débit binaire
  - Espace colorimétrique
  - Taille du fichier (en octets et Mo)
  - Date de modification
  - Métadonnées techniques supplémentaires des en-têtes de fichier
- Prend en charge plusieurs formats multimédias :
  - Vidéo : .mp4, .avi, .mkv, .mov
  - Audio : .mp3, .wav, .flac, .m4a, .aac
  - Formats supportés : .mp3, .mp4, .avi, .mkv, .mov, .wav, .flac, .m4a, .aac
- Analyse récursive des répertoires
- Exclut les fichiers cachés (ceux commençant par '.')
- Fournit :
  - Nombre total de fichiers multimédias
  - Taille totale en Go
  - Métadonnées détaillées pour chaque fichier
  - Résultats sauvegardés dans un fichier Excel avec ajustement automatique des largeurs de colonnes
  - Sortie JSON optionnelle avec conversion de type appropriée
  - Mémorise le dernier répertoire utilisé via un fichier temporaire
  - Suivi de la progression avec pourcentage d'avancement
  - Support de l'annulation
  - Gestion des erreurs pour les fichiers problématiques
  - Interface graphique avec :
    - Documentation extensible
    - Lien GitHub
    - Améliorations de style
    - Validation des entrées

## Installation

1. Clonez le dépôt :
```bash
git clone https://github.com/thiswillbeyourgithub/MediaMetadataParser.git
cd MediaMetadataParser
```

2. Installez les dépendances :
```bash
pip install -r requirements.txt
```

## Utilisation

Exécutez le script pour lancer l'application graphique :
```bash
python MediaMetadataParser.py
```

1. Sélectionnez un dossier contenant des fichiers multimédias
2. Choisissez un emplacement de sortie
3. Cliquez sur 'Démarrer le traitement' pour commencer l'extraction des métadonnées

L'application va :
- Analyser le répertoire sélectionné
- Extraire les métadonnées de tous les fichiers multimédias supportés
- Sauvegarder les résultats dans un fichier Excel
- Optionnellement sauvegarder les résultats au format JSON

## Prérequis

- Python 3.8+
- Paquets requis :
  - moviepy (pour l'extraction des métadonnées multimédias)
  - openpyxl (pour la création de fichiers Excel)
  - tkinter (pour l'interface graphique)
  - json (pour la sortie JSON optionnelle)

## Contribution

Les contributions sont les bienvenues ! Veuillez suivre ces étapes :

1. Forkez le dépôt
2. Créez une nouvelle branche (`git checkout -b feature/NomDeVotreFonctionnalité`)
3. Committez vos changements (`git commit -m 'Ajouter une fonctionnalité'`)
4. Poussez vers la branche (`git push origin feature/NomDeVotreFonctionnalité`)
5. Créez une nouvelle Pull Request

## Licence

Ce projet est sous licence GPLv3 - voir le fichier [LICENCE](LICENSE) pour plus de détails.

## Support

Si vous trouvez ce projet utile, pensez à mettre une étoile au dépôt ⭐

Pour les problèmes ou demandes de fonctionnalités, veuillez ouvrir un issue sur GitHub.

## Exemple de sortie

L'application génère une feuille de calcul Excel détaillée avec les métadonnées de chaque fichier multimédia, incluant :
- Nom et chemin du fichier
- Taille du fichier en octets et Mo
- Date de modification
- Durée en secondes et au format HH:MM:SS
- Résolution (pour les fichiers vidéo)
- FPS (pour les fichiers vidéo)
- Informations sur le codec
- Détails techniques supplémentaires
