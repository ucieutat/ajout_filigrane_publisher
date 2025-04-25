# Filigrane Publisher

**Filigrane Publisher** est une application Python qui permet d'ajouter un filigrane personnalisé à des fichiers Microsoft Publisher (`.pub`) et de les exporter au format PDF. L'interface utilisateur est construite avec [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) pour une expérience moderne et intuitive.

---

## 🚀 Fonctionnalités

- Ajout d'un filigrane personnalisé à plusieurs fichiers Publisher.
- Exportation des fichiers Publisher en PDF.
- Interface utilisateur moderne avec basculement entre les modes clair et sombre.
- Création automatique d'un dossier pour les fichiers exportés.
- Journalisation des actions dans un fichier `programme.log`.

---

## 🛠️ Prérequis

Avant de commencer, assurez-vous d'avoir les éléments suivants installés sur votre machine :

1. **Python 3.8 ou supérieur** : [Télécharger Python](https://www.python.org/downloads/)
2. **Microsoft Publisher** : Nécessaire pour manipuler les fichiers `.pub`.
3. **Dépendances Python** :
   - `customtkinter`
   - `pywin32`

---

## 📦 Installation

1. Clonez ce dépôt ou téléchargez les fichiers :
   ```bash
   git clone https://github.com/ucieutat/ajout_filigrane_publisher.git
   cd ajout_filigrane_publisher