C'est parfait. Pour que le code de ton **TP4** (et des autres TP) fonctionne, l'utilisateur doit impérativement installer **Google OR-Tools** (le solveur) et **OpenPyXL** (pour lire les fichiers Excel).

Voici le `README.md` complet, "excellent" et prêt à l'emploi, incluant la section **Installation des bibliothèques** :

---

```markdown
# TP-RO : Travaux Pratiques de Recherche Opérationnelle

Ce dépôt contient l'ensemble des exercices, codes sources et ressources pour le module de **Recherche Opérationnelle (RO)**.

## 📁 Structure du Projet

Le projet est organisé par répertoires correspondant à chaque séance de TP :

* **TP1** : Initiation et exercices de base.
* **TP2** : Suite des implémentations algorithmiques.
* **TP3** : Code principal du TP3 (Note : une version de développement existe sur la branche `test`).
* **TP4** : Optimisation (Bin Packing) avec Google OR-Tools.
* **exemple** : Dossier contenant des modèles d'exemples.

### 📊 Fichiers de données (Racine)
| Fichier | Description |
| :--- | :--- |
| `exercice1.xlsx` | Données pour les calculs de l'exercice 1 (TP4) |
| `exercice2.xlsx` | Données pour les calculs de l'exercice 2 |
| `exercice3.xlsx` | Données pour les calculs de l'exercice 3 |

---

## 💻 Langages et Technologies
* **Langage principal** : Python 3.x
* **Solveur** : Google OR-Tools (Linear Programming / MIP)
* **Gestion de versions** : Git / GitHub

---

## 🛠️ Installation et Prérequis

Avant de lancer les scripts, vous devez installer les dépendances nécessaires.

### 1. Installation des bibliothèques Python
Ouvrez votre terminal et exécutez la commande suivante :
```bash
pip install ortools openpyxl pandas

```

### 2. Cloner le dépôt

```bash
git clone https://github.com/mohameden19961/TP-RO.git

```

---

## ⚙️ Utilisation (Exemple pour le TP4)

Pour exécuter l'exercice d'optimisation du TP4 :

1. **Entrer dans le dossier :**
```bash
cd TP-RO/TP4

```


2. **Lancer le script :**
```bash
python exercice1.py

```



> **Note :** Assurez-vous que les fichiers `.xlsx` sont bien présents à la racine du projet ou dans le dossier indiqué par le script pour éviter les erreurs `FileNotFoundError`.

---

## 📋 Fonctionnalités du code (TP4)

Le script utilise le solveur **CBC (Mixed Integer Programming)** pour :

* Lire les données de poids et de capacité depuis un fichier Excel.
* Minimiser le nombre de conteneurs (vols) nécessaires.
* Sauvegarder automatiquement les résultats dans l'onglet "Résultats" du fichier Excel.

---


```
