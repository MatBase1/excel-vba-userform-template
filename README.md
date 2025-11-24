# Modèle de UserForm Excel VBA

Ce dépôt propose un exemple simple et propre de **UserForm Excel VBA** pour illustrer une structuration correcte d’interface utilisateur et de code.

Le formulaire permet de saisir des informations de type “client” (nom, e-mail, pays) et d’enregistrer les données dans une feuille Excel dédiée.

---

## Fonctionnalités

- Formulaire utilisateur (UserForm) avec :
  - Champ **Nom**
  - Champ **E-mail**
  - Liste déroulante **Pays**
  - Boutons **Enregistrer** et **Annuler**
- Validation basique des champs (nom et e-mail obligatoires)
- Enregistrement automatique des données dans la feuille `Clients`
- Ajout à la **première ligne vide disponible**
- Code VBA commenté pour servir de modèle

---

## Prérequis

- Microsoft Excel (version Desktop)
- Macros activées (`.xlsm`)
- Connaissances de base en VBA (optionnel mais recommandé)

---

## Installation

1. Télécharger le fichier `clients.xlsm`.
2. Ouvrir le fichier dans Excel.
3. Activer les macros si nécessaire.
4. Appuyer sur `Alt + F11` pour ouvrir l’éditeur VBA et consulter le code.

---

## Utilisation

1. Depuis Excel, exécuter la macro :

   - `Outils > Macro > Macros…`  
   - Sélectionner `AfficherFormulaireClient`  
   - Cliquer sur **Exécuter**

2. Remplir les champs du formulaire :

   - Nom
   - E-mail
   - Choisir un pays dans la liste

3. Cliquer sur **Enregistrer** :

   - Les données sont ajoutées à la première ligne vide de la feuille `Clients`.

4. Cliquer sur **Annuler** pour fermer le formulaire sans enregistrer.

---

## Structure du code

- **UserForm** : `frmClient`
  - `txtNom` (TextBox)
  - `txtEmail` (TextBox)
  - `cboPays` (ComboBox)
  - `cmdEnregistrer` (CommandButton)
  - `cmdAnnuler` (CommandButton)

- **Module standard** : `modFormulaireClient`
  - `Sub AfficherFormulaireClient()`
  - Procédures utilitaires pour trouver la première ligne vide, etc.

---

## Avertissement

Ce projet est un **exemple pédagogique**.  
Il ne contient aucun code client réel ni logique métier sensible.  
Tu peux l’utiliser comme base pour tes propres développements Excel/VBA.

---

## Auteur

**Matthieu Chenal – DOPHIS**  
Expert Excel & VBA – Développement d’outils métier, automatisation & optimisation de fichiers.
