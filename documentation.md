# Structure du Plan pour la Documentation

## 1. Introduction

Ce projet a mis en œuvre une application sous Visual Basic pour gérer une bibliothèque de bandes dessinées. L'application permet aux utilisateurs d'ajouter et de supprimer des livres référencés, ainsi que suivre les livres déjà en leur possession. De plus, l'application permet aux utilisateurs de organier les albums commander en ligne pour les retirer chez un libraire ou en magasin.

## 2. Formulaire Principal

### Nom du Formulaire : `MainForm`

#### Description
Le formulaire principal, nommé `MainForm`, est la page centrale de l’application. Il offre une vue globale des fonctionnalités disponibles et permet à l'utilisateur d'accéder rapidement aux formulaires et actions nécessaires pour gérer la bibliothèque de bandes dessinées. L’objectif principal de ce formulaire est d’organiser et de simplifier l’accès aux différents modules du système.

#### Structure et Présentation
Le `MainForm` est conçu avec une interface intuitive qui regroupe 7 boutons principaux, chacun représentant une action ou une fonctionnalité clé. Chaque bouton ouvre un formulaire dédié ou déclenche une procédure spécifique.


### Boutons et Actions Associées

1. **Ajout Éditeur**
   - **Objectif** : Permet d'ajouter un éditeur dans la base de données.
   - **Action** : Ouvre le formulaire `AjoutEditeurForm`. Ce formulaire contient un champ pour saisir le nom d’un éditeur et un bouton pour valider l’ajout.
   - **Étapes après clic** :
     - L'utilisateur saisit un nom pour l'éditeur.
     - Après validation, l’éditeur est ajouté à la feuille `EDITEURS` avec un identifiant unique généré automatiquement.


2. **Suppression Éditeur**
   - **Objectif** : Supprimer un éditeur existant de la base de données.
   - **Action** : Ouvre le formulaire `SuppEditeurForm`. Une liste déroulante affiche les éditeurs existants, permettant à l'utilisateur de sélectionner celui à supprimer.
   - **Étapes après clic** :
     - L'utilisateur choisit un éditeur dans la liste déroulante.
     - Une vérification empêche la suppression si l'éditeur est lié à un livre référencé.
     - Si l'éditeur peut être supprimé, il est retiré de la feuille `EDITEURS`.

3. **Ajout Collection**
   - **Objectif** : Ajouter une nouvelle collection dans la base de données.
   - **Action** : Ouvre le formulaire `CollectionForm`. Ce formulaire permet de saisir le nom de la collection et de l’enregistrer.
   - **Étapes après clic** :
     - L'utilisateur entre le nom de la collection.
     - Après validation, la collection est ajoutée à la feuille `COLLECTIONS` avec un identifiant unique.


4. **Ajout Livre Référencé**
   - **Objectif** : Ajouter un nouveau livre référencé dans la base de données.
   - **Action** : Ouvre le formulaire `LivreRefForm`. Ce formulaire permet de saisir les détails du livre, notamment le titre, le prix, l'éditeur, et la collection via des champs de texte et des listes déroulantes.
   - **Étapes après clic** :
     - L'utilisateur renseigne les informations du livre.
     - Une fois validé, le livre est ajouté à la feuille `LIVRES` avec un identifiant unique, un prix, et des liens vers l’éditeur et la collection correspondants.


5. **Ajout Livre Possédé**
   - **Objectif** : Marquer un livre comme possédé par l'utilisateur.
   - **Action** : Ouvre le formulaire `AjoutLivrePForm`. Une liste déroulante affiche les livres référencés qui ne sont pas encore possédés (`Possédé = False`).
   - **Étapes après clic** :
     - L'utilisateur sélectionne un livre dans la liste déroulante.
     - Après validation, le statut `Possédé` du livre est mis à jour à `True` dans la feuille `LIVRES`.




6. **Ajout Point de Livraison**
   - **Objectif** : Ajouter un nouveau point de livraison avec son jour de disponibilité.
   - **Action** : Ouvre le formulaire `FormAjoutPointdeLivraison`. Ce formulaire permet de renseigner le nom, l'adresse, et le jour de disponibilité du point de livraison.
   - **Étapes après clic** :
     - L'utilisateur saisit les informations nécessaires.
     - Une validation garantit que le jour est au format correct (jour de la semaine, ex. "Lundi").
     - Le point de livraison est enregistré dans la feuille `POINTS DE LIVRAISON`.

    7. **Commande**
        - **Objectif** : Générer une nouvelle commande incluant plusieurs livres et créer une nouvelle feuille avec les détails de la commande.
        - **Action** : Ouvre le formulaire `CommandeUserForm`. Ce formulaire permet de sélectionner des livres et une date de livraison, puis de générer une commande.
        - **Étapes après clic** :
          - L'utilisateur sélectionne un ou plusieurs livres et une date de livraison.
          - Seuls les points de livraison disponibles pour la date sélectionnée sont affichés.
          - Après validation, une nouvelle feuille est créée pour la commande, avec les informations suivantes :
             - Les livres commandés, classés par éditeur et par titre.
             - Le total HT, la TVA (20%), et le total TTC.
             - Le point de livraison associé.


<!-- ### Interactions et Flux de Travail
1. **Navigation Simple** :
   - Chaque bouton du formulaire principal ouvre un formulaire distinct, facilitant l'accès direct à une fonctionnalité spécifique.

2. **Validation des Données** :
   - Tous les formulaires incluent des validations pour garantir que les données saisies sont complètes et conformes (ex. : format des dates ou saisie obligatoire).

3. **Suivi des Données** :
   - Les informations saisies ou modifiées via le formulaire principal sont directement mises à jour dans les feuilles de calcul correspondantes.

4. **Création Dynamique des Commandes** :
   - Le bouton "Commande" offre une fonctionnalité clé en générant une nouvelle feuille pour chaque commande, permettant un suivi organisé et une transparence des informations.


### Résumé
Le `MainForm` est le cœur de l'application, regroupant toutes les fonctionnalités majeures en un seul endroit. Chaque bouton est lié à une action spécifique qui s’appuie sur les feuilles de calcul Excel pour gérer les éditeurs, collections, livres, points de livraison, et commandes. Grâce à son interface claire et ses validations robustes, il garantit une gestion efficace et intuitive des données. -->



# 3. Formulaires et leurs Procédures

## 3.1 Ajout Éditeur

### **Description**
Le formulaire `AjoutEditeurForm` permet à l'utilisateur d'ajouter un nouvel éditeur dans la base de données. Il contient un champ de texte pour saisir le nom de l'éditeur et un bouton pour valider l'ajout.

### **Objets du Formulaire**
- **TextBoxNomEditeur** : Champ de texte pour saisir le nom de l'éditeur.
- **ButtonAddEditeur** : Bouton pour valider et enregistrer l'éditeur dans la base de données.

### **Procédure**
#### `ButtonAddEditeur_Click`
1. Valide que le champ `TextBoxNomEditeur` n'est pas vide.
2. Ajoute l'éditeur dans la feuille `EDITEURS` avec un identifiant unique auto-incrémenté.
3. Affiche un message de confirmation à l'utilisateur.
4. Réinitialise le champ de texte pour permettre une nouvelle saisie.

---

## 3.2 Suppression Éditeur

### **Description**
Le formulaire `SuppEditeurForm` permet de supprimer un éditeur existant. Une liste déroulante affiche tous les éditeurs disponibles, et l'utilisateur peut en sélectionner un à supprimer.

### **Objets du Formulaire**
- **ComboBoxSuppEditeur** : Liste déroulante affichant les éditeurs disponibles.
- **ButtonSupEditeur** : Bouton pour valider la suppression de l'éditeur.

### **Procédure**
#### `ButtonSupEditeur_Click`
1. Valide que l'utilisateur a sélectionné un éditeur dans `ComboBoxSuppEditeur`.
2. Vérifie que l'éditeur n'est pas associé à un livre dans la feuille `LIVRES`.
3. Supprime l'éditeur sélectionné de la feuille `EDITEURS`.
4. Affiche un message de confirmation.
5. Met à jour la liste déroulante pour refléter les modifications.

---

## 3.3 Ajout Livre Référencé

### **Description**
Le formulaire `LivreRefForm` permet d'ajouter un nouveau livre dans la base de données. L'utilisateur peut saisir le titre, le prix, et sélectionner l'éditeur et la collection via des listes déroulantes.

### **Objets du Formulaire**
- **TitreBox** : Champ de texte pour saisir le titre du livre.
- **PrixBox** : Champ de texte pour saisir le prix du livre.
- **EditeurComboBox** : Liste déroulante pour sélectionner l'éditeur.
- **CollectionComboBox** : Liste déroulante pour sélectionner la collection.
- **AjouterLivreRefButton** : Bouton pour valider et ajouter le livre.

### **Procédure**
#### `AjouterLivreRefButton_Click`
1. Valide que tous les champs (titre, prix, éditeur, collection) sont correctement remplis.
2. Ajoute le livre dans la feuille `LIVRES` avec un identifiant unique et ses propriétés (titre, prix, éditeur, collection).
3. Affiche un message de confirmation.
4. Réinitialise les champs pour permettre une nouvelle saisie.

---

## 3.4 Ajout Livre Possédé

### **Description**
Le formulaire `AjoutLivrePForm` permet de marquer un livre référencé comme possédé. L'utilisateur sélectionne un livre dans une liste déroulante des livres non possédés.

### **Objets du Formulaire**
- **ComboBoxLivrePossedee** : Liste déroulante affichant les livres non possédés.
- **PossedeeButton** : Bouton pour valider et marquer le livre comme possédé.

### **Procédure**
#### `PossedeeButton_Click`
1. Valide que l'utilisateur a sélectionné un livre dans `ComboBoxLivrePossedee`.
2. Met à jour la colonne `Possédé` à `True` dans la feuille `LIVRES` pour le livre sélectionné.
3. Affiche un message de confirmation.
4. Met à jour la liste déroulante pour refléter les modifications.

---

## 3.5 Ajout Collection

### **Description**
Le formulaire `CollectionForm` permet à l'utilisateur d'ajouter une nouvelle collection dans la base de données.

### **Objets du Formulaire**
- **TextBoxNomCollection** : Champ de texte pour saisir le nom de la collection.
- **ButtonAddCollection** : Bouton pour valider et enregistrer la collection.

### **Procédure**
#### `ButtonAddCollection_Click`
1. Valide que le champ `TextBoxNomCollection` n'est pas vide.
2. Ajoute la collection dans la feuille `COLLECTIONS` avec un identifiant unique.
3. Affiche un message de confirmation.
4. Réinitialise le champ de texte pour permettre une nouvelle saisie.

---

## 3.6 Ajout Point de Livraison

### **Description**
Le formulaire `FormAjoutPointdeLivraison` permet d'ajouter un point de livraison dans la base de données. L'utilisateur renseigne le nom, l'adresse, et le jour disponible.

### **Objets du Formulaire**
- **TextBoxNomPL** : Champ de texte pour saisir le nom du point de livraison.
- **TextBoxAddresse** : Champ de texte pour saisir l'adresse du point de livraison.
- **TextBoxJourDispo** : Champ de texte pour saisir le jour disponible (ex. : "Lundi").
- **ButtonAddPoint** : Bouton pour valider et enregistrer le point de livraison.

### **Procédure**
#### `ButtonAddPoint_Click`
1. Valide que tous les champs (nom, adresse, jour disponible) sont remplis.
2. Vérifie que le jour disponible correspond à un jour valide (ex. : "Lundi", "Mardi").
3. Ajoute le point de livraison dans la feuille `POINTS DE LIVRAISON`.
4. Affiche un message de confirmation.
5. Réinitialise les champs pour permettre une nouvelle saisie.

---

## 3.7 Commande

### **Description**
Le formulaire `CommandeUserForm` permet de générer une nouvelle commande. L'utilisateur sélectionne des livres et un point de livraison pour créer une commande.

### **Objets du Formulaire**
- **ComboBoxLivre** : Liste déroulante permettant de sélectionner plusieurs livres.
- **ComboBoxPointLivraison** : Liste déroulante affichant les points de livraison disponibles pour le jour actuel.
- **ButtonGenerateOrder** : Bouton pour valider et créer la commande.
- **LabelTVA** : Étiquette affichant le montant de la TVA calculée.

### **Procédure**
#### `ButtonGenerateOrder_Click`
1. Valide que l'utilisateur a sélectionné au moins un livre et un point de livraison.
2. Vérifie que le point de livraison est disponible pour le jour actuel.
3. Ajoute les détails de la commande dans une nouvelle feuille nommée avec l'`ID_Commande`.
4. Classe les livres par éditeur et par titre dans la commande.
5. Calcule le total HT, la TVA (20%), et le total TTC.
6. Affiche le montant de la TVA dans `LabelTVA`.
7. Affiche un message de confirmation.

# 4. Base de Données

La base de données est le fondement de notre application. Elle est organisée en différentes tables qui permettent de stocker et de gérer toutes les informations nécessaires à la gestion de la bibliothèque, des commandes, et des points de livraison. Chaque table a un rôle précis, et elles sont reliées entre elles pour garantir la cohérence des données.

---

## **Les Éditeurs**
La table des éditeurs est dédiée à l’enregistrement des maisons d’édition. Chaque éditeur est identifié par un numéro unique généré automatiquement. Cette table contient également le nom de chaque éditeur. Elle est essentielle pour attribuer un éditeur à chaque livre dans la bibliothèque. Par exemple, si un livre est publié par une maison spécifique, cette information est directement liée à la table des éditeurs.

---

## **Les Collections**
Les collections regroupent des séries ou catégories de livres. Chaque collection a un identifiant unique et un nom descriptif. Par exemple, si plusieurs livres font partie d’une même série ou thématique, cette table permet de les regrouper sous une même entité. Les collections sont utilisées pour organiser les livres et faciliter leur consultation.

---

## **Les Livres**
La table des livres est au cœur de l’application. Elle contient les informations détaillées sur chaque livre référencé :
- Son titre et son prix, qui sont des données importantes pour les commandes.
- Une indication pour savoir si le livre est déjà en possession de l'utilisateur ou non. 
- Des liens directs avec les éditeurs et les collections, ce qui permet de structurer les données et de regrouper les livres par thématique ou éditeur.

Par exemple, si un utilisateur souhaite ajouter un livre à sa bibliothèque, il peut rapidement vérifier s'il est déjà possédé ou non grâce à cette table.

---

## **Les Points de Livraison**
Cette table enregistre les lieux où les commandes peuvent être récupérées. Chaque point de livraison est identifié par un numéro unique et contient des informations comme :
- Son nom, pour identifier facilement le point.
- Son adresse complète, pour savoir où se rendre.
- Les jours de disponibilité (par exemple, "Lundi", "Mardi"), ce qui permet d’assurer que la commande est compatible avec le jour actuel.

Par exemple, si une commande est passée un lundi, seuls les points disponibles ce jour-là seront proposés.


## **Les Commandes**
La table des commandes conserve l’historique des commandes générées par l’utilisateur. Chaque commande possède :
- Un identifiant unique.
- La date à laquelle elle a été créée.
- Les montants hors taxes, la TVA (calculée automatiquement à 20 %), et le montant TTC.
- Le point de livraison choisi, lié directement à la table des points de livraison.

Cette table joue un rôle crucial pour le suivi des commandes, en conservant toutes les informations nécessaires pour les consulter ou les modifier ultérieurement.

---


## **Relation entre les Tables**
Les tables de la base de données ne fonctionnent pas de manière isolée ; elles sont reliées entre elles pour garantir une cohérence des données :
- Les livres sont liés aux éditeurs et aux collections, ce qui permet une organisation structurée.
- Les commandes sont associées aux livre et points de livraison, ce qui garantit que chaque commande a un lieu de récupération précis.

Ces relations rendent la base de données flexible et permettent à l’utilisateur de naviguer facilement entre les informations, par exemple en consultant tous les livres d’un éditeur ou toutes les commandes récupérées à un point de livraison spécifique.

---
