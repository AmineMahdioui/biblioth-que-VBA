# ReadMe

## Aperçu du Projet

Ce projet implique le développement et la maintenance d'une application VBA. L'application comprend divers formulaires et modules pour gérer différentes fonctionnalités.

## À Faire

- [ ] Corriger `Jour_Disponible` dans le formulaire `livraison`
- [ ] Formulaire `commande` : Ajouter une userform
- [ ] 
## Base de Données


### **Table EDITEURS**

| Nom de colonne   | Type          | Description                                       |
|------------------|---------------|---------------------------------------------------|
| `ID_Editeur`     | NuméroAuto    | Clé primaire, identifiant des éditeurs           |
| `Nom_Editeur`    | Texte         | Nom de l'éditeur                                 |

---

### **Table COLLECTIONS**
| Nom de colonne   | Type          | Description                                       |
|------------------|---------------|---------------------------------------------------|
| `ID_Collection`  | NuméroAuto    | Clé primaire, identifiant des collections         |
| `Nom_Collection` | Texte         | Nom de la collection                             |

---

### **Table LIVRES**
| Nom de colonne      | Type          | Description                                       |
|---------------------|---------------|---------------------------------------------------|
| `ID_Livre`          | Texte         | Clé primaire, identifiant unique (ex : ISBN)      |
| `Titre_Livre`       | Texte         | Titre du livre                                   |
`Prix` | Numérique | Prix du livre
| `Possede`              | Booléen       | Indique si le livre est en possession |
| `ID_Collection`     | Numérique     | Clé étrangère de la table COLLECTIONS            |
| `ID_Editeur`        | Numérique     | Clé étrangère de la table EDITEURS               |

---

### **Table POINTS_DE_LIVRAISON**
| Nom de colonne       | Type          | Description                                       |
|----------------------|---------------|---------------------------------------------------|
| `ID_Point`           | NuméroAuto    | Clé primaire, identifiant des points de livraison |
| `Nom_Point`          | Texte         | Nom ou description du point de livraison         |
| `Adresse`            | Texte         | Adresse complète du point de livraison           |
| `Jour_Disponible`    | Texte         | Jours où le point est opérationnel               |

---

### **Table COMMANDES**
| Nom de colonne       | Type          | Description                                       |
|----------------------|---------------|---------------------------------------------------|
| `ID_Commande`        | NuméroAuto    | Clé primaire, identifiant des commandes          |
| `Date_Commande`      | Date          | Date de la commande                              |
| `Total_HT`           | Numérique     | Montant total hors taxes                         |
| `TVA`                | Numérique     | Montant de la TVA                                |
| `Total_TTC`          | Numérique     | Montant total toutes taxes comprises             |
| `ID_Point_Livraison` | Numérique     | Clé étrangère de la table POINTS_DE_LIVRAISON     |

---

### **Table DETAILS_COMMANDES**
| Nom de colonne       | Type          | Description                                       |
|----------------------|---------------|---------------------------------------------------|
| `ID_Commande`        | Numérique     | Clé étrangère de la table COMMANDES              |
| `ID_Livre`           | Texte         | Clé étrangère de la table LIVRES                 |
| `Sous_Total`         | Numérique     | Sous-total pour cet article                      |
