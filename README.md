# 📦 Réconciliation de stock LEONI

Application web (Streamlit) qui vérifie, article par article, que

```
Stock final = Stock initial + Mouvements        →        Écart = SF − SI − Mouvement
```

et exporte le résultat dans un fichier Excel.

## Ce que fait l'application

1. **Plant** : liste déroulante (29, 47, 48, 66, 77). Chaque plant a sa propre table des signes
   (+1 entrée, −1 sortie, 0 neutre).
   Au choix du plant, un message d'accueil animé s'affiche en haut de la page puis s'évapore
   (29 : Iheb · 47 : Youssef · 48 : Adnen · 66 : Jawher ; dictionnaire `SALUTATIONS` dans `app.py`).
   Le plant 77 est partagé : l'utilisateur choisit d'abord **Sami** ou **Nejma** (cartes cliquables) et le message devient
   « Bonjour Sami ! » ou « Bonjour Nejma ! » (dictionnaire `PERSONNES_PAR_PLANT`).
2. **Stocks** : dépôt du stock initial et du stock final (rapports FMALS015).
3. **Mouvements** : dépôt d'**un ou deux** fichiers (rapports FMALS032).
   Un choix en haut de l'étape permet d'**appliquer le redressement CW** (détecté et appliqué fichier par fichier ;
   un fichier sans CW n'est pas modifié) ou de **laisser les fichiers tels quels** (le magasin « CW » reste alors dans la liste).
   Le choix est tracé dans l'onglet Contrôles et dans la feuille `Paramètres`.
4. **Magasins à supprimer** : l'application liste les magasins présents dans les fichiers de
   mouvements (après redressement). On coche ceux à supprimer ; si on ne coche rien, rien n'est supprimé.
   Les zéros à gauche sont ignorés (`010953` = `10953`).
   **Aperçu avant suppression** : pour les magasins cochés, l'application affiche les valeurs de la colonne
   « by warehouses » associées (lignes + quantité). Chaque ligne a une case **Effacer** : cochée = supprimée,
   décochée = conservée. Les exceptions sont tracées dans la feuille `Paramètres`.
5. **Résultats** : indicateurs, tableau de réconciliation (recherche, filtre « seulement les écarts »),
   plus gros écarts, **magasins à écarts**, TCD par fichier, contrôles, et **téléchargement Excel**.

### Onglet « Magasins à écarts »

Classe les magasins selon les écarts des articles qui y ont bougé :

- **Écart € concerné** : écarts (valeur absolue) des articles en écart présents dans le magasin ;
- **Écart € imputé** : cet écart réparti entre les magasins au prorata de leur volume de mouvement (la somme = écart total) ;
- **Articles annulés** : articles dont l'écart disparaît si on ignore le mouvement net du magasin (tolérance 2 %) —
  indice le plus fort que le magasin est la cause ; à vérifier, puis à cocher dans « Magasins à supprimer » ;
- **Niveau** : 🔴 ≥ 20 % de l'écart total · 🟠 ≥ 8 % · 🟢 sinon.

Un sélecteur affiche le détail article par article d'un magasin. Chaque ligne = un couple **magasin + « by warehouses »**. Les feuilles `Magasins à écarts` et `Détail magasins`
(50 000 lignes max) sont ajoutées à l'Excel.

### Contrôles affichés (onglet « Contrôles »)

- lignes lues / CW détectées / supprimées / retenues, par fichier de mouvements ;
- **mouvements absents de la table des signes** (signe 0 appliqué : ils sont ignorés dans le calcul, mais signalés) ;
- **PartNumbers qui ont des mouvements mais n'existent dans aucun des deux stocks** (non repris dans la réconciliation) ;
- magasins supprimés.

### Fichier Excel

Feuilles : `Réconcil`, `TCD1` (`TCD2`), `Paramètres` (plant, fichiers, magasins supprimés, date),
`Magasins à écarts`, `Détail magasins`, `Magasins supprimés` (magasins cochés + « by warehouses » liés, lignes, quantité,
statut Supprimé / Conservé, par fichier ; présente seulement si un magasin est coché) et selon le cas `Mouvements inconnus` et `Hors stock`. Les écarts non nuls sont en rouge.

## Format des fichiers attendus

| Fichier | Rapport | Lignes d'en-tête ignorées | Colonnes utilisées |
|---|---|---|---|
| Mouvements | FMALS032 (CSV `;`) | 12 | `PartNr.`, `BookType`, `Quantity`, `Type`, `Warehouse` (col. I), `by warehouses` (col. M) |
| Stock initial / final | FMALS015 (CSV `;`) | 9 | col. A = PartNumber, col. D = Type, col. G = Stock, col. J = Value |

Le **PU** est calculé par `Value / Stock` (valeur du stock final en priorité, sinon du stock initial).

## Déploiement gratuit (Streamlit Community Cloud)

1. Créez un dépôt GitHub et déposez-y, **à la racine**, les trois fichiers : `app.py`, `requirements.txt`, `README.md`.
2. Allez sur <https://share.streamlit.io> et connectez-vous avec GitHub.
3. **Create app** → choisissez le dépôt, la branche (`main`) et le fichier principal **`app.py`**.
4. Cliquez sur **Deploy** : l'application obtient une adresse publique `https://<nom>.streamlit.app`.

Une application inactive peut être mise en veille ; un clic sur « réveiller » la relance.

## Lancer en local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Modifier les règles de signes

Dans l'application (étape 4 « Signes des mouvements ») : chaque signe peut être changé à la volée (1, 0 ou −1) ;
les types de mouvement hors table présents dans les fichiers sont ajoutés avec le signe 0. Un bouton rétablit la table
par défaut, et les signes modifiés sont tracés dans la feuille `Paramètres` de l'Excel. Ces changements ne durent que pour la session.

Pour modifier la table par défaut, ou ajouter un plant (ex. 29) :

1. Ouvrir `app.py` et éditer le dictionnaire `SIGNES_PAR_PLANT` (une table par plant).
2. Pour un nouveau plant : ajouter sa table dans `SIGNES_PAR_PLANT` **et** son numéro dans la liste `PLANTS`.

Clé = `"<BookType>_<Type>"` (ex. `"12_DESPATCHED"`), valeur = `1`, `-1` ou `0`.
Un type absent de la table reçoit le signe `SIGNE_PAR_DEFAUT` (0).

## Confidentialité

- Les fichiers déposés sont traités **en mémoire** : l'application n'écrit rien sur disque.
- Un cache mémoire (1 heure maximum) évite de relire les fichiers à chaque clic.
- L'adresse d'une application Streamlit Community Cloud est publique : toute personne qui l'ouvre
  peut déposer ses propres fichiers. Si vos données sont confidentielles, restreignez l'accès à
  l'application dans les paramètres de partage de Streamlit Community Cloud.

## Dépannage

| Message | Cause probable |
|---|---|
| `colonnes manquantes [...]` | ce n'est pas un rapport FMALS032, ou les 12 lignes d'en-tête ont changé |
| `colonne(s) requise(s) en double` | deux colonnes ont le même nom après suppression des espaces (ex. `Type` et `Type `) |
| `au moins 10 colonnes sont attendues` | ce n'est pas un rapport FMALS015 |
| Avertissement « le nom du stock indique le plant XX » | le plant choisi ne correspond pas au nom du fichier de stock |
| Beaucoup de mouvements inconnus | un plant a de nouveaux types de mouvement : ajoutez-les dans `SIGNES_PAR_PLANT` |
