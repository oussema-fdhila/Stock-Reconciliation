# 📦 Réconciliation de stock LEONI

Application web (Streamlit) qui vérifie, article par article, que

```
Stock final = Stock initial + Mouvements        →        Écart = SF − SI − Mouvement
```

et exporte le résultat dans un fichier Excel.

## Ce que fait l'application

1. **Plant** : liste déroulante (29, 47, 48, 66, 77). Chaque plant a sa propre table des signes
   (+1 entrée, −1 sortie, 0 neutre).
2. **Stocks** : dépôt du stock initial et du stock final (rapports FMALS015).
3. **Mouvements** : dépôt d'**un ou deux** fichiers (rapports FMALS032).
   Le **redressement CW est détecté et appliqué automatiquement**, fichier par fichier :
   un fichier sans CW (ou déjà redressé) n'est pas modifié.
4. **Magasins à supprimer** : l'application liste les magasins présents dans les fichiers de
   mouvements (après redressement). On coche ceux à supprimer ; si on ne coche rien, rien n'est supprimé.
   Les zéros à gauche sont ignorés (`010953` = `10953`).
5. **Résultats** : indicateurs, tableau de réconciliation (recherche, filtre « seulement les écarts »),
   plus gros écarts, TCD par fichier, contrôles, et **téléchargement Excel**.

### Contrôles affichés (onglet « Contrôles »)

- lignes lues / CW détectées / supprimées / retenues, par fichier de mouvements ;
- **mouvements absents de la table des signes** (signe 0 appliqué : ils sont ignorés dans le calcul, mais signalés) ;
- **PartNumbers qui ont des mouvements mais n'existent dans aucun des deux stocks** (non repris dans la réconciliation) ;
- magasins supprimés.

### Fichier Excel

Feuilles : `Réconcil`, `TCD1` (`TCD2`), `Paramètres` (plant, fichiers, magasins supprimés, date),
et selon le cas `Mouvements inconnus` et `Hors stock`. Les écarts non nuls sont en rouge.

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
