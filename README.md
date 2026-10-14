# StockFlow — LEONI Stock Reconciliation

Application Streamlit pour réconcilier les stocks FMALS015 et les mouvements FMALS032.

## Fonctionnalités
- Plants 47, 48, 66 et 77 avec tables de signes spécifiques.
- Import du stock initial, du stock final et d'un ou deux fichiers de mouvements.
- Détection dynamique des magasins dans les fichiers de mouvements et sélection par cases à cocher.
- Redressement CW automatique optionnel.
- Réconciliation `Écart = SF - SI - Mouvement`.
- Export Excel : `Réconcil`, `TCD1`, `TCD2` (si deux fichiers) et `Contrôles`.

## Lancer localement
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Déploiement gratuit
1. Crée un dépôt GitHub.
2. Ajoute `app.py` et `requirements.txt`.
3. Ouvre https://share.streamlit.io/ et connecte ton compte GitHub.
4. Clique **Create app**, choisis le dépôt, la branche et `app.py`, puis déploie.
5. Partage l'URL publique générée.

Documentation officielle : https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy

## Confidentialité
Ne publie jamais de fichiers réels de stock ou de mouvements dans GitHub. Une application publique est accessible à toute personne disposant de l'URL. N'utilise pas de données d'entreprise confidentielles sur un service public sans autorisation explicite. Le code et les règles métier peuvent être visibles si le dépôt est public.
