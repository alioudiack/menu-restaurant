# Menu restaurant avec Streamlit

Cette application affiche le menu digital du Restaurant La Galette avec photos, descriptions, compositions, prix, filtres par categorie et generation de QR code.

## Lancer en local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Modifier le menu

Les plats sont stockes dans `Menu la Galette.xlsx`, feuille `Restaurant`.
Vous pouvez modifier le menu directement dans Excel ou depuis l'onglet `Administration` de l'application.

Les colonnes utilisees sont :

- `Plat` : nom du plat
- `Description` : courte description
- `Catégorie` : categorie
- `Composition` : ingredients ou composition
- `image` : chemin de l'image, par exemple `image\Touffé.jfif`
- `Prix (FCFA)` : prix affiche au client
- `Disponible` : `Oui` pour afficher le plat, `Non` pour le cacher du menu client
- `Service` : `Déjeuner`, `Dîner` ou `Déjeuner et dîner`

## Menu temporel

Par defaut, le menu client filtre automatiquement les plats selon l'heure :

- de 08h00 a 17h00 : plats du `Déjeuner`
- apres 17h00 et avant 08h00 : plats du `Dîner`
- les plats `Déjeuner et dîner` restent visibles tout le temps

Dans la barre laterale, vous pouvez passer en mode manuel pour tester un autre service.

## Administration

Dans l'application, ouvrez `Administration`.

Mot de passe local par defaut :

```text
admin123
```

Avant un deploiement public, changez ce mot de passe avec un secret Streamlit :

```toml
ADMIN_PASSWORD = "votre-mot-de-passe"
```

L'administrateur peut :

- indiquer si un plat est disponible ou indisponible
- indiquer si un plat est servi au dejeuner, au diner ou aux deux
- modifier le prix affiche
- modifier le nom, la categorie, la description, la composition et l'image
- ajouter un nouveau plat avec une image

## QR code

1. Deployeez l'application, par exemple sur Streamlit Community Cloud.
2. Copiez l'URL publique de l'application.
3. Ouvrez l'onglet `QR code`.
4. Collez l'URL et telechargez le QR code.
5. Imprimez ce QR code pour les tables du restaurant.
