# planning-tallard

Planning de la voiture de service pour Tallard : https://valleejc.github.io/planning-tallard/

- `generer.py` : lit les EDT ADE, garde les cours en présentiel des enseignants habilités et génère `docs/index.html` + `docs/data.json` (config en haut du fichier).
- `template.html` : toute l'interface (HTML/CSS/JS sans dépendance) ; `__DATA__` est remplacé par les données.
- `.github/workflows/planning.yml` : lance `generer.py` chaque matin et publie `index.html` + `data.json` sur la branche `gh-pages`.
  `data.json` est relu à la synchro suivante pour détecter les cours ajoutés, déplacés ou annulés.

En local : `pip install -r requirements.txt` puis `python generer.py`, et ouvrir `docs/index.html`.

`update_planning.py` et `update_planning_v2.py` sont les anciennes versions (plus utilisées).
