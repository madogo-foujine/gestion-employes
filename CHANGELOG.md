# Changelog

Tous les changements notables de ce projet sont documentés dans ce fichier.

## [Non publié]

### Sécurité
- Mots de passe : remplacement de SHA-256 par **scrypt** (mémoire-dur, `hashlib`),
  avec sel aléatoire et compatibilité ascendante (les anciens hachages SHA-256
  restent vérifiables). Comparaison à temps constant (`secrets.compare_digest`).
- Validation des valeurs du fichier de configuration (plages contrôlées,
  retour aux valeurs par défaut si invalides).
- Validation des saisies utilisateur (champs numériques ≥ 0, dates valides).
- Journalisation des erreurs (`logging`) au lieu d'`except … : pass` silencieux.

### Ajouté
- Tableau de bord enrichi : 5 indicateurs (Effectif, Masse brute, Masse nette,
  Net moyen, Total IR).
- Suite de tests unitaires (`tests/`) avec `pytest` : moteur de paie + sécurité.
- Intégration continue (GitHub Actions) : tests + build de l'exécutable Windows.
- Documentation : `CHANGELOG.md` et `docs/FEATURES.md`.

## [1.0] - 2026-06-29

### Ajouté
- Gestion des employés reliée à un fichier Excel (`employes.xlsx`).
- Moteur de paie marocain : CNSS, AMO, frais professionnels, IR (barème
  progressif paramétrable), net à payer.
- Pointage mensuel, calendrier annuel, congés, avances sur salaire.
- Bulletins de paie (PDF/HTML), attestations, contrat de travail, déclaration
  CNSS, état de paie (HTML/Excel) — avec logo et signature.
- Numérotation séquentielle des documents et registre.
- Tableau de bord, graphiques, évolution de la masse salariale.
- Import/export CSV, sauvegardes automatiques, mode sombre, mot de passe et
  rôles, photos et documents par employé.
- Distribution : exécutable `.exe` (PyInstaller) + installateur (Inno Setup).
