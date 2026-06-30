# Changelog

Tous les changements notables de ce projet sont documentés dans ce fichier.

## [Non publié]

### Interface
- Interface modernisée avec **ttkbootstrap** (thèmes clair *flatly* / sombre
  *darkly*, boutons colorés avec effet *hover*).
- **Splash screen** au démarrage (logo, version, barre de chargement).
- **Raccourcis clavier** : Ctrl+N (nouveau), Ctrl+S (enregistrer),
  Ctrl+F (recherche), Ctrl+P (bulletin).
- **Tooltips** sur la barre d'outils, **notifications Toast** non intrusives.

### Ajouté
- **Journal d'audit** : chaque opération (ajout, modification, suppression,
  archivage, congés) est tracée (date, utilisateur) — lecture seule.
- **Archivage des employés** (suppression douce) avec restauration.
- **Système de congés** : demandes, approbation/refus, types
  (annuel/maladie/sans solde), solde lié au calcul des congés.
- **Heures (entrée/sortie)** : saisie par jour, calcul automatique des heures
  travaillées, du retard et des **heures supplémentaires** (×1,25), report
  possible sur les primes.
- Tableau de bord enrichi : 5 indicateurs (effectif, masse brute/nette,
  net moyen, total IR).
- Recherche élargie (ID, e-mail), tests `pytest` (paie, sécurité, temps),
  intégration continue GitHub Actions, `CHANGELOG.md`, `docs/FEATURES.md`.

### Sécurité
- Mots de passe : **scrypt** (au lieu de SHA-256), sel aléatoire, comparaison
  à temps constant, compatibilité ascendante.
- Validation de la configuration et des saisies utilisateur.
- Journalisation des erreurs (`logging`).

## [1.0] - 2026-06-29

### Ajouté
- Gestion des employés reliée à un fichier Excel (`employes.xlsx`).
- Moteur de paie marocain : CNSS, AMO, frais professionnels, IR (barème
  progressif paramétrable), net à payer.
- Pointage mensuel, calendrier annuel, congés (solde), avances sur salaire.
- Bulletins de paie (PDF/HTML), attestations, contrat de travail, déclaration
  CNSS, état de paie (HTML/Excel) — avec logo et signature.
- Numérotation séquentielle des documents et registre.
- Tableau de bord, graphiques, évolution de la masse salariale.
- Import/export CSV, sauvegardes automatiques, mode sombre, mot de passe et
  rôles, photos et documents par employé.
- Distribution : exécutable `.exe` (PyInstaller) + installateur (Inno Setup).
