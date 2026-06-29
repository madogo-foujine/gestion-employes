# Gestion des Employés & Paie

Application de bureau (Windows) pour la gestion du personnel et de la paie,
adaptée au contexte marocain (CNSS / AMO / IR), reliée à un fichier Excel.

## Fonctionnalités

- **Fiches employés** : informations personnelles, contact, administratif, salaire.
- **Moteur de paie marocain** : calcul automatique CNSS (4,48 %, plafond 6000),
  AMO (2,26 %), frais professionnels, IR (barème progressif), net à payer.
  Taux et barème IR **paramétrables** depuis l'application.
- **Pointage** mensuel + **calendrier annuel** de présence.
- **Congés** : solde acquis / pris / restant.
- **Avances** sur salaire avec échéancier et prélèvement mensuel.
- **Bulletins de paie** (PDF & HTML), **attestations** de travail / de salaire,
  **contrat de travail** (CDI/CDD), avec logo et signature/cachet.
- **Numérotation séquentielle** des documents + **registre**.
- **État de paie** mensuel (HTML / Excel), **déclaration CNSS**.
- **Tableau de bord**, graphique des salaires, évolution de la masse salariale.
- Import / export **CSV**, **sauvegardes automatiques**, **mode sombre**,
  **mot de passe** et rôles (admin / comptable), **documents** par employé.

## Données

Les données sont stockées dans `employes.xlsx` (feuille « Employes ») dans le
dossier de l'utilisateur. La configuration (société, logo, taux, thème…) est
dans `~/.employee_manager.json`. Ces fichiers ne sont pas inclus dans le dépôt.

## Lancer depuis le code source

```bash
pip install -r requirements.txt
python employee_manager.py
```

## Générer l'exécutable (.exe)

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name GestionEmployes --icon icon.ico employee_manager.py
```

L'exécutable est créé dans `dist/`.

## Générer l'installateur (Setup.exe)

Avec [Inno Setup](https://jrsoftware.org/isinfo.php) installé :

```bash
ISCC.exe installer.iss
```

L'installateur est créé dans `installer/`.

## Dépendances

- Python 3.10+
- openpyxl, Pillow, fpdf2 (voir `requirements.txt`)
