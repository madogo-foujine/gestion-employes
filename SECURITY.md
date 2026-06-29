# Sécurité

Cette application est un logiciel **de bureau, mono-poste**, sans serveur ni
réseau. Le modèle de menace se limite donc à un accès local à la machine.

## Mesures en place

- **Mots de passe** : hachés avec **scrypt** (`hashlib.scrypt`, fonction
  mémoire-dur), sel aléatoire de 16 octets, comparaison à temps constant
  (`secrets.compare_digest`). Les anciens hachages SHA-256 restent acceptés
  pour compatibilité, mais tout nouveau mot de passe utilise scrypt.
- **Validation de la configuration** : les valeurs lues depuis
  `~/.employee_manager.json` (taux, plafonds, barème IR…) sont contrôlées ;
  toute valeur invalide ou hors plage est ignorée au profit du défaut.
- **Validation des saisies** : les champs numériques doivent être des nombres
  positifs et les dates valides avant enregistrement.
- **Journalisation** : les erreurs d'écriture de fichiers sont enregistrées
  dans `~/.employee_manager.log` (plus de `except: pass` silencieux).
- Aucune utilisation d'`eval`, `exec`, `os.system`, `subprocess` ni `pickle`.

## Données au repos (limite connue)

Les données (`employes.xlsx`, fichiers JSON, documents) sont stockées en clair
dans le profil de l'utilisateur. Elles sont protégées par les **droits d'accès
du compte Windows**, mais **ne sont pas chiffrées**.

Pour un usage en entreprise, il est recommandé de :
- activer le chiffrement du disque (**BitLocker**) ;
- protéger la session Windows par mot de passe ;
- restreindre l'accès au dossier de données aux utilisateurs autorisés.

Un chiffrement applicatif complet nécessiterait un mot de passe maître pour
dériver une clé ; il n'est pas implémenté à ce jour.

## Signaler un problème

Ouvrez une *issue* sur le dépôt GitHub.
