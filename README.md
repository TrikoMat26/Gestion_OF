# Gestionnaire d'Ordres de Fabrication (OF) Multi-Clients

Ce projet est un outil autonome permettant de gérer les associations entre les Ordres de Fabrication (OF) et les Numéros de Série (SN) pour différents clients.

## Contenu du Projet

- **Gestion_OF_MultiClients.ps1** : Script principal avec interface graphique (WinForms).
- **Global_OF_Registry.json** : Base de données globale structurée par client.
- **Liste_OF.ps1** : Utilitaire de segmentation et de détection de manquants dans les listes de SN.
- **OF_Registry.json** : (Optionnel) Ancien registre Wattsy pour importation initiale.

## Fonctionnalités

1. **Gestion Multi-Clients** : Déclarez plusieurs clients et gardez leurs OF séparés.
2. **Importation Facilitée** : Bouton d'importation pour migrer les données depuis l'ancien format.
3. **Ergonomie** : Actions rapides par double-clic et touches clavier (`Entrée`, `Echap`).
4. **Recherche Globale** : Retrouvez instantanément à quel client et quel OF appartient un numéro de série.

## Utilisation

Faites un clic droit sur `Gestion_OF_MultiClients.ps1` et choisissez **"Exécuter avec PowerShell"**.
