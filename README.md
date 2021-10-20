# Get-MailboxesWithSuspiciousRule

Script PowerShell pour rechercher les boîtes aux lettres qui comporteraient une règle trouvée provenant d'une campagne d'hammeçonnage détectée.

## Prérequis

| Logiciel | Version |
| --- | --- |
| PowerShell | 3+ |
| [Module PowerShell « EXO » (ExchangeOnlineManagement)](https://www.powershellgallery.com/packages/ExchangeOnlineManagement/) | 2.X |

## Instructions

1. Installer le module de gestion d'Exchange Online avec la commande `Install-Module -Name ExchangeOnlineManagement` dans une console PowerShell avec privilèges administratifs
2. Lancer le script avec la commande `.\Get-MailboxesWithSuspiciousRule.ps1`
3. Répondre aux questions pour fournir le nom d'utilisateur exact à utiliser (ex.: `admin@contoso.com`) puis le nom exact de la règle à chercher (ex.: `///`)
4. Attendre et vérifier le résultat dans le fichier `BoîtesCompromises.csv` qui est créé dans le répertoire de travail
