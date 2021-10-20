# Prérequis : Install-Module -Name ExchangeOnlineManagement
# Modifier la variable $user avec la bonne adresse courriel

function DisconnectAndExit
{
	Write-Host "Déconnexion d'Exchange Online..."
	Disconnect-ExchangeOnline
	Write-Host "Exchange Online déconnecté."
	Return
}

function NewConsoleLine
{
	Write-Host
}

try
{
	Write-Host "Importation du module de gestion d'Exchange Online..."
	Import-Module ExchangeOnlineManagement
	Write-Host "Module importé."
}
catch
{
	Write-Error "Impossible d'importer le module « ExchangeOnlineManagement ». Veuillez l'installer à l'aide de la commande « Install-Module -Name ExchangeOnlineManagement »."
	Return
}

try
{
	# Source : https://docs.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps
	$user = Read-Host -Prompt "Entrez votre nom d'utilisateur"
	Write-Host "Connexion à ExchangeOnline..."
	Write-Warning "Il se peut qu'une fenêtre de connexion à Microsoft apparaîsse pour l'authentification multi-facteur. S.v.p. y répondre pour compléter la connexion."
	Connect-ExchangeOnline -UserPrincipalName $script:user
	Write-Host "Connecté à ExchangeOnline avec $script:user."
	NewConsoleLine
}
catch
{
	Write-Error "Échec de connexion. Vérifiez le nom d'utilisateur et l'accès au serveur."
	Return
}

$suspiciousRule = Read-Host -Prompt "Entrez le nom exact de la règle à rechercher"
Write-Host "Nom de la règle suspecte à rechercher : $script:suspiciousRule"
NewConsoleLine

Write-Host "Obtention des boîtes courriel..."
$list = (Get-Mailbox -ResultSize unlimited -SortBy Name | Sort-Object -Property Name | Select-Object -ExpandProperty Name)
Write-Host "$($script:list.Count) adresses trouvées."
NewConsoleLine

$outputFile = "BoîtesCompromises.csv"
Write-Host "Création du fichier CSV pour lister les comptes compromis..."
try {
	Set-Content -Path $script:outputFile -Value "Alias"
	Write-Host "Fichier CSV de liste de comptes compromis créé : $script:outputFile ."
	Write-Host "Ce fichier contiendra tout les comptes qui ont été compromis avec la règle suspecte."
	NewConsoleLine
}
catch {
	Write-Error "Impossible d'écrire le fichier. Vérifiez les droits d'accès à $script:outputFile puis réessayez."
	DisconnectAndExit
}

$count = 0
$index = 0
Write-Host "Balayage à travers les comptes courriel..."
Write-Warning "Seulement les comptes comportant la règle avec le nom recherché seront affichés et seront écris dans le fichier."
NewConsoleLine
foreach ($alias in $script:list)
{

	Write-Progress -Id 1 -Activity "Balayage des règles des comptes courriels" -Status "Vérification des règles pour $script:alias" -PercentComplete ($script:index / $script:list.Count * 100)

	$rules = (Get-InboxRule -Mailbox $script:alias | Select-Object -ExpandProperty Name)
	if ($script:rules.Count -gt 0) {
		if ($script:rules.Contains($script:suspiciousRule))
		{
			Write-Host $script:alias
			Add-Content -Path $script:outputFile -Value $script:alias
			$script:count++
		}
	}
	$script:index++
}
Write-Progress -Id 1 -Activity "Balayage des règles des comptes courriels" -PercentComplete (100)

Write-Host "Script terminé. $script:count boîtes aux lettres semblent être compromises."
Get-Content $script:outputFile
DisconnectAndExit
