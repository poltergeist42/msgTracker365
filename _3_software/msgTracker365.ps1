<#
Infos
=====

    :Projet:             msgTracker365
    :Nom du fichier:     msgTracker365.ps1
    :depot GitHub:       
    :documentation:      
    :Autheur:            `Poltergeist42 <https://github.com/poltergeist42>`_
    :Version:            20170912

####

    :Licence:            CC-BY-NC-SA
    :Liens:              https://creativecommons.org/licenses/by-nc-sa/4.0/

####

    :dev language:      Powershell
    :framework:         
    
####

Descriptif
==========

    :Projet:            Se projet est un projet PowerShell. L'objectif est de creer un
                        script qui se connect automatiquement a Office365, interoge le
                        Suivie de Message et envoie automatiquement le resultat par mail

####

Reference Web
=============

    * https://support.office.com/fr-fr/article/Gestion-d-Office-365-et-d-Exchange-Online-avec-Windows-PowerShell-06a743bb-ceb6-49a9-a61d-db4ffdf54fa6
        # Gestion d'Office 365 et d'Exchange Online avec Windows PowerShell
        
    * https://technet.microsoft.com/library/dn975125.aspx
        # Se connecter à Office 365 PowerShell
    
    * https://technet.microsoft.com/fr-fr/library/dn568015.aspx
        # Connexion a  tous les services Office 365
        # a l'aide d'une seule fenetre Windows PowerShell
    
####

#>

cls
Write-Host "Debut du script"


#########################
#                       #
#     Configuration     #
#                       #
#########################

$vCfgUser365 = "user@domain.dom"
    # Login utilise pour l'authentification sur le compte Office365 / Exchange Online
    #
    # N.B : Le compte Office365 utilise doit au minimum faire partie des groupes :
    # * Gestion de l’organisation (Organization management)
    # * Gestion de la conformite (Compliance Management)
    #
    # Attention : "user@domain.dom" doit etre remplacer par votre nom d'utilisateur
    # dans la version en production de ce script
    
$vCfgPwd365 = ConvertTo-SecureString -String "P@sSwOrd" -AsPlainText -Force
    # Mot de passe utiliser avec le login du compte  Office365 / Exchange Online
    #
    # Attention : "P@sSwOrd" doit etre remplace par votre mot de passe
    # dans la version en production de ce script
    
$vDomain = "*@domain.dom"
    # Nom de domaine a auditer  sur Office365 / Exchange Online
    #
    # Attention : "*@domain.dom" doit etre remplacer par votre domain
    # dans la version en production de ce script
    
$vCfgStartDate = 7
    # Cette valeur (en nombre de jour) permet de definir la date a partir de la quelle on
    # recupere les informations. Il s'agit de la date la plus ancienne. Cette date ne peut
    # depasser 90 jours

$vCfgEndDate = 0
    # Cette valeur (en nombre de jour) permet de definir la date jusqu'a laquelle on
    # récupère les informations. il s'agit de la date la plus recente. Si cette valeur est
    # egale a 0, la date de fin serat la date actuelle

$vCfgPath = "C:\utilSRV\Scripts\msgTracker"
    # Chemin utiliser pour enregistrer les fichiers identifier
    # par $vCfgExpCSV et $vCfgExpBody
    #
    # N.B : le chemin doit exister sur le PC avant l'execution de se script
    
$vCfgExpCSV = "Suivie_de_message.csv"
    # Nom du fichier contenant le resultat de la requette. Les valeurs contenues dans se
    # fichier sont separer par des virgules. la convention veut donc que l'extention
    # du fichier soit au format "CSV" (Comma-separated values). Ce fichier est genere
    # a l'endroit pointe par "$vCfgPath"
    
$vCfgExpBody = "Body.txt"
    # Nom du fichier contenant le corp de l'email. Ce fichier peut etre utilise pour
    # envoyer un email depuis un logiciel tiers (ex : smtpsend). Ce fichier est genere
    # a l'endroit pointe par "$vCfgPath"
    

##########################
#                        #
# variables de requette  #
#                        #
##########################

$Credential365 = New-Object -TypeName "System.Management.Automation.PSCredential" -ArgumentList $vCfgUser365, $vCfgPwd365
    # il faut utiliser l'option : -Credential $Credential365
    # Pour pouvoir l'utiliser dans une requette

$vStartDate = (Get-Date).adddays(-1 * $vCfgStartDate)
$vEndDate = (Get-Date).adddays(-1 * $vCfgEndDate)

$vEndDateShort = $vEndDate.ToShortDateString()
$vStartDateShort = $vStartDate.ToShortDateString()

$vCSV_FQFN = "$vCfgPath\$vCfgExpCSV"
$vBody_FQFM = "$vCfgPath\$vCfgExpBody"
    
    
##########################
#                        #
# Requettes et formatage #
#                        #
##########################

## Connection Ã  office 365
Import-Module MsOnline
Connect-MsolService -Credential $Credential365


## Connection a Exchange Online
$exchangeSession =  New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/"`
-Credential $Credential365 -Authentication "Basic" -AllowRedirection

Import-PSSession $exchangeSession -DisableNameChecking

## Suivie de message
$vMsgTrace = Get-MessageTrace -RecipientAddress $vDomain -StartDate $vStartDate -EndDate $vEndDate | Sort-Object `
-Property Received | Select-Object Received, RecipientAddress, SenderAddress
    # Attention, Get-MessageTrace parcour la liste du plus ancien (StartDate) vers le plus recent (EndDate)

$vMsgTraceMeasure = $vMsgTrace | measure
$vMsgTraceCount = $vMsgTraceMeasure.Count

$vMsgTrace | Export-Csv -Path $vCSV_FQFN
$vBody = "Bonjour.`n`nNombre total de courriels reçu entre $vStartDateShort et $vEndDateShort pour le domain` `'$vDomain`' : $vMsgTraceCount`n`nVous trouverez le détail dans la pièce jointe nommée : `'$vCfgExpCSV`'`n`nCordialement,` l'équipe ICS"

$vBody | Out-File -FilePath $vBody_FQFM

<# TODO :
envoyer automatiquement un mail avec $vBody dans le corp du message et $vCSV_FQFN en pièce jointe
#>


##########################
#                        #
#     Fin de tache       #
#                        #
##########################

Write-Host "Fin du script"
## Fermeture de toutes les sessions distantes (les PSSession)
Get-PSSession | Remove-PSSession

    