<#
Infos
=====

    :Projet:             msgTracker365
    :Nom du fichier:     msgTracker365.ps1
    :depot_GitHub:       https://github.com/poltergeist42/msgTracker365
    :Documentation:      https://poltergeist42.github.io/msgTracker365/
    :Auteur:            `Poltergeist42 <https://github.com/poltergeist42>`_
    :Version:            20170920

####

    :Licence:            CC-BY-NC-SA
    :Liens:              https://creativecommons.org/licenses/by-nc-sa/4.0/

####

    :dev language:      Powershell
    :framework:         
    
####

Descriptif
==========

    :Projet:            Ce projet est un projet PowerShell. L'objectif est de cr�er un
                        Script qui se connecte automatiquement � Office365, interroge le
                        Suivie de Message et envoie automatiquement le r�sultat par mail

####

Reference Web
=============

    * https://support.office.com/fr-fr/article/Gestion-d-Office-365-et-d-Exchange-Online-avec-Windows-PowerShell-06a743bb-ceb6-49a9-a61d-db4ffdf54fa6
        # Gestion d'Office 365 et d'Exchange Online avec Windows PowerShell
        
    * https://technet.microsoft.com/library/dn975125.aspx
        # Se connecter � Office 365 PowerShell
    
    * https://technet.microsoft.com/fr-fr/library/dn568015.aspx
        # Connexion �  tous les services Office 365
        # � l'aide d'une seule fen�tre Windows PowerShell
        
    * https://technet.microsoft.com/EN-US/library/dn621038(v=exchg.160).aspx
        # Liste des commandes pour Exchange Online (module MsOnline)
    
####

Liste des modules externes
==========================

    * MsOnline

####

#>

cls
Write-Host "`t## D�but du script : msgTracker365 ##"


#########################
#                       #
#     Configuration     #
#                       #
#########################

## Param�tres pour l'authentification sur Office365 / Exchange Online
$vCfgUser365 = "user@domain.dom"
    # Login utilis� pour l'authentification sur le compte Office365 / Exchange Online
    #
    # N.B : Le compte Office365 utilis� doit au minimum faire partie des groupes :
    # * Gestion de l�organisation (Organization management)
    # * Gestion de la conformit� (Compliance Management)
    #
    # Attention : "user@domain.dom" doit �tre remplac� par votre nom d'utilisateur
    # dans la version en production de ce script
    
$vCfgPwd365 = "P@sSwOrd"
    # Mot de passe utiliser avec le login du compte  Office365 / Exchange Online
    #
    # Attention : "P@sSwOrd" doit �tre remplac� par votre mot de passe
    # dans la version en production de ce script
   
   
## Param�tre pour le domaine � auditer
$vDomain = "*@domain.dom"
    # Nom de domaine � auditer  sur Office365 / Exchange Online
    #
    # Attention : "*@domain.dom" doit �tre remplac� par votre domaine
    # dans la version en production de ce script
    

## Param�tre de configuration de la p�riode � auditer (du plus ancien au plus r�cent)
$vCfgStartDate = 7
    # Cette valeur (en nombre de jour) permet de d�finir la date � partir de laquelle on
    # r�cup�re les informations. Il s'agit de la date la plus ancienne. Cette date ne peut
    # d�passer 90 jours

$vCfgEndDate = 0
    # Cette valeur (en nombre de jour) permet de d�finir la date jusqu'� laquelle on
    # r�cup�re les informations. Il s'agit de la date la plus r�cente. Si cette valeur est
    # �gale � 0, la date de fin sera la date actuelle

    
## Param�tre pour la quantit� d'�l�ment � afficher dans le suivie de message
$vCfgNumOfPage = 1
    # Cette valeur permet de d�terminer le nombre de pages affich�es
    # dans le suivie de message. Cette valeur est comprise entre 1 et 5000.
    # La valeur par d�faut est de 1.
    
$vCfgItemPerPage = 1000
    # Cette valeur permet de d�terminer le nombre d'�l�ment � afficher
    # dans le suivie de message. Cette valeur est comprise entre 1 et 5000.
    # La valeur par d�faut est de 1000.

    
## Param�tres pour la g�n�ration des fichiers    
$vCfgPath = ".\"
    # Chemin utilis� pour enregistrer les fichiers identifi�
    #par $vCfgExpCSV et $vCfgExpBody. Si 'vCfgPath' vaut '.\', les fichiers seront cr�es
    # dans le r�pertoire d'ex�cution de ce script
    #
    # N.B : le chemin doit exister sur le PC avant l'ex�cution de se script. Ce chemin
    # peut �tre relatif ou absolu
    
$vCfgExpCSV = "Suivie_de_message.csv"
    # Nom du fichier contenant le r�sultat de la requ�te. Les valeurs contenues dans se
    # fichier sont s�parer par des virgules. La convention veut donc que l'extension
    # du fichier soit au format "CSV" (Comma-separated values). Ce fichier est g�n�r�
    # � l'endroit point� par "$vCfgPath"
    
$vCfgExpBody = "Body.txt"
    # Nom du fichier contenant le corps de l'email. Ce fichier peut �tre utilise pour
    # envoyer un email depuis un logiciel tiers (ex : smtpsend). Ce fichier est g�n�r�
    # � l'endroit pointe par "$vCfgPath"
    
$vCfgEncoding = "Default"
    # Permet de d�finir l'encodage des fichiers et du mail. La valeur "Default",
    # r�cup�re l'encodage du syst�me depuis lequel est ex�cuter ce script.
    # Les valeurs accept�es sont :
    # "Unicode", "UTF7", "UTF8", "ASCII", "UTF32",
    # "BigEndianUnicode", "Default", "OEM"

    
## Param�tre de configuration de l'envoie de Mail
$vCfgSendMail = $False
    # Permet d'activer ou de d�sactiver l'envoie automatique du fichier '.csv' par mail.
    # Les valeurs accept�es sont :
    # * $TRUE   --> Envoie de mail activ�
    # * $FALSE  --> Envoie de mail d�sactiv�
    
$vCfgSendMailFrom = "user01@example.com"
    # Adresse mail de l'exp�diteur
    #
    # Attention : "user01@example.com" doit �tre remplac� par l'adresse de l'exp�diteur
    # dans la version en production de ce script
    
$vCfgSendMailTo = "user02@example.com"
    # Adresse Mail du destinataire
    #
    # Attention : "user02@example.com" doit �tre remplac� par l'adresse du destinataire
    # dans la version en production de ce script
    
$vCfgSendMailSmtp = "smtp.serveur.com"
    # Serveur SMTP � utiliser pour l'envoie de Mail
    #
    # Attention : "smtp.serveur.com" doit �tre remplac� par votre serveur SMTP
    # dans la version en production de ce script
    
$vCfgSendMailPort = 25
    # Num�ro de port utilis� par le serveur SMTP
    
$vCfgSendMailAuth = $TRUE
    # Permet d'activer ou de d�sactiver l'authentification sur le SMTP.
    # Les valeurs accept�es sont :
    # * $FALSE  --> Pas d'authentification
    # * $TRUE   --> Authentification
    #
    # N.B : Si le serveur SMTP n�cessite une authentification ($TRUE), les variables :
    # 'vCfgSendMailUsr' et 'vCfgSendMailPwd' seront �galement � renseigner
    
$vCfgSendMailUsr = "user@domain.dom"
    # Login utilis� pour l'authentification du SMTP
    #
    # Attention : "user@domain.dom" doit �tre remplac� par votre nom d'utilisateur
    # dans la version en production de ce script
    

$vCfgSendMailPwd = "P@sSwOrd"
    # Mot de passe utiliser avec le login du compte  d'authentification SMTP
    #
    # Attention : "P@sSwOrd" doit �tre remplace par votre mot de passe
    # dans la version en production de ce script
    

##########################
#                        #
# Variables de requ�te   #
#                        #
##########################

$vPwd365 = ConvertTo-SecureString -String $vCfgPwd365 -AsPlainText -Force
$Credential365 = New-Object -TypeName "System.Management.Automation.PSCredential" -ArgumentList $vCfgUser365, $vPwd365
    # il faut utiliser l'option : -Credential $Credential365
    # Pour pouvoir l'utiliser dans une requ�te

if ($vCfgSendMailAuth) {
    $vMailPwd = ConvertTo-SecureString -String $vCfgSendMailPwd -AsPlainText -Force
    $CredentialMail = New-Object -TypeName "System.Management.Automation.PSCredential" -ArgumentList $vCfgSendMailUsr, $vMailPwd
    }
    
$vStartDate = (Get-Date).adddays(-1 * $vCfgStartDate)
$vEndDate = (Get-Date).adddays(-1 * $vCfgEndDate)

$vStartDateShort = $vStartDate.ToShortDateString()
$vEndDateShort = $vEndDate.ToShortDateString()

$vCSV_FQFN = "$vCfgPath\$vCfgExpCSV"
$vBody_FQFM = "$vCfgPath\$vCfgExpBody"
    
    
##########################
#                        #
# Requ�tes et formatage  #
#                        #
##########################

## Connection �  office 365
Import-Module MsOnline
Connect-MsolService -Credential $Credential365


## Connection a Exchange Online
$exchangeSession =  New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/"`
-Credential $Credential365 -Authentication "Basic" -AllowRedirection

Import-PSSession $exchangeSession -DisableNameChecking

## Suivie de message
$vMsgTrace = Get-MessageTrace -RecipientAddress $vDomain -StartDate $vStartDate -EndDate $vEndDate -Page $vCfgNumOfPage -PageSize $vCfgItemPerPage | Sort-Object  `
-Property Received | Select-Object Received, RecipientAddress, SenderAddress, Subject
    # Attention, Get-MessageTrace parcourt la liste du plus ancien (StartDate) vers le plus r�cent (EndDate)

$vMsgTraceMeasure = $vMsgTrace | measure
$vMsgTraceCount = $vMsgTraceMeasure.Count

$vMsgTrace | Export-Csv -Path $vCSV_FQFN  -Delimiter ";" -Encoding $vCfgEncoding
$vBody = "Bonjour.`n`nNombre total de courriels re�u entre $vStartDateShort et $vEndDateShort pour le domaine `'$vDomain`' : $vMsgTraceCount`n`nVous trouverez le d�tail dans la pi�ce jointe nomm�e : `'$vCfgExpCSV`'`n`nCordialement, l'�quipe ICS"

$vBody | Out-File -FilePath $vBody_FQFM -Encoding $vCfgEncoding


##########################
#                        #
#     Envoie de Mail     #
#                        #
##########################

if ($vCfgSendMail) {
    if ($vCfgSendMailAuth) {
        Send-MailMessage -From $vCfgSendMailFrom `
        -Encoding $vCfgEncoding `
        -To $vCfgSendMailTo `
        -Subject "Rapport de Suivie de message entre le $vStartDateShort et le $vEndDateShort" `
        -Body $vBody `
        -Attachments $vCSV_FQFN `
        -Credential $CredentialMail `
        -SmtpServer $vCfgSendMailSmtp `
        -Port $vCfgSendMailPort
    }
    else {
        Send-MailMessage -From $vCfgSendMailFrom `
        -Encoding $vCfgEncoding `
        -To $vCfgSendMailTo `
        -Subject "Rapport de Suivie de message entre le $vStartDateShort et le $vEndDateShort" `
        -Body $vBody `
        -Attachments $vCSV_FQFN `
        -SmtpServer $vCfgSendMailSmtp `
        -Port $vCfgSendMailPort
    }}

##########################
#                        #
#     Fin de tache       #
#                        #
##########################

## Fermeture de toutes les sessions distantes (les PSSession)
Get-PSSession | Remove-PSSession

Write-Host "`n`t## Fin du script : msgTracker365 ##"
