====================================
Informations générales msgTracker365
====================================

:Auteur:            `Poltergeist42 <https://github.com/poltergeist42>`_
:Projet:             msgTracker365
:Dépôt_GitHub:       https://github.com/poltergeist42/msgTracker365
:Documentation:      https://poltergeist42.github.io/msgTracker365/
:Licence:            CC BY-NC-SA 4.0
:Liens:              https://creativecommons.org/licenses/by-nc-sa/4.0/

####

Descriptions
===========

Ce projet est un projet PowerShell. L'objectif est de créer un script qui se connecte
automatiquement à Office365, interroge le Suivie de Message et envoie automatiquement le
résultat par mail.

####

Téléchargement / Installation
=============================

Vous pouvez télécharger le projet entier directement depuis son `dépôt GitHub <https://github.com/poltergeist42/msgTracker365.git>`_ .
ou récupérer juste le script depuis le dossier `_3_software du dépôt GitHub <https://github.com/poltergeist42/msgTracker365/tree/master/_3_software>`_ .

Le script n'a pas besoin d'installation, il doit simplement être exécuté. Voir 'Prérequis' et 'Utilisation'
   
####   
 
Prérequis
=========

    #. Ce script vous permettant de vous connecter à Office365 au travers de PowerShell,
       il faut avoir installé les logiciels de l'étape 1 de la page Web suivante :
       
        * https://technet.microsoft.com/library/dn975125.aspx
    
    #. Vous devez disposer d'un compte Office365 ayant les autorisations ci-dessous,
       Conformément au tableau décris dans le `document Technet <https://technet.microsoft.com/fr-fr/library/jj200673(v=exchg.150).aspx>`_ :

        * Gestion de l’organisation (Organization management)
        * Gestion de la conformité (Compliance Management)
    
    #. Votre machine doit être autorisée à exécuter des scripts (`Set-ExecutionPolicy <https://docs.microsoft.com/fr-fr/powershell/module/Microsoft.PowerShell.Security/Set-ExecutionPolicy?view=powershell-5.1>`_)

####
    
Utilisation
===========

    #. **Personnalisation**
    
        Se script doit être personnalisé. Pour faciliter cette personnalisation, l'ensemble
        des variables à modifier ont été placées au début du script sous
        l'entête "Configuration".
       
        **Détail des variables à personnaliser** :
       
        :vCfgUser365:
            Login utilisé pour l'authentification sur le compte Office365 / Exchange Online.

            **N.B** : Le compte Office365 utilisé doit au minimum faire partie des groupes :
                * Gestion de l’organisation (Organization management)
                * Gestion de la conformité (Compliance Management)

            **Attention** : "user@domain.dom" doit être remplacé par votre nom
            d'utilisateur dans la version en production de ce script.
            
        :vCfgPwd365:
            Mot de passe utiliser avec le login du compte  Office365 / Exchange Online.

            **Attention** : "P@sSwOrd" doit être remplacé par votre mot de passe
            dans la version en production de ce script.
            
        :vDomain:
            Nom de domaine a auditer  sur Office365 / Exchange Online.

            **Attention** : "*@domain.dom" doit être remplacé par votre domaine
            dans la version en production de ce script.
    
        :vCfgStartDate:
            Cette valeur (en nombre de jour) permet de définir la date a partir de
            la quelle on récupère les informations. Il s'agit de la date la plus ancienne.
            Cette date ne peut dépasser 90 jours.
            
        :vCfgNumOfPage:
            Cette valeur permet de déterminer le nombre de pages affichées
            dans le suivie de message. Cette valeur est comprise entre 1 et 5000.
            La valeur par défaut est de 1.
    
        :vCfgItemPerPage:
            Cette valeur permet de déterminer le nombre d'élément à afficher
            dans le suivie de message. Cette valeur est comprise entre 1 et 5000.
            La valeur par défaut est de 1000.
            
        :vCfgEndDate:
            Cette valeur (en nombre de jour) permet de définir la date jusqu'a laquelle on
            récupère les informations. Il s'agit de la date la plus récente. Si cette
            valeur est égale a 0, la date de fin sera la date actuelle.
            
        :vCfgPath:
            Chemin utilisé pour enregistrer les fichiers identifié
            par $vCfgExpCSV et $vCfgExpBody. Si 'vCfgPath' vaut '.\',
            les fichiers seront créés dans le répertoire d'exécution de ce script.

            **N.B** : le chemin doit exister sur le PC avant l'exécution de se script.
            Ce chemin peut être relatif ou absolu.
            
        :vCfgExpCSV:
            Nom du fichier contenant le résultat de la requête. Les valeurs contenues
            dans ce fichier sont séparée par des virgules. La convention veut donc que
            l'extension du fichier soit au format "CSV" (Comma-separated values). Ce
            fichier est généré à l'endroit pointer par "$vCfgPath".
            
        :vCfgExpBody:
            Nom du fichier contenant le corps de l'email. Ce fichier peut être utilisé
            pour envoyer un email depuis un logiciel tiers (ex : smtpsend).
            Ce fichier est généré à l'endroit pointé par "$vCfgPath".
            
        :vCfgEncoding:
            Permet de définir l'encodage des fichiers et du mail. La valeur "Default",
            récupère l'encodage du système depuis lequel est exécuter ce script.
            Les valeurs acceptées sont :
                * "Unicode", "UTF7", "UTF8", "ASCII", "UTF32", "BigEndianUnicode", "Default", "OEM"
    
        :vCfgSendMail:
            Permet d'activer ou de désactiver l'envoie automatique du fichier '.csv' par
            mail. Les valeurs acceptées sont :
                * $TRUE   --> Envoie de mail activé
                * $False  --> Envoie de mail désactivé
                
            **N.B** : Cette fonctionnalité est désactivé par défaut ($FALSE)
    
        :vCfgSendMailFrom:
            Adresse mail de l'expéditeur

            **Attention** : "user01@example.com" doit être remplacé par l'adresse de
            l'expéditeur dans la version en production de ce script.
    
        :vCfgSendMailTo:
            Adresse Mail du destinataire

            **Attention** : "user02@example.com" doit être remplacé par l'adresse
            du destinataire dans la version en production de ce script.
    
        :vCfgSendMailSmtp:
            Serveur SMTP à utiliser pour l'envoie de Mail.

            **Attention** : "smtp.serveur.com" doit être remplacé par votre serveur SMTP
            dans la version en production de ce script.
    
        :vCfgSendMailPort:
            Numéro de port utilisé par le serveur SMTP.
        
        :vCfgSendMailAuth:
            Permet d'activer ou de désactiver l'authentification sur le SMTP.
            Les valeurs acceptées sont :
                * $FALSE  --> Pas d'authentification
                * $TRUE   --> Authentification

            **N.B** : Si le serveur SMTP nécessite une authentification ($TRUE),
            les variables : 'vCfgSendMailUsr' et 'vCfgSendMailPwd'
            seront également à renseigner. Ce mode est activé par défaut en cas d'envoie
            automatique d'un mail depuis se script.
    
        :vCfgSendMailUsr:
            Login utilisé pour l'authentification du SMTP.

            **Attention** : "user@domain.dom" doit être remplacé par votre nom
            d'utilisateur dans la version en production de ce script.

        :vCfgSendMailPwd:
            Mot de passe utiliser avec le login du compte  d'authentification SMTP.

            **Attention** : "P@sSwOrd" doit être remplace par votre mot de passe
            dans la version en production de ce script.
    

    
    #. **Automatisation et planification**
    
        Si la tâche doit être effectuée régulièrement, il faut créer une tache planifié.
        On peut s'aider de la page ci-dessous pour exécuter un script PowerShell dans une
        tâche planifiée.
        
            * https://www.adminpasbete.fr/executer-script-powershell-via-tache-planifiee/
    
Arborescence du projet
======================

Pour aider à la compréhension de mon organisation, voici un bref descriptif de
L'arborescence de ce projet. Cette arborescence est à reproduire si vous récupérez ce dépôt
depuis GitHub. ::

	openFile               # Dossier racine du projet (non versionner)
	|
	+--project             # (branch master) contient l'ensemble du projet en lui même
	|  |
	|  +--_1_userDoc       # Contiens toute la documentation relative au projet
	|  |   |
	|  |   \--source       # Dossier réunissant les sources utilisées par Sphinx
	|  |
	|  +--_2_modelisation  # Contiens tous les plans et toutes les modélisations du projet
	|  |
	|  +--_3_software      # Contiens toute la partie programmation du projet
	|  |
	|  \--_4_PCB           # Contient toutes les parties des circuits imprimés (routage,
	|                      # Implantation, typon, fichier de perçage, etc.
	|
	\--webDoc              # Dossier racine de la documentation qui doit être publiée
	   |
	   \--html             # (branch gh-pages) C'est dans ce dossier que Sphinx vat
	                       # générer la documentation à publier sur internet

