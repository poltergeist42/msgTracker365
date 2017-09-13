====================================
Informations générales msgTracker365
====================================

:Autheur:            `Poltergeist42 <https://github.com/poltergeist42>`_
:Projet:             msgTracker365
:dépôt GitHub:       https://github.com/poltergeist42/msgTracker365
:documentation:      https://poltergeist42.github.io/msgTracker365/
:Licence:            CC BY-NC-SA 4.0
:Liens:              https://creativecommons.org/licenses/by-nc-sa/4.0/

####

Déscription
===========

Se projet est un projet PowerShell. L'objectif est de créer un script qui se connect
automatiquement à Office365, intéroge le Suivie de Message et envoie automatiquement le
résultat par mail

**N.B** : dans la version "20170913", l'envoie de mail doit être effectué depuis
un logiciel tiers.
 
####

Téléchargement / Installation
=============================

Vous pouvez télécharger le projet entier directement depuis son `dépôt GitHub <https://github.com/poltergeist42/msgTracker365.git>`_ .
ou récupérer juste le script depuis le dossier `_3_software du dépôt GitHub <https://github.com/poltergeist42/msgTracker365/tree/master/_3_software>`_ .

Le script n'a pas besoin d'installation, il doit simplement être éxécuté. Voir 'Prérequis' et 'Utilisation'
   
####   
 
Prérequis
=========

    #. Ce script vous permettant de vous connecter à Office365 au travers de PowerShell,
       il faut avoir installer les logiciels de l'étape 1 de la page Web suivante :
       
        * https://technet.microsoft.com/library/dn975125.aspx
    
    #. Vous devez disposez d'un compte Office365 ayant les autorisations ci-dessous,
       conformément aut tableau décris dans le `document Technet <https://technet.microsoft.com/fr-fr/library/jj200673(v=exchg.150).aspx>`_ :

        * Gestion de l’organisation (Organization management)
        * Gestion de la conformité (Compliance Management)
    
    #. Votre machine doit être autorisée à éxecuter des scripts (`Set-ExecutionPolicy <https://docs.microsoft.com/fr-fr/powershell/module/Microsoft.PowerShell.Security/Set-ExecutionPolicy?view=powershell-5.1>`_)

####
    
Utilisation
===========

    #. **Personalisation**
    
        Se script doit être personalisé. Pour faciliter cette personnalisation, l'ensemble
        des variables à modifier ont été placées au début du script sous
        l'entête "Configuration".
       
        **Détail des variables à personaliser** :
       
        :vCfgUser365:
            Login utilisé pour l'authentification sur le compte Office365 / Exchange Online.

            **N.B** : Le compte Office365 utilisé doit au minimum faire partie des groupes :
                * Gestion de l’organisation (Organization management)
                * Gestion de la conformité (Compliance Management)

            **Attention** : "user@domain.dom" doit être remplacer par votre nom
            d'utilisateur dans la version en production de ce script.
            
        :vCfgPwd365:
            Mot de passe utiliser avec le login du compte  Office365 / Exchange Online.

            **Attention** : "P@sSwOrd" doit être remplacé par votre mot de passe
            dans la version en production de ce script.
            
        :vDomain:
            Nom de domaine a auditer  sur Office365 / Exchange Online.

            **Attention** : "*@domain.dom" doit être remplacer par votre domain
            dans la version en production de ce script.
    
        :vCfgStartDate:
            Cette valeur (en nombre de jour) permet de définir la date a partir de
            la quelle on récupere les informations. Il s'agit de la date la plus ancienne.
            Cette date ne peut depasser 90 jours.
            
        :vCfgEndDate:
            Cette valeur (en nombre de jour) permet de définir la date jusqu'a laquelle on
            récupère les informations. il s'agit de la date la plus recente. Si cette
            valeur est égale a 0, la date de fin serat la date actuelle.
            
        :vCfgPath:
            Chemin utilisé pour enregistrer les fichiers identifié
            par $vCfgExpCSV et $vCfgExpBody. Si 'vCfgPath' vaut '.\',
            les fichiers seront créés dans le répertoire d'éxecution de ce script.

            **N.B** : le chemin doit éxister sur le PC avant l'éxecution de se script.
            Ce chemin peut être rélatif ou absolu.
            
        :vCfgExpCSV:
            Nom du fichier contenant le résultat de la requette. Les valeurs contenues
            dans se fichier sont séparée par des virgules. la convention veut donc que
            l'extention du fichier soit au format "CSV" (Comma-separated values). Ce
            fichier est généré à l'endroit pointer par "$vCfgPath".
            
        :vCfgExpBody:
            Nom du fichier contenant le corp de l'email. Ce fichier peut être utilisé
            pour envoyer un email depuis un logiciel tiers (ex : smtpsend).
            Ce fichier est généré à l'endroit pointé par "$vCfgPath".
    
    #. **Automatisation et planification**
    
        Si la tâche doit être effectuée régulierement, il faut créer une tache planifié.
        On peut s'aider de la page cidessous pour executer un script PowerShell dans une
        tâche planifiée.
        
            * https://www.adminpasbete.fr/executer-script-powershell-via-tache-planifiee/
    
Arborescence du projet
======================

Pour aider à la compréhension de mon organisation, voici un bref déscrptif de
l'arborescence de se projet.Cette arborescence est à reproduire si vous récupérez ce dépôt
depuis GitHub. ::

	openFile               # Dossier racine du projet (non versionner)
	|
	+--project             # (branch master) contient l'ensemble du projet en lui même
	|  |
	|  +--_1_userDoc       # Contien toute la documentation relative au projet
	|  |   |
	|  |   \--source       # Dossier réunissant les sources utilisées par Sphinx
	|  |
	|  +--_2_modelisation  # contien tous les plans et toutes les modélisations du projet
	|  |
	|  +--_3_software      # Contien toute la partie programmation du projet
	|  |
	|  \--_4_PCB           # Contient toutes les parties des circuits imprimés (routage,
	|                      # implantation, typon, fichier de perçage, etc
	|
	\--webDoc              # Dossier racine de la documentation qui doit être publiée
	   |
	   \--html             # (branch gh-pages) C'est dans se dosier que Sphinx vat
	                       # générer la documentation à publié sur internet

