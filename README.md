# Synchronisation locale vers SharePoint avec Microsoft Graph

Ce script Python synchronise récursivement le contenu d’un répertoire local vers une bibliothèque de documents SharePoint en utilisant Microsoft Graph et une authentification applicative Microsoft Entra ID.

Le script permet également de récupérer et d’afficher les `Drive ID` des bibliothèques de documents disponibles sur le site SharePoint.

## Fonctionnalités

- Authentification auprès de Microsoft Entra ID avec `CLIENT_ID`, `CLIENT_SECRET` et `TENANT_ID`.
- Récupération automatique du `Site ID` SharePoint.
- Affichage de toutes les bibliothèques SharePoint du site avec :
  - le nom de la bibliothèque ;
  - le `Drive ID` ;
  - le type de lecteur ;
  - l’URL web.
- Création automatique des dossiers SharePoint manquants.
- Parcours récursif du répertoire local.
- Téléversement des fichiers dans SharePoint en conservant l’arborescence locale.
- Journalisation dans un fichier horodaté et dans la console.
- Fonction disponible pour supprimer le contenu d’un dossier SharePoint.

## Prérequis

- Python 3.9 ou supérieur recommandé.
- Une application enregistrée dans Microsoft Entra ID.
- Un secret client valide.
- Des permissions Microsoft Graph adaptées au mode applicatif.
- Un accès réseau aux services Microsoft Entra ID et Microsoft Graph.

## Installation

Créer un environnement virtuel :

```bash
python -m venv .venv
```

Sous Windows PowerShell :

```powershell
.\.venv\Scripts\Activate.ps1
```

Installer les dépendances :

```bash
pip install requests msal
```

Une alternative consiste à créer un fichier `requirements.txt` :

```text
msal
requests
```

Puis à exécuter :

```bash
pip install -r requirements.txt
```

## Configuration Microsoft Entra ID

L’application doit être enregistrée dans Microsoft Entra ID et disposer de permissions Microsoft Graph applicatives permettant d’accéder aux sites et aux fichiers SharePoint concernés.

Selon la politique de sécurité et le périmètre attendu, les permissions sont généralement accordées à l’application, puis validées par un administrateur. Pour limiter l’accès à certains sites uniquement, privilégier un modèle d’autorisation restreint plutôt qu’un accès global à tous les sites.

## Configuration du script

Modifier les constantes au début du script :

```python
CLIENT_ID = 'VOTRE_CLIENT_ID'
CLIENT_SECRET = 'VOTRE_CLIENT_SECRET'
TENANT_ID = 'VOTRE_TENANT_ID'

AUTHORITY = f'https://login.microsoftonline.com/{TENANT_ID}'

SHAREPOINT_SITE_URL = (
    'https://graph.microsoft.com/v1.0/sites/'
    'votre-domaine-sharepoint:/sites/nom-du-site'
)

LOCAL_DIRECTORY = 'E:/workspace/EXTRACT/opt'
SHAREPOINT_BASE_FOLDER = '/data_tree'
```

### Signification des paramètres

- `CLIENT_ID` : identifiant de l’application Microsoft Entra ID.
- `CLIENT_SECRET` : secret associé à l’application.
- `TENANT_ID` : identifiant du tenant Microsoft Entra ID.
- `SHAREPOINT_SITE_URL` : endpoint Microsoft Graph utilisé pour récupérer le site SharePoint.
- `LOCAL_DIRECTORY` : répertoire local à synchroniser.
- `SHAREPOINT_BASE_FOLDER` : dossier de destination dans la bibliothèque SharePoint par défaut.

## Sécurité des identifiants

Ne pas conserver un véritable `CLIENT_SECRET` directement dans le code source ou dans Git.

Il est préférable d’utiliser des variables d’environnement :

```python
CLIENT_ID = os.environ['SP_CLIENT_ID']
CLIENT_SECRET = os.environ['SP_CLIENT_SECRET']
TENANT_ID = os.environ['SP_TENANT_ID']
```

Exemple sous PowerShell :

```powershell
$env:SP_CLIENT_ID = "votre-client-id"
$env:SP_CLIENT_SECRET = "votre-secret"
$env:SP_TENANT_ID = "votre-tenant-id"
python .\sharepoint_sync.py
```

Ajouter également les fichiers sensibles et les journaux au `.gitignore` :

```gitignore
.venv/
__pycache__/
*.pyc
.env
logs/
```

## Exécution

Enregistrer le script, par exemple sous le nom :

```text
sharepoint_sync.py
```

Puis lancer :

```bash
python sharepoint_sync.py
```

Le programme réalise les opérations suivantes :

1. création du fichier de journalisation ;
2. obtention d’un jeton Microsoft Graph ;
3. récupération du `Site ID` ;
4. récupération et affichage des bibliothèques et de leurs `Drive ID` ;
5. création éventuelle de `/data_tree` et de ses sous-dossiers ;
6. téléversement des fichiers du répertoire local.

## Exemple d’affichage des Drive ID

```text
================================================================================
SHAREPOINT DRIVES
================================================================================

Name     : Documents
Drive ID : b!xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
Type     : documentLibrary
Web URL  : https://tenant.sharepoint.com/sites/site/Documents

================================================================================
```

La valeur à saisir dans un champ `Drive ID` est la valeur complète affichée après `Drive ID`, généralement sous une forme commençant par `b!`.

## Bibliothèque SharePoint utilisée par le script

Les appels du script utilisent actuellement cette forme :

```text
/sites/{site_id}/drive/root:...
```

Cela cible la bibliothèque de documents par défaut du site SharePoint.

Si la destination se trouve dans une autre bibliothèque, il faut sélectionner son `Drive ID`, puis remplacer les endpoints concernés par :

```text
/drives/{drive_id}/root:...
```

Exemple :

```python
check_url = (
    f"https://graph.microsoft.com/v1.0/drives/{drive_id}"
    f"/root:{folder_path_encoded}"
)
```

Dans ce cas, les fonctions de synchronisation doivent recevoir `drive_id` en paramètre à la place de dépendre de la bibliothèque par défaut.

## Arborescence créée

Avec cette configuration :

```python
LOCAL_DIRECTORY = 'E:/workspace/EXTRACT/opt'
SHAREPOINT_BASE_FOLDER = '/data_tree'
```

Un fichier local comme :

```text
E:/workspace/EXTRACT/opt/site1/config/settings.json
```

sera envoyé vers :

```text
/data_tree/site1/config/settings.json
```

## Journaux

Les journaux sont créés dans :

```text
logs/synchro/
```

Le nom du fichier suit ce format :

```text
sharepoint_sync_YYYYMMDD_HHMMSS.log
```

Les mêmes informations sont aussi affichées dans la console.

## Suppression du contenu distant

La fonction suivante est présente dans le script :

```python
delete_folder_contents(access_token, site_id, folder_path)
```

Elle supprime les éléments directement présents dans le dossier indiqué. Malgré le commentaire du code, l’implémentation actuelle ne parcourt pas elle-même chaque sous-dossier avant suppression. Microsoft Graph peut supprimer un dossier avec son contenu lorsque l’élément dossier est supprimé, mais il est recommandé de tester ce comportement dans un environnement non productif avant utilisation.

La fonction n’est pas appelée dans le programme principal. La synchronisation actuelle ajoute ou remplace les fichiers portant le même nom, mais ne supprime pas automatiquement les fichiers distants absents du répertoire local.

## Limites importantes

### Taille des fichiers

Le script utilise un envoi simple avec :

```http
PUT /content
```

Pour des fichiers volumineux, une session d’upload Microsoft Graph est préférable afin d’envoyer le fichier en plusieurs fragments et de reprendre un transfert interrompu.

### Pagination

La fonction `list_drive_ids()` lit uniquement la première page retournée par Microsoft Graph. Si la réponse contient `@odata.nextLink`, il faut suivre cette URL pour obtenir les pages suivantes.

La même remarque s’applique à la liste des éléments d’un dossier si cette fonction est étendue à de gros volumes.

### Encodage des noms

Les chemins de dossiers sont encodés avec `urllib.parse.quote`. Pour les noms de fichiers contenant des espaces, `#`, `%` ou d’autres caractères réservés, il est plus robuste d’encoder également chaque nom de fichier avant de construire l’URL d’upload.

### Écrasement et suppression

- Un fichier portant le même chemin peut être remplacé lors du téléversement.
- Les fichiers présents dans SharePoint mais absents localement ne sont pas supprimés.
- Il ne s’agit donc pas d’une synchronisation miroir complète.

## Erreur fréquente dans un copier-coller HTML

Dans le code Python, vérifier que la comparaison suivante contient bien le caractère `>` :

```python
if len(path_parts) > 1:
```

Si le fichier contient littéralement `&gt;`, Python générera une erreur de syntaxe. Cette séquence provient généralement d’un copier-coller depuis une page HTML.

## Dépannage

### Erreur 401

Vérifier :

- le `CLIENT_ID` ;
- le `CLIENT_SECRET` ;
- le `TENANT_ID` ;
- l’expiration du secret ;
- l’obtention effective du jeton.

### Erreur 403

Vérifier :

- les permissions Microsoft Graph de l’application ;
- le consentement administrateur ;
- l’autorisation de l’application sur le site SharePoint ciblé si l’accès est restreint.

### Erreur 404 sur le site

Vérifier la construction de `SHAREPOINT_SITE_URL`, notamment :

```text
https://graph.microsoft.com/v1.0/sites/{hostname}:/sites/{site-name}
```

### Erreur 404 sur un dossier

Vérifier :

- que le dossier appartient bien à la bibliothèque par défaut ;
- la valeur de `SHAREPOINT_BASE_FOLDER` ;
- l’encodage du chemin ;
- le `Drive ID` si le dossier se trouve dans une autre bibliothèque.

### Aucun Drive ID affiché

Vérifier que l’application peut appeler :

```text
GET /sites/{site_id}/drives
```

et qu’au moins une bibliothèque de documents est accessible sur le site.

## Recommandations de production

Avant une utilisation en production :

- déplacer les secrets vers des variables d’environnement ou un coffre de secrets ;
- ajouter un délai maximal aux requêtes `requests` ;
- utiliser `response.raise_for_status()` ou une gestion homogène des erreurs ;
- ajouter des tentatives automatiques pour les erreurs temporaires et les réponses `429` ;
- implémenter les sessions d’upload pour les gros fichiers ;
- gérer la pagination Microsoft Graph ;
- encoder les noms de fichiers dans les URLs ;
- tester les suppressions dans un site non productif ;
- limiter les permissions de l’application au strict nécessaire.

## Licence

Projet interne. Adapter la section de licence aux règles de votre organisation.
