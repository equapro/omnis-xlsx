# omnis-xlsx
Omnis Studio JavaScript worker module for [SheetJS](https://sheetjs.com).

## Usage

Créer un objet (oXlsx par exemple) ayant comme superclass `.OW3.OW3 Worker Objects\JAVASCRIPTWorker`.

### Ecriture

```
# oXlslx.$write

# parameters
#   - pData         List       Données à écrire
#   - pHeader       List       En-têtes des colonnes (optionnel)
#   - pPath         String     Chemin complet vers le fichier
#   - pSheetName    String     Nom de la feuille (optionnel - Défaut : "Feuil1")
# variables
#   - lParam        Row        Paramètres envoyés au Js Worker

Do lParam.$cols.$add('filename',kCharacter,kSimplechar)
Do lParam.$cols.$add('sheetName',kCharacter,kSimplechar)
Do lParam.$cols.$add('data',kList)
Calculate lParam.filename as pPath
Calculate lParam.sheetName as pSheetName
Calculate lParam.data as pData
If pHeader.$linecount
  Do lParam.$cols.$add('rowHeader',kList)
  Calculate lParam.rowHeader as pHeader
End If

Do $cinst.$callmethod('xlsx','write',lParam,kTrue) Returns #F
```


## nmp registry
https://www.npmjs.com/package/@equapro/omnis-xlsx

Une github action publie automatiquement le package sur npmjs.com.

L'action est basée sur le push d'un tag `vX.Y.Z`
