Option Explicit

'=============================================================
' MODULE : IMPRESSION
'
' Fonction :
'   - Récupérer les informations du formulaire Dashboard
'   - Les transférer vers PrintSheet
'   - Afficher la photo dans PrintSheet
'   - Imprimer uniquement la feuille PrintSheet
'
'=============================================================

Sub Imprimer()

    Dim wsDashboard As Worksheet
    Dim wsPrint As Worksheet

    Dim id As String
    Dim photoPath As String

    On Error GoTo GestionErreur

    '=========================================================
    ' DÉFINIR LES FEUILLES
    '=========================================================

    Set wsDashboard = ThisWorkbook.Sheets("Dashboard")
    Set wsPrint = ThisWorkbook.Sheets("PrintSheet")


    '=========================================================
    ' VÉRIFIER QU'UN ÉLÈVE EST CHARGÉ
    '=========================================================

    id = Trim(CStr(wsDashboard.Range("D8").Value))

    If id = "" Then

        MsgBox "Aucun élève n'est chargé dans le formulaire." & _
               vbCrLf & vbCrLf & _
               "Veuillez d'abord rechercher ou enregistrer un élève.", _
               vbExclamation, "Impression"

        Exit Sub

    End If


    '=========================================================
    ' TRANSFERT DES INFORMATIONS VERS PRINTSHEET
    '=========================================================

    '---------------------------------------------------------
    ' Informations générales
    '---------------------------------------------------------

    ' ID
    wsPrint.Range("C5").Value = wsDashboard.Range("D8").Value

    ' Nom
    wsPrint.Range("C6").Value = wsDashboard.Range("D10").Value

    ' Post-nom
    wsPrint.Range("C7").Value = wsDashboard.Range("D12").Value

    ' Prénom
    wsPrint.Range("C8").Value = wsDashboard.Range("D14").Value

    ' Statut
    wsPrint.Range("C9").Value = wsDashboard.Range("J10").Value

    ' Classe
    wsPrint.Range("C10").Value = wsDashboard.Range("J12").Value

    ' Option
    wsPrint.Range("C11").Value = wsDashboard.Range("J14").Value

    ' Dossier
    wsPrint.Range("C12").Value = wsDashboard.Range("J16").Value

    ' Lieu et date de naissance
    wsPrint.Range("C13").Value = wsDashboard.Range("D16").Value


    '=========================================================
    ' INFORMATIONS COMPLÉMENTAIRES
    '=========================================================

    ' École de provenance
    wsPrint.Range("C15").Value = wsDashboard.Range("J8").Value

    ' Adresse
    wsPrint.Range("C16").Value = wsDashboard.Range("D18").Value

    ' Nom du père
    wsPrint.Range("C17").Value = wsDashboard.Range("G8").Value

    ' Nom de la mère
    wsPrint.Range("C18").Value = wsDashboard.Range("G10").Value

    ' Responsable / Tuteur
    wsPrint.Range("C19").Value = wsDashboard.Range("G12").Value

    ' Fonction du responsable
    wsPrint.Range("C20").Value = wsDashboard.Range("G14").Value

    ' Contact du responsable
    wsPrint.Range("C21").Value = wsDashboard.Range("G16").Value

    ' Origine
    wsPrint.Range("C22").Value = wsDashboard.Range("G18").Value

    ' Observation
    wsPrint.Range("C23").Value = wsDashboard.Range("J18").Value


    '=========================================================
    ' GESTION DE LA PHOTO
    '=========================================================

    ' Récupérer le chemin de la photo actuellement chargée
    photoPath = Trim(CStr(photoPath))


    '---------------------------------------------------------
    ' Vider d'abord la zone photo de PrintSheet
    '---------------------------------------------------------

    wsPrint.Shapes("PhotoShape").Fill.Solid


    '---------------------------------------------------------
    ' Si une photo existe
    '---------------------------------------------------------

    If photoPath <> "" Then

        ' Vérifier que le fichier existe réellement
        If Dir(photoPath) <> "" Then

            ' Afficher la photo dans PrintSheet
            With wsPrint.Shapes("PhotoShape")

                .Fill.Visible = msoTrue
                .Fill.UserPicture photoPath

            End With

        Else

            MsgBox "La photo de l'élève est introuvable." & _
                   vbCrLf & vbCrLf & _
                   "Chemin :" & vbCrLf & _
                   photoPath, _
                   vbExclamation, "Photo introuvable"

        End If

    End If


    '=========================================================
    ' CONFIGURATION DE L'IMPRESSION
    '=========================================================

    With wsPrint.PageSetup

        ' Orientation portrait
        .Orientation = xlPortrait

        ' Ajuster à une seule page en largeur
        .FitToPagesWide = 1

        ' Une seule page en hauteur
        .FitToPagesTall = 1

        ' Désactiver l'échelle manuelle
        .Zoom = False

    End With


    '=========================================================
    ' CONFIGURER LA ZONE D'IMPRESSION
    '=========================================================

    wsPrint.PageSetup.PrintArea = "$A$1:$E$23"


    '=========================================================
    ' ACTIVER LA FEUILLE D'IMPRESSION
    '=========================================================

    wsPrint.Activate


    '=========================================================
    ' APERÇU AVANT IMPRESSION
    '=========================================================

    wsPrint.PrintPreview


    Exit Sub


'=============================================================
' GESTION DES ERREURS
'=============================================================

GestionErreur:

    MsgBox "Une erreur est survenue lors de la préparation de l'impression." & _
           vbCrLf & vbCrLf & _
           "Numéro d'erreur : " & Err.Number & _
           vbCrLf & _
           "Description : " & Err.Description, _
           vbCritical, "Erreur d'impression"

End Sub

