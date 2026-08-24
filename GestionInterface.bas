Option Explicit

'=============================================================
' MODULE : INTERFACE
'
' Contient :
'   1. ViderChamp
'   2. PleinEcrant
'   3. QuiterPleinEcran
'
'=============================================================


'=============================================================
' 1. VIDER LES CHAMPS
'=============================================================

Sub ViderChamp()

    Dim ws As Worksheet

    On Error GoTo GestionErreur

    '---------------------------------------------------------
    ' Feuille Dashboard
    '---------------------------------------------------------
    Set ws = ThisWorkbook.Sheets("Dashboard")


    '---------------------------------------------------------
    ' Vider tous les champs du formulaire
    '---------------------------------------------------------

    ws.Range("D8,D10,D12,D14,D16,D18," & _
             "G8,G10,G12,G14,G16,G18," & _
             "J8,J10,J12,J14,J16,J18").ClearContents


    '---------------------------------------------------------
    ' Vider la photo
    '---------------------------------------------------------

    ws.Shapes("PhotoShape").Fill.Solid


    '---------------------------------------------------------
    ' Réinitialiser le chemin de la photo
    '---------------------------------------------------------

    photoPath = ""


    '---------------------------------------------------------
    ' Replacer le curseur sur le champ ID
    '---------------------------------------------------------

    ws.Range("D8").Select


    Exit Sub


'=============================================================
' GESTION DES ERREURS
'=============================================================

GestionErreur:

    MsgBox "Une erreur est survenue lors de la suppression des champs." & _
           vbCrLf & vbCrLf & _
           "Erreur n° : " & Err.Number & _
           vbCrLf & _
           "Description : " & Err.Description, _
           vbCritical, "Erreur"

End Sub



'=============================================================
' 2. PASSER EN MODE PLEIN ÉCRAN
'=============================================================

Sub PleinEcrant()

    On Error GoTo GestionErreur

    '---------------------------------------------------------
    ' Cacher le ruban Excel
    '---------------------------------------------------------

    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"


    '---------------------------------------------------------
    ' Cacher la barre de formule
    '---------------------------------------------------------

    Application.DisplayFormulaBar = False


    '---------------------------------------------------------
    ' Cacher les en-têtes de lignes et de colonnes
    '---------------------------------------------------------

    ActiveWindow.DisplayHeadings = False


    '---------------------------------------------------------
    ' Cacher les onglets des feuilles
    '---------------------------------------------------------

    ActiveWindow.DisplayWorkbookTabs = False


    '---------------------------------------------------------
    ' Cacher la barre de défilement horizontale et verticale
    '---------------------------------------------------------

    ActiveWindow.DisplayHorizontalScrollBar = False
    ActiveWindow.DisplayVerticalScrollBar = False


    '---------------------------------------------------------
    ' Masquer la barre d'état
    '---------------------------------------------------------

    Application.DisplayStatusBar = False


    '---------------------------------------------------------
    ' Masquer les boutons de l'interface Excel
    '---------------------------------------------------------

    Application.DisplayFullScreen = True


    Exit Sub


'=============================================================
' GESTION DES ERREURS
'=============================================================

GestionErreur:

    MsgBox "Une erreur est survenue lors du passage en plein écran." & _
           vbCrLf & vbCrLf & _
           "Erreur n° : " & Err.Number & _
           vbCrLf & _
           "Description : " & Err.Description, _
           vbCritical, "Erreur"

End Sub



'=============================================================
' 3. QUITTER LE MODE PLEIN ÉCRAN
'=============================================================

Sub QuiterPleinEcran()

    On Error GoTo GestionErreur

    '---------------------------------------------------------
    ' Quitter le mode plein écran
    '---------------------------------------------------------

    Application.DisplayFullScreen = False


    '---------------------------------------------------------
    ' Afficher le ruban Excel
    '---------------------------------------------------------

    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"


    '---------------------------------------------------------
    ' Afficher la barre de formule
    '---------------------------------------------------------

    Application.DisplayFormulaBar = True


    '---------------------------------------------------------
    ' Afficher les en-têtes de lignes et colonnes
    '---------------------------------------------------------

    ActiveWindow.DisplayHeadings = True


    '---------------------------------------------------------
    ' Afficher les onglets des feuilles
    '---------------------------------------------------------

    ActiveWindow.DisplayWorkbookTabs = True


    '---------------------------------------------------------
    ' Afficher les barres de défilement
    '---------------------------------------------------------

    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayVerticalScrollBar = True


    '---------------------------------------------------------
    ' Afficher la barre d'état
    '---------------------------------------------------------

    Application.DisplayStatusBar = True


    Exit Sub


'=============================================================
' GESTION DES ERREURS
'=============================================================

GestionErreur:

    MsgBox "Une erreur est survenue lors de la sortie du plein écran." & _
           vbCrLf & vbCrLf & _
           "Erreur n° : " & Err.Number & _
           vbCrLf & _
           "Description : " & Err.Description, _
           vbCritical, "Erreur"

End Sub

