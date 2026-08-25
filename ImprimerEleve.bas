
Option Explicit

Sub ImprimerEleve()

    Dim wsDashboard As Worksheet
    Dim wsPrint As Worksheet
    Dim photoShape As Shape
    
    Dim photoPath As String
    Dim valeurID As String
    
    On Error GoTo ErrorHandler

    '=========================================================
    ' FEUILLES
    '=========================================================
    Set wsDashboard = ThisWorkbook.Sheets("Dashboard")
    Set wsPrint = ThisWorkbook.Sheets("PrintSheet")

    '=========================================================
    ' RÉCUPÉRER LE CHEMIN DE LA PHOTO
    ' La colonne T contient le chemin de la photo
    '=========================================================
    photoPath = Trim(CStr(wsDashboard.Range("T23").Value))

    '=========================================================
    ' TRANSFERT DES INFORMATIONS VERS PrintSheet
    '=========================================================

    '---------------------------------------------------------
    ' Informations principales
    '---------------------------------------------------------
    wsPrint.Range("C5").Value = wsDashboard.Range("D8").Value
    wsPrint.Range("C6").Value = wsDashboard.Range("D10").Value
    wsPrint.Range("C7").Value = wsDashboard.Range("D12").Value
    wsPrint.Range("C8").Value = wsDashboard.Range("D14").Value

    '---------------------------------------------------------
    ' Informations supplémentaires
    '---------------------------------------------------------
    wsPrint.Range("C9").Value = wsDashboard.Range("J10").Value
    wsPrint.Range("C10").Value = wsDashboard.Range("J12").Value
    wsPrint.Range("C11").Value = wsDashboard.Range("J14").Value
    wsPrint.Range("C12").Value = wsDashboard.Range("J16").Value

    '---------------------------------------------------------
    ' Autres informations
    '---------------------------------------------------------
    wsPrint.Range("C13").Value = wsDashboard.Range("D16").Value

    ' C14 laissé vide
    wsPrint.Range("C14").ClearContents

    '---------------------------------------------------------
    ' Suite des informations
    '---------------------------------------------------------
    wsPrint.Range("C15").Value = wsDashboard.Range("J8").Value
    wsPrint.Range("C16").Value = wsDashboard.Range("D18").Value
    wsPrint.Range("C17").Value = wsDashboard.Range("G8").Value
    wsPrint.Range("C18").Value = wsDashboard.Range("G10").Value
    wsPrint.Range("C19").Value = wsDashboard.Range("G12").Value
    wsPrint.Range("C20").Value = wsDashboard.Range("G14").Value
    wsPrint.Range("C21").Value = wsDashboard.Range("G16").Value
    wsPrint.Range("C22").Value = wsDashboard.Range("G18").Value
    wsPrint.Range("C23").Value = wsDashboard.Range("J18").Value


    '=========================================================
    ' CHARGEMENT DE LA PHOTO
    '=========================================================

    On Error Resume Next
    Set photoShape = wsPrint.Shapes("PhotoShape")
    On Error GoTo ErrorHandler

    '---------------------------------------------------------
    ' Vérifier que PhotoShape existe
    '---------------------------------------------------------
    If photoShape Is Nothing Then
        
        MsgBox "La zone PhotoShape est introuvable dans la feuille PrintSheet." & _
               vbCrLf & vbCrLf & _
               "Vérifiez que la forme contenant la photo porte exactement le nom :" & _
               vbCrLf & _
               "PhotoShape", _
               vbExclamation, "Photo introuvable"
        
    Else
        
        '-----------------------------------------------------
        ' Nettoyer l'ancienne photo
        '-----------------------------------------------------
        On Error Resume Next
        photoShape.Fill.Visible = msoFalse
        photoShape.Line.Visible = msoTrue
        On Error GoTo ErrorHandler
        
        '-----------------------------------------------------
        ' Vérifier le chemin de la photo
        '-----------------------------------------------------
        If photoPath <> "" Then
            
            '-------------------------------------------------
            ' Vérifier si le fichier existe réellement
            '-------------------------------------------------
            If Len(Dir(photoPath)) > 0 Then
                
                '-------------------------------------------------
                ' Insérer la nouvelle photo
                '-------------------------------------------------
                With photoShape
                    .Fill.Visible = msoTrue
                    .Fill.UserPicture photoPath
                    
                    ' Garder les proportions de la photo
                    .Fill.TextureTile = msoFalse
                End With
                
            Else
                
                MsgBox "La photo de l'élève est introuvable." & _
                       vbCrLf & vbCrLf & _
                       "Chemin enregistré :" & _
                       vbCrLf & _
                       photoPath, _
                       vbExclamation, "Photo introuvable"
                
            End If
            
        Else
            
            MsgBox "Aucun chemin de photo n'est enregistré pour cet élève.", _
                   vbInformation, "Photo absente"
        
        End If
        
    End If


    '=========================================================
    ' CONFIGURATION DE L'IMPRESSION
    '=========================================================

    With wsPrint.PageSetup
        
        .PrintArea = "$A$1:$E$23"
        
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        
        ' Centrage horizontal uniquement
        .CenterHorizontally = True
        
        ' Ne pas centrer verticalement :
        ' l'impression commence en haut de la page
        .CenterVertically = False
        
    End With


    '=========================================================
    ' APERÇU AVANT IMPRESSION
    '=========================================================
    wsPrint.PrintPreview

    Exit Sub


'=============================================================
' GESTION DES ERREURS
'=============================================================
ErrorHandler:

    MsgBox "Une erreur s'est produite lors de l'impression." & _
           vbCrLf & vbCrLf & _
           "Erreur n° : " & Err.Number & _
           vbCrLf & _
           "Description : " & Err.Description, _
           vbCritical, "Erreur"

End Sub

