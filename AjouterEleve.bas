Option Explicit

Public photoPath As String

Sub Ajouter()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim id As String
    Dim fs As Object
    Dim folderPath As String
    Dim fileName As String
    Dim newPhotoPath As String
    Dim found As Range

    Set ws = ThisWorkbook.Sheets("Dashboard")

    '---------------------------------------
    ' Récupération de l'ID
    '---------------------------------------
    id = Trim(ws.Range("D8").Value)

    If id = "" Then
        MsgBox "Veuillez saisir l'ID de l'élève.", vbExclamation
        Exit Sub
    End If

    '---------------------------------------
    ' Vérification de l'unicité de l'ID
    '---------------------------------------
    Set found = ws.Range("B23:B" & ws.Rows.Count).Find( _
                    What:=id, _
                    LookIn:=xlValues, _
                    LookAt:=xlWhole)

    If Not found Is Nothing Then
        MsgBox "Cet ID existe déjà. Veuillez utiliser un ID unique ou cliquez sur le bouton Modifier.", vbExclamation
        Exit Sub
    End If

    '---------------------------------------
    ' Vérification que le classeur est enregistré
    '---------------------------------------
    If ThisWorkbook.Path = "" Then
        MsgBox "Veuillez d'abord enregistrer le classeur avant d'ajouter un élève.", vbExclamation
        Exit Sub
    End If

    '---------------------------------------
    ' Détermination de la prochaine ligne
    '---------------------------------------
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row

    If lastRow < 22 Then
        lastRow = 22
    End If

    lastRow = lastRow + 1

    '---------------------------------------
    ' Gestion de la photo
    '---------------------------------------
    newPhotoPath = ""

    If photoPath <> "" Then

        Set fs = CreateObject("Scripting.FileSystemObject")

        ' Dossier Photos situé à côté du classeur
        folderPath = ThisWorkbook.Path & Application.PathSeparator & "Photos"

        ' Création du dossier s'il n'existe pas
        If Not fs.FolderExists(folderPath) Then
            fs.CreateFolder folderPath
        End If

        ' Nom de la photo
        fileName = id & "_" & Format(Now, "yyyymmdd_hhnnss") _
                   & "." & fs.GetExtensionName(photoPath)

        newPhotoPath = folderPath & Application.PathSeparator & fileName

        ' Copie de la photo
        fs.CopyFile photoPath, newPhotoPath, True

    End If

    '---------------------------------------
    ' Enregistrement des informations
    '---------------------------------------

    ws.Cells(lastRow, 2).Value = ws.Range("D8").Value
    ws.Cells(lastRow, 3).Value = ws.Range("D10").Value
    ws.Cells(lastRow, 4).Value = ws.Range("D12").Value
    ws.Cells(lastRow, 5).Value = ws.Range("D14").Value
    ws.Cells(lastRow, 6).Value = ws.Range("D16").Value
    ws.Cells(lastRow, 7).Value = ws.Range("D18").Value

    ws.Cells(lastRow, 8).Value = ws.Range("G8").Value
    ws.Cells(lastRow, 9).Value = ws.Range("G10").Value
    ws.Cells(lastRow, 10).Value = ws.Range("G12").Value
    ws.Cells(lastRow, 11).Value = ws.Range("G14").Value
    ws.Cells(lastRow, 12).Value = ws.Range("G16").Value
    ws.Cells(lastRow, 13).Value = ws.Range("G18").Value

    ws.Cells(lastRow, 14).Value = ws.Range("J8").Value
    ws.Cells(lastRow, 15).Value = ws.Range("J10").Value
    ws.Cells(lastRow, 16).Value = ws.Range("J12").Value
    ws.Cells(lastRow, 17).Value = ws.Range("J14").Value
    ws.Cells(lastRow, 18).Value = ws.Range("J16").Value
    ws.Cells(lastRow, 19).Value = ws.Range("J18").Value

    '---------------------------------------
    ' Enregistrement du chemin de la photo
    ' Colonne T = 20
    '---------------------------------------
    ws.Cells(lastRow, 20).Value = newPhotoPath

    '---------------------------------------
    ' NE PAS écrire le numéro de ligne
    '---------------------------------------
    ' ws.Cells(lastRow, 21).Value = lastRow

    '---------------------------------------
    ' Message de confirmation
    '---------------------------------------
    MsgBox "Données enregistrées avec succès !", vbInformation

    '---------------------------------------
    ' Nettoyage du formulaire
    '---------------------------------------
    ws.Range("D8,D10,D12,D14,D16,D18," & _
             "G8,G10,G12,G14,G16,G18," & _
             "J8,J10,J12,J14,J16,J18").ClearContents

    ' Réinitialisation de la photo
    photoPath = ""

    ws.Shapes("PhotoShape").Fill.Solid

End Sub


Sub AjouterPhoto()

    Dim picPath As Variant
    Dim shp As Shape

    picPath = Application.GetOpenFilename( _
        "Images (*.jpg;*.jpeg;*.png),*.jpg;*.jpeg;*.png", _
        , _
        "Sélectionner une photo")

    If picPath <> False Then

        photoPath = CStr(picPath)

        Set shp = ThisWorkbook.Sheets("Dashboard").Shapes("PhotoShape")

        shp.Fill.UserPicture photoPath

    End If

End Sub


Sub ViderImage()

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets("Dashboard")

    photoPath = ""

    ws.Shapes("PhotoShape").Fill.Solid

End Sub



