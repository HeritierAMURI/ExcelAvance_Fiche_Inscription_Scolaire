Option Explicit

Sub Modifier()

    Dim ws As Worksheet
    Dim found As Range

    Dim id As String
    Dim oldPhotoPath As String
    Dim newPhotoPath As String

    Dim fs As Object
    Dim folderPath As String
    Dim fileName As String

    On Error GoTo GestionErreur

    '=========================================================
    ' FEUILLE
    '=========================================================
    Set ws = ThisWorkbook.Sheets("Dashboard")


    '=========================================================
    ' RÉCUPÉRER L'ID
    '=========================================================
    id = Trim(CStr(ws.Range("D8").Value))

    If id = "" Then

        MsgBox "Veuillez saisir l'identifiant de l'élève à modifier.", _
               vbExclamation, "Modification"

        ws.Range("D8").Select

        Exit Sub

    End If


    '=========================================================
    ' RECHERCHER L'ÉLÈVE DANS LE TABLEAU
    ' Colonne B = ID
    '=========================================================
    Set found = ws.Range("B23:B" & _
                         ws.Cells(ws.Rows.Count, "B").End(xlUp).Row).Find( _
                            What:=id, _
                            LookIn:=xlValues, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False)


    '=========================================================
    ' VÉRIFIER SI L'ÉLÈVE EXISTE
    '=========================================================
    If found Is Nothing Then

        MsgBox "Impossible de modifier cet élève." & _
               vbCrLf & vbCrLf & _
               "Aucun élève ne correspond à l'ID : " & id, _
               vbExclamation, "Élève introuvable"

        Exit Sub

    End If


    '=========================================================
    ' CONFIRMATION AVANT MODIFICATION
    '=========================================================
    If MsgBox("Voulez-vous vraiment enregistrer les modifications " & _
              "apportées à l'élève :" & vbCrLf & vbCrLf & _
              "ID : " & id & vbCrLf & vbCrLf & _
              "Cette opération va remplacer les anciennes informations.", _
              vbQuestion + vbYesNo, _
              "Confirmation de modification") = vbNo Then

        Exit Sub

    End If


    '=========================================================
    ' MODIFICATION DES INFORMATIONS
    '=========================================================

    '---------------------------------------
    ' Colonne C à G
    '---------------------------------------

    ' D10 -> C
    ws.Cells(found.Row, 3).Value = ws.Range("D10").Value

    ' D12 -> D
    ws.Cells(found.Row, 4).Value = ws.Range("D12").Value

    ' D14 -> E
    ws.Cells(found.Row, 5).Value = ws.Range("D14").Value

    ' D16 -> F
    ws.Cells(found.Row, 6).Value = ws.Range("D16").Value

    ' D18 -> G
    ws.Cells(found.Row, 7).Value = ws.Range("D18").Value


    '---------------------------------------
    ' Colonne H à M
    '---------------------------------------

    ' G8 -> H
    ws.Cells(found.Row, 8).Value = ws.Range("G8").Value

    ' G10 -> I
    ws.Cells(found.Row, 9).Value = ws.Range("G10").Value

    ' G12 -> J
    ws.Cells(found.Row, 10).Value = ws.Range("G12").Value

    ' G14 -> K
    ws.Cells(found.Row, 11).Value = ws.Range("G14").Value

    ' G16 -> L
    ws.Cells(found.Row, 12).Value = ws.Range("G16").Value

    ' G18 -> M
    ws.Cells(found.Row, 13).Value = ws.Range("G18").Value


    '---------------------------------------
    ' Colonne N à S
    '---------------------------------------

    ' J8 -> N
    ws.Cells(found.Row, 14).Value = ws.Range("J8").Value

    ' J10 -> O
    ws.Cells(found.Row, 15).Value = ws.Range("J10").Value

    ' J12 -> P
    ws.Cells(found.Row, 16).Value = ws.Range("J12").Value

    ' J14 -> Q
    ws.Cells(found.Row, 17).Value = ws.Range("J14").Value

    ' J16 -> R
    ws.Cells(found.Row, 18).Value = ws.Range("J16").Value

    ' J18 -> S
    ws.Cells(found.Row, 19).Value = ws.Range("J18").Value


    '=========================================================
    ' GESTION DE LA PHOTO
    ' Colonne T = 20
    '=========================================================

    ' Ancien chemin de la photo
    oldPhotoPath = Trim(CStr(ws.Cells(found.Row, 20).Value))


    '---------------------------------------------------------
    ' Si l'utilisateur a sélectionné une nouvelle photo
    ' photopath contient alors son emplacement original
    '---------------------------------------------------------
    If photoPath <> "" Then

        ' Vérifier que le classeur est enregistré
        If ThisWorkbook.Path = "" Then

            MsgBox "Veuillez d'abord enregistrer le classeur avant " & _
                   "de modifier la photo.", _
                   vbExclamation, "Classeur non enregistré"

            Exit Sub

        End If


        '-----------------------------------------------------
        ' Création du dossier Photos
        '-----------------------------------------------------
        Set fs = CreateObject("Scripting.FileSystemObject")

        folderPath = ThisWorkbook.Path & _
                     Application.PathSeparator & _
                     "Photos"


        If Not fs.FolderExists(folderPath) Then

            fs.CreateFolder folderPath

        End If


        '-----------------------------------------------------
        ' Nom de la nouvelle photo
        '-----------------------------------------------------
        fileName = id & "_" & _
                   Format(Now, "yyyymmdd_hhnnss") & "." & _
                   fs.GetExtensionName(photoPath)


        newPhotoPath = folderPath & _
                       Application.PathSeparator & _
                       fileName


        '-----------------------------------------------------
        ' Copier la nouvelle photo
        '-----------------------------------------------------
        fs.CopyFile photoPath, newPhotoPath, True


        '-----------------------------------------------------
        ' Enregistrer le nouveau chemin dans la colonne T
        '-----------------------------------------------------
        ws.Cells(found.Row, 20).Value = newPhotoPath


        '-----------------------------------------------------
        ' Afficher la nouvelle photo
        '-----------------------------------------------------
        ws.Shapes("PhotoShape").Fill.UserPicture newPhotoPath


        '-----------------------------------------------------
        ' Mettre à jour la variable photopath
        '-----------------------------------------------------
        photoPath = newPhotoPath


    Else

        '-----------------------------------------------------
        ' Aucune nouvelle photo sélectionnée :
        ' conserver automatiquement l'ancienne photo
        '-----------------------------------------------------
        ws.Cells(found.Row, 20).Value = oldPhotoPath

    End If


    '=========================================================
    ' ID
    '=========================================================
    ' L'ID n'est volontairement PAS modifié.
    '
    ' L'ID de l'élève reste celui qui a servi à la recherche.
    '=========================================================


    '=========================================================
    ' MESSAGE DE CONFIRMATION
    '=========================================================
    MsgBox "Les informations de l'élève ont été modifiées " & _
           "avec succès !" & vbCrLf & vbCrLf & _
           "ID : " & id, _
           vbInformation, "Modification réussie"


    Exit Sub


'=============================================================
' GESTION DES ERREURS
'=============================================================

GestionErreur:

    MsgBox "Une erreur est survenue lors de la modification." & _
           vbCrLf & vbCrLf & _
           "Numéro d'erreur : " & Err.Number & _
           vbCrLf & _
           "Description : " & Err.Description, _
           vbCritical, "Erreur"

End Sub

