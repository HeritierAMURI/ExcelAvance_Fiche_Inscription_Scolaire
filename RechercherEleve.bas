Option Explicit

Sub Rechercher()

    Dim ws As Worksheet
    Dim id As String
    Dim found As Range
    Dim photoPath As String
    Dim fullPhotoPath As String
    Dim lastRow As Long

    On Error GoTo GestionErreur

    '=========================================================
    ' FEUILLE
    '=========================================================
    Set ws = ThisWorkbook.Sheets("Dashboard")


    '=========================================================
    ' RÉCUPÉRER L'ID SAISI
    '=========================================================
    id = Trim(CStr(ws.Range("D8").Value))

    If id = "" Then

        MsgBox "Veuillez saisir l'identifiant de l'élève dans le champ ID.", _
               vbExclamation, "Recherche"

        ws.Range("D8").Select

        Exit Sub

    End If


    '=========================================================
    ' DÉTERMINER LA DERNIÈRE LIGNE DU TABLEAU
    '=========================================================
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    If lastRow < 23 Then

        MsgBox "Aucun élève n'est encore enregistré dans le tableau.", _
               vbInformation, "Recherche"

        Exit Sub

    End If


    '=========================================================
    ' RECHERCHER L'ID DANS LA COLONNE B
    '=========================================================
    Set found = ws.Range("B23:B" & lastRow).Find( _
                    What:=id, _
                    LookIn:=xlValues, _
                    LookAt:=xlWhole, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False)


    '=========================================================
    ' ÉLÈVE INTROUVABLE
    '=========================================================
    If found Is Nothing Then

        MsgBox "Aucun élève ne correspond à l'identifiant :" & _
               vbCrLf & vbCrLf & _
               id, _
               vbInformation, "Élève introuvable"


        ' Nettoyer les champs du formulaire
        ws.Range("D10,D12,D14,D16,D18," & _
                 "G8,G10,G12,G14,G16,G18," & _
                 "J8,J10,J12,J14,J16,J18").ClearContents


        ' Vider la zone photo
        ws.Shapes("PhotoShape").Fill.Solid


        Exit Sub

    End If


    '=========================================================
    ' CHARGER LES INFORMATIONS DE L'ÉLÈVE
    '=========================================================

    '---------------------------------------
    ' Informations D
    '---------------------------------------

    ' B -> D8 : ID
    ws.Range("D8").Value = found.Offset(0, 0).Value

    ' C -> D10
    ws.Range("D10").Value = found.Offset(0, 1).Value

    ' D -> D12
    ws.Range("D12").Value = found.Offset(0, 2).Value

    ' E -> D14
    ws.Range("D14").Value = found.Offset(0, 3).Value

    ' F -> D16
    ws.Range("D16").Value = found.Offset(0, 4).Value

    ' G -> D18
    ws.Range("D18").Value = found.Offset(0, 5).Value


    '---------------------------------------
    ' Informations G
    '---------------------------------------

    ' H -> G8
    ws.Range("G8").Value = found.Offset(0, 6).Value

    ' I -> G10
    ws.Range("G10").Value = found.Offset(0, 7).Value

    ' J -> G12
    ws.Range("G12").Value = found.Offset(0, 8).Value

    ' K -> G14
    ws.Range("G14").Value = found.Offset(0, 9).Value

    ' L -> G16
    ws.Range("G16").Value = found.Offset(0, 10).Value

    ' M -> G18
    ws.Range("G18").Value = found.Offset(0, 11).Value


    '---------------------------------------
    ' Informations J
    '---------------------------------------

    ' N -> J8
    ws.Range("J8").Value = found.Offset(0, 12).Value

    ' O -> J10
    ws.Range("J10").Value = found.Offset(0, 13).Value

    ' P -> J12
    ws.Range("J12").Value = found.Offset(0, 14).Value

    ' Q -> J14
    ws.Range("J14").Value = found.Offset(0, 15).Value

    ' R -> J16
    ws.Range("J16").Value = found.Offset(0, 16).Value

    ' S -> J18
    ws.Range("J18").Value = found.Offset(0, 17).Value


    '=========================================================
    ' RÉCUPÉRER LA PHOTO
    ' COLONNE T = 20
    '=========================================================

    ' IMPORTANT :
    ' On récupère le chemin enregistré dans la colonne T.
    photoPath = Trim(CStr(found.Offset(0, 18).Value))


    '---------------------------------------
    ' Réinitialiser PhotoShape
    '---------------------------------------

    ws.Shapes("PhotoShape").Fill.Solid


    '---------------------------------------
    ' Vérifier si une photo est enregistrée
    '---------------------------------------

    If photoPath <> "" Then

        '-----------------------------------------------------
        ' CAS 1 : chemin absolu
        ' Exemple :
        ' C:\Users\...\Photos\ELV001_20260811.jpg
        '-----------------------------------------------------

        If InStr(1, photoPath, ":\", vbTextCompare) > 0 _
           Or Left(photoPath, 2) = "\\" Then

            fullPhotoPath = photoPath


        '-----------------------------------------------------
        ' CAS 2 : chemin relatif
        ' Exemple :
        ' Photos\ELV001_20260811.jpg
        '-----------------------------------------------------

        Else

            fullPhotoPath = ThisWorkbook.Path & _
                            Application.PathSeparator & _
                            photoPath

        End If


        '-----------------------------------------------------
        ' Vérifier que le fichier existe
        '-----------------------------------------------------

        If Dir(fullPhotoPath) <> "" Then

            '-------------------------------------------------
            ' Afficher la photo dans PhotoShape
            '-------------------------------------------------

            With ws.Shapes("PhotoShape")

                .Fill.Visible = msoTrue
                .Fill.UserPicture fullPhotoPath

            End With


        Else

            MsgBox "L'élève a bien été trouvé, mais sa photo est introuvable." & _
                   vbCrLf & vbCrLf & _
                   "Chemin enregistré :" & vbCrLf & _
                   photoPath & vbCrLf & vbCrLf & _
                   "Chemin recherché :" & vbCrLf & _
                   fullPhotoPath, _
                   vbExclamation, "Photo introuvable"

        End If

    End If


    '=========================================================
    ' MESSAGE DE CONFIRMATION
    '=========================================================

    MsgBox "Élève trouvé avec succès !" & _
           vbCrLf & vbCrLf & _
           "ID : " & id, _
           vbInformation, "Recherche"


    Exit Sub


'=============================================================
' GESTION DES ERREURS
'=============================================================

GestionErreur:

    MsgBox "Une erreur est survenue lors de la recherche." & _
           vbCrLf & vbCrLf & _
           "Numéro d'erreur : " & Err.Number & _
           vbCrLf & _
           "Description : " & Err.Description, _
           vbCritical, "Erreur"

End Sub

