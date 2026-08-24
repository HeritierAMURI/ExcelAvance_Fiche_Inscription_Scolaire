Option Explicit

Sub SupprimerEleve()

    Dim ws As Worksheet
    Dim found As Range

    Dim id As String
    Dim photoPath As String
    Dim fullPhotoPath As String

    Dim reponse As VbMsgBoxResult

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

        MsgBox "Veuillez d'abord rechercher l'élève à supprimer." & _
               vbCrLf & vbCrLf & _
               "Saisissez son ID dans le champ prévu, puis cliquez " & _
               "sur le bouton Rechercher.", _
               vbExclamation, "Suppression"

        ws.Range("D8").Select

        Exit Sub

    End If


    '=========================================================
    ' RECHERCHER L'ÉLÈVE DANS LA COLONNE B
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

        MsgBox "Aucun élève ne correspond à l'ID :" & _
               vbCrLf & vbCrLf & _
               id, _
               vbExclamation, "Élève introuvable"

        Exit Sub

    End If


    '=========================================================
    ' RÉCUPÉRER LE CHEMIN DE LA PHOTO
    ' Colonne T = 20
    '=========================================================
    photoPath = Trim(CStr(ws.Cells(found.Row, 20).Value))


    '=========================================================
    ' CONFIRMATION DE SUPPRESSION
    '=========================================================
    reponse = MsgBox( _
        "ATTENTION !" & vbCrLf & vbCrLf & _
        "Vous êtes sur le point de supprimer définitivement " & _
        "l'élève ayant l'ID :" & vbCrLf & vbCrLf & _
        "        " & id & vbCrLf & vbCrLf & _
        "Toutes les informations de cet élève seront supprimées." & _
        vbCrLf & _
        "Sa photo sera également supprimée du dossier Photos." & _
        vbCrLf & vbCrLf & _
        "Voulez-vous continuer ?", _
        vbQuestion + vbYesNo + vbDefaultButton2, _
        "Confirmation de suppression")


    If reponse = vbNo Then

        MsgBox "Suppression annulée.", _
               vbInformation, "Suppression"

        Exit Sub

    End If


    '=========================================================
    ' SUPPRESSION DE LA PHOTO
    '=========================================================
    If photoPath <> "" Then

        '-----------------------------------------------------
        ' Déterminer si le chemin est absolu ou relatif
        '-----------------------------------------------------
        If InStr(1, photoPath, ":\", vbTextCompare) > 0 _
           Or Left(photoPath, 2) = "\\" Then

            ' Chemin absolu
            fullPhotoPath = photoPath

        Else

            ' Chemin relatif
            fullPhotoPath = ThisWorkbook.Path & _
                            Application.PathSeparator & _
                            photoPath

        End If


        '-----------------------------------------------------
        ' Vérifier que la photo existe avant suppression
        '-----------------------------------------------------
        If Dir(fullPhotoPath) <> "" Then

            Kill fullPhotoPath

        End If

    End If


    '=========================================================
    ' SUPPRIMER LA LIGNE DE L'ÉLÈVE
    '=========================================================
    ws.Rows(found.Row).Delete


    '=========================================================
    ' NETTOYER LE FORMULAIRE
    '=========================================================

    ws.Range("D8,D10,D12,D14,D16,D18," & _
             "G8,G10,G12,G14,G16,G18," & _
             "J8,J10,J12,J14,J16,J18").ClearContents


    '=========================================================
    ' VIDER LA ZONE PHOTO
    '=========================================================

    ws.Shapes("PhotoShape").Fill.Solid


    ' Réinitialiser la variable photo
    photoPath = ""


    '=========================================================
    ' MESSAGE DE CONFIRMATION
    '=========================================================

    MsgBox "L'élève a été supprimé avec succès." & _
           vbCrLf & vbCrLf & _
           "ID supprimé : " & id, _
           vbInformation, "Suppression réussie"


    Exit Sub


'=============================================================
' GESTION DES ERREURS
'=============================================================

GestionErreur:

    MsgBox "Une erreur est survenue lors de la suppression." & _
           vbCrLf & vbCrLf & _
           "Numéro d'erreur : " & Err.Number & _
           vbCrLf & _
           "Description : " & Err.Description, _
           vbCritical, "Erreur"

End Sub

