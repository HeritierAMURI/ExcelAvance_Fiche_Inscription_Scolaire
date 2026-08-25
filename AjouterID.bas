Option Explicit

Sub AjouterID()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim valeurID As Variant
    Dim maxID As Long
    Dim newID As Long

    On Error GoTo ErrorHandler

    '=========================================================
    ' FEUILLE
    '=========================================================
    Set ws = ThisWorkbook.Sheets("Dashboard")


    '=========================================================
    ' DÉTERMINER LA DERNIÈRE LIGNE DU TABLEAU
    ' Les élèves commencent à la ligne 23
    '=========================================================
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row


    '=========================================================
    ' INITIALISER LE PLUS GRAND ID
    '=========================================================
    maxID = 0


    '=========================================================
    ' PARCOURIR LES IDs EXISTANTS
    '=========================================================
    If lastRow >= 23 Then

        For i = 23 To lastRow

            valeurID = ws.Cells(i, "B").Value

            ' Vérifier si la cellule contient un nombre
            If IsNumeric(valeurID) And Trim(CStr(valeurID)) <> "" Then

                ' Garder uniquement le plus grand ID
                If CLng(valeurID) > maxID Then
                    maxID = CLng(valeurID)
                End If

            End If

        Next i

    End If


    '=========================================================
    ' GÉNÉRER LE PROCHAIN ID
    '=========================================================
    newID = maxID + 1


    '=========================================================
    ' PLACER LE NOUVEL ID DANS D8
    '=========================================================
    ws.Range("D8").Value = newID


    '=========================================================
    ' FORMAT DE LA CELLULE ID
    '=========================================================
    ws.Range("D8").NumberFormat = "0"


    Exit Sub


'=============================================================
' GESTION DES ERREURS
'=============================================================

ErrorHandler:

    MsgBox "Une erreur s'est produite lors de la génération de l'identifiant." & _
           vbCrLf & vbCrLf & _
           "Erreur n° : " & Err.Number & _
           vbCrLf & _
           "Description : " & Err.Description, _
           vbCritical, "Erreur"

End Sub

