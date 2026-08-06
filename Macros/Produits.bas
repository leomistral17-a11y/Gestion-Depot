REM  *****  BASIC  *****

Option Explicit

Function PrixProduit(ligneProduit As Integer, situation As String) As Double

    Dim col As Integer

    Select Case situation
        Case "p": col = 1
        Case "c": col = 2
        Case "i": col = 3
        Case "n": col = 4
        Case Else: col = 1
    End Select

    PrixProduit = Sheet("Produits").getCellByPosition(col, ligneProduit).Value

End Function

