REM  *****  BASIC  *****

Option Explicit

Function CalculTotalVente(ligne As Long) As Double

    Dim total As Double
    Dim i As Integer
    Dim situation As String

    situation = Sheet("Ventes").getCellRangeByName("D" & ligne).String
    total = 0

    For i = 0 To 11

        Dim qte As Double
        qte = Sheet("Ventes").getCellByPosition(4 + i * 2, ligne - 1).Value

        If qte > 0 Then
            total = total + qte * PrixProduit(2 + i, situation)
        End If

    Next i

    CalculTotalVente = total

End Function

