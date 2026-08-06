REM  *****  BASIC  *****

Option Explicit

Function Sheet(nom As String) As Object
    Sheet = ThisComponent.Sheets.getByName(nom)
End Function


Function DerniereLigne(nomFeuille As String) As Long

    Dim i As Long
    i = 1

    Do While Sheet(nomFeuille).getCellRangeByName("A" & i).String <> ""
        i = i + 1
    Loop

    DerniereLigne = i - 1

End Function

