Sub CalculerStatsClient()

    Dim nomClient As String
    Dim i As Long
    Dim ligneStats As Long
    Dim totalVentes As Double
    Dim totalPaiements As Double
    
    Dim shStats As Object
    Dim shVentes As Object
    Dim shPaiements As Object
    
    shStats = Sheet("Stats")
    shVentes = Sheet("Ventes")
    shPaiements = Sheet("Paiements")
    
    ' Lire le client sélectionné
    nomClient = shStats.getCellRangeByName("C4").String
    
    ' Nettoyer anciennes données
    'RAZ_Stats
    
    totalVentes = 0
    totalPaiements = 0
    ligneStats = 10
    
    ' ========================
    ' RECHERCHE DANS VENTES
    ' ========================
    
    For i = 3 To DerniereLigne("Ventes")
        
        If shVentes.getCellRangeByName("A" & i).String = nomClient Then
            
            ' Date en colonne B
            shStats.getCellRangeByName("B" & ligneStats).Value = _
                shVentes.getCellRangeByName("B" & i).Value
            
            ' Total en colonne AB
            shStats.getCellRangeByName("D" & ligneStats).Value = _
            	shVentes.getCellRangeByName("Q" & i).Value
                'shVentes.getCellRangeByName("AC" & i).Value
            
            totalVentes = totalVentes + _
                shVentes.getCellRangeByName("Q" & i).Value
                'shVentes.getCellRangeByName("AC" & i).Value
            ligneStats = ligneStats + 1
            
        End If
        
    Next i
    
    ' Total ventes
    shStats.getCellRangeByName("D9").Value = totalVentes
    
    ' ========================
    ' RECHERCHE DANS PAIEMENTS
    ' ========================
    
    ligneStats = 10
    
    For i = 3 To DerniereLigne("Paiements")
        
        If shPaiements.getCellRangeByName("A" & i).String = nomClient Then
      
            ' Date paiement
            shStats.getCellRangeByName("G" & ligneStats).Value = _
                shPaiements.getCellRangeByName("C" & i).Value
            
            ' Montant paiement (col G)
            shStats.getCellRangeByName("H" & ligneStats).Value = _
                shPaiements.getCellRangeByName("G" & i).Value
            
            totalPaiements = totalPaiements + _
                shPaiements.getCellRangeByName("G" & i).Value
            
            ligneStats = ligneStats + 1
            
        End If
        
    Next i
    
    ' Total paiements
    shStats.getCellRangeByName("H9").Value = totalPaiements
    
    ' ========================
    ' SOLDE
    ' ========================
    
    shStats.getCellRangeByName("H5").Value = totalVentes - totalPaiements

End Sub
'-----------------------------------------------------------------------------------------------------

Sub BalanceGenerale()

    Dim shStats As Object
    Dim shVentes As Object
    Dim shPaiements As Object
    
    Dim i As Long
    Dim j As Long
    Dim nom As String
    Dim existe As Boolean
    Dim ligneStats As Long
       Dim sh As Object
   		 sh = Sheet("Stats") 
    
    
    shStats = ThisComponent.Sheets.getByName("Stats")
    shVentes = ThisComponent.Sheets.getByName("Ventes")
    shPaiements = ThisComponent.Sheets.getByName("Paiements")
    
   ' Nettoyage zone affichage
    
        For i = 10 To 200
        sh.getCellRangeByName("K" & i).Value = 0
        sh.getCellRangeByName("M" & i).Value = 0
	    Next i
       
    ligneStats = 10
    
    ' =========================
    ' CONSTRUIRE LISTE UNIQUE
    ' =========================
    
    For i = 3 To DerniereLigne("Ventes")
        
        nom = Trim(shVentes.getCellRangeByName("A" & i).String)
        
        If nom <> "" Then
            
            existe = False
            
            ' Vérifier si déjà affiché
            For j = 10 To ligneStats - 1
                If shStats.getCellRangeByName("K" & j).String = nom Then
                    existe = True
                    Exit For
                End If
            Next j
            
            If Not existe Then
                
                Dim totalV As Double
                Dim totalP As Double
                
                totalV = SommeClientVentes(nom)
                totalP = SommeClientPaiements(nom)
                
                If totalV - totalP <> 0 Then
                    
                    shStats.getCellRangeByName("K" & ligneStats).String = nom
                    shStats.getCellRangeByName("L" & ligneStats).Value = totalV
                    shStats.getCellRangeByName("M" & ligneStats).Value = totalV - totalP
                    
                    ligneStats = ligneStats + 1
                    
                End If
                
            End If
            
        End If
        
    Next i

End Sub




'-----------------------------------------------------------------------------------------------------------------------------
Function SommeClientVentes(nomClient As String) As Double

    Dim i As Long
    Dim total As Double
    
    For i = 3 To DerniereLigne("Ventes")     
       
       If Sheet("Ventes").getCellRangeByName("A" & i).String = nomClient Then
            total = total + Sheet("Ventes").getCellRangeByName("Q" & i).Value
            'total = total + Sheet("Ventes").getCellRangeByName("AC" & i).Value
      	End If
    Next i
    
    SommeClientVentes = total

End Function
'-----------------------------------------------------------------------------------------------------------------------------
Function SommeClientPaiements(nomClient As String) As Double

    Dim i As Long
    Dim total As Double
    
    For i = 3 To DerniereLigne("Paiements")
        
        If Sheet("Paiements").getCellRangeByName("A" & i).String = nomClient Then
            total = total + Sheet("Paiements").getCellRangeByName("G" & i).Value
        End If
        
    Next i
    
    SommeClientPaiements = total
'msgbox (SommeClientPaiements)
End Function
'----------------------------------------------------------------------------------------------------------------------------
Sub RAZ_Stats()

    Dim sh As Object
    sh = Sheet("Stats")

    
    sh.getCellRangeByName("C4").String =""
    sh.getCellRangeByName("D9").String =""
    sh.getCellRangeByName("H9").String =""
    sh.getCellRangeByName("H5").String =""
    
    For i = 10 To 200
        sh.getCellRangeByName("B" & i).Value = 0
        sh.getCellRangeByName("D" & i).Value = 0
        sh.getCellRangeByName("G" & i).Value = 0
        sh.getCellRangeByName("H" & i).Value = 0
        sh.getCellRangeByName("K" & i).String = ""
        sh.getCellRangeByName("M" & i).Value = 0
    Next i

   

End Sub




