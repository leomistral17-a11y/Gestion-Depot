REM  *****  BASIC  *****

Option Explicit

Function GenererNumeroFacture() As String

    Dim an As String
    Dim cpt As Long

    an = Right(Year(Now), 2)
    cpt = Sheet("Produits").getCellRangeByName("K7").Value

    GenererNumeroFacture = an & "-" & cpt     ' "F-" & an & "-" & cpt

    Sheet("Produits").getCellRangeByName("K7").Value = cpt + 1

End Function


'Sub RemplirJournalFactures(ligneVente As Long, numFact As String)
Function RemplirJournalFactures(ligneVente As Long, numFact As String) As Long

    Dim f As Object
    Dim l As Long
    Dim i As Integer
    Dim qte As Double
    Dim pu As Double
    Dim situation As String

    
    f = Sheet("Factures")
    
    ' Prochaine ligne libre
    l = DerniereLigne("Factures") + 1
    
    ' ---------
    ' En-tête
    ' ---------
    

    f.getCellRangeByName("A" & l).String = numFact
    f.getCellRangeByName("B" & l).String = Sheet("Ventes").getCellRangeByName("A" & ligneVente).String   'nom client
    f.getCellRangeByName("D" & l).String = Sheet("Ventes").getCellRangeByName("D" & ligneVente).String
    
    situation = Sheet("Ventes").getCellRangeByName("D" & ligneVente).String
    f.getCellRangeByName("D" & l).String = situation
    
    f.getCellRangeByName("E" & l).String = Date
    
    ' ---------
    ' Détail produits
    ' ---------
    
    For i = 0 To 11
        
        'Lire quantité
        'qte = Sheet("Ventes").getCellByPosition(4 + i * 2, ligneVente - 1).Value
        qte = Sheet("Ventes").getCellByPosition(4 + i , ligneVente - 1).Value
      
        
        ' Lire prix unitaire
        pu = PrixProduit(1 + i, situation)
         
        ' Stocker quantité (colonne G + i*2)
        f.getCellByPosition(6 + i * 2, l - 1).Value = qte  '6
    
        
        ' Stocker prix unitaire (colonne H + i*2)
         If qte <> "0" Then
       f.getCellByPosition(7 +i * 2, l - 1).Value = pu  '7
		end If
        
    Next i
    
    ' Total (colonne F)
    f.getCellRangeByName("F" & l).Value = Sheet("Ventes").getCellRangeByName("Q" & ligneVente).String   ' CalculTotalVente(ligneVente)

    ' 👉 on retourne la ligne créée
    RemplirJournalFactures = l

End Function
'End Sub



Sub CreerFactureDepuisSelection()

    Dim oSel As Object
    Dim ligneVente As Long
    Dim numFact As String
    Dim ligneFacture As Long
    Dim oCell As Object
    
    NettoyerFactBls
    
    ' Récupérer ligne vente sélectionnée
    oSel = ThisComponent.getCurrentSelection()
    ligneVente = oSel.CellAddress.Row + 1
    
    ' Vérifier déjà facturé
 	' If Sheet("Ventes").getCellRangeByName("AE" & ligneVente).String = "FACTURÉ" Then
  	' Sheet("Ventes").getCellRangeByName("AE" & ligneVente).String = numFact Then
    ' MsgBox "Cette vente est déjà facturée."
  	 'Exit Sub
     'End If  

    
    ' Générer numéro
    numFact = GenererNumeroFacture()
	'msgbox (numFact)
    ' Créer ligne facture et récupérer son numéro de ligne
    ligneFacture = RemplirJournalFactures(ligneVente, numFact)
    
    ' Marquer comme facturé
    Sheet("Ventes").getCellRangeByName("S" & ligneVente).String = numFact  ' "FACTURÉ"
    'Sheet("Ventes").getCellRangeByName("AE" & ligneVente).String = numFact

    
    ' 👉 Aller sur la feuille Factures
    ThisComponent.CurrentController.setActiveSheet(Sheet("Factures"))
    
    ' 👉 Sélectionner la cellule du numéro facture créé
    oCell = Sheet("Factures").getCellRangeByName("A" & ligneFacture)
    ThisComponent.CurrentController.select(oCell)
    
    ' 👉 Remplir FactBls depuis cette ligne
    RemplirFactBlsDepuisJournal ligneFacture
    
    MsgBox "Facture " & numFact & " créée et prête à imprimer."

End Sub




Sub ReimprimerFactureDepuisSelection()
	NettoyerFactBls
	
    Dim oSel As Object
    Dim ligne As Long
    
        Sheet("FactBls").getCellRangeByName("B11").String = "Facture"
        Sheet("FactBls").getCellRangeByName("B14").String = "Matériels"
    
    ' Récupérer la cellule sélectionnée
    oSel = ThisComponent.getCurrentSelection()
    
    ' Numéro de ligne (attention index commence à 0)
    ligne = oSel.CellAddress.Row + 1
    
    ' Appel de la vraie macro avec paramètre
    RemplirFactBlsDepuisJournal ligne

End Sub


