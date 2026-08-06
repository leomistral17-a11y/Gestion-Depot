REM  *****  BASIC  *****
Option Explicit
Dim nomDepositaire As String
Sub RemplirFactBlsDepot()  'ligne As Long

    Dim fb As Object
    Dim i As Integer
    Dim lDest As Integer
    Dim situation As String
    Dim ligne As Long
    Dim oSel As Object
	
    
    fb = Sheet("FactBls")
    
        ' Récupérer la cellule sélectionnée
    oSel = ThisComponent.getCurrentSelection()
   'msgbox (oSel)
    ' Numéro de ligne (attention index commence à 0)
    ligne = oSel.CellAddress.Row + 1
   '
    ' --- Infos générales ---

   ' fb.getCellRangeByName("D11").String = numFact
    fb.getCellRangeByName("G4").String = Sheet("Depots").getCellRangeByName("A" & ligne).String 
    fb.getCellRangeByName("I11").String = Date
	nomDepositaire = Sheet("Depots").getCellRangeByName("A" & ligne).String 
        
        Sheet("FactBls").getCellRangeByName("B11").String = "Bon  de dépôt"
        Sheet("FactBls").getCellRangeByName("B14").String = "Matériels"
     
   
    situation = Sheet("Depots").getCellRangeByName("D" & ligne).String
   ' msgbox (Situation)
    ' Ligne de départ pour les produits
    lDest = 15
    
    ' On boucle sur 13 produits
    For i = 0 To 11
        
        Dim qte As Double
        Dim prix As Double
        Dim totalLigne As Double
      
        
        ' Lecture quantité dans Depots
        'qte = Sheet("Depots").getCellByPosition(4 + i * 2, ligne - 1).Value
        ' i de 5 à 12
        qte = Sheet("Depots").getCellByPosition(4 + i , ligne - 1).Value
    
        
        If qte > 0 Then
            
            ' Nom produit (colonne A feuille Produits)
            fb.getCellRangeByName("B" & lDest).String = _
                Sheet("Produits").getCellByPosition(0, 1 + i).String
            
            ' Quantité
            fb.getCellRangeByName("G" & lDest).Value = qte
            
            ' Prix unitaire
            prix = PrixProduit(1 + i, situation)
            fb.getCellRangeByName("I" & lDest).Value = prix
            
            ' Total ligne
            totalLigne = qte * prix
            fb.getCellRangeByName("J" & lDest).Value = totalLigne
            
            lDest = lDest + 1
            
        End If
        
    Next i
    
End Sub

Sub ExporterDepotPDF()

    Dim oDoc As Object
    Dim sURL As String
    Dim oArgs(0) As Object
    Dim nomDepositaire As String
    
		     Dim oModele As Object
   			 Dim oSheet As Object

   			 Dim cheminModele As String
  			  Dim cheminPDF As String

  			  Dim args()
   			Dim pdfArgs(0) As Variant
    
    
    oDoc = ThisComponent
    
    ' On relit le nom depuis la feuille FactBls
    nomDepositaire = ThisComponent.Sheets.getByName("FactBls") _
        .getCellRangeByName("G4").String
    
    If nomDepositaire = "" Then
        MsgBox "Nom dépositaire vide."
        Exit Sub
    End If
    
    
    
 	If Sheet("FactBls").getCellRangeByName("B11").String = "Facture" Then
    'MsgBox "Informations Export incomplètes ou erronnées."
    MsgBox "Mauvais bouton !!!"
    Exit Sub
	End If   
    
    
'    sURL = ConvertToURL("/home/jmp/Bureau/Global-dépot-25-26/Dépôts/2026/ " & nomDepositaire & ".pdf")
'
 '   oArgs(0) = CreateUnoStruct("com.sun.star.beans.PropertyValue")
 '   oArgs(0).Name = "FilterName"
 '   oArgs(0).Value = "calc_pdf_Export"

 '   oDoc.storeToURL sURL, oArgs()
 '   MsgBox "Votre Bon de dépôt " & nomDepositaire & " a bien été exporté !"
 
 	' création d'un fichier intermédiaire momentanné pour séparer la partie "gestion" de la partie édition des factures
    ' -----------------------------
    ' chemin modèle
    ' -----------------------------

    cheminModele = ConvertToURL( _
           "/home/jmp/DONNEES/a#Rad/Entreprise/GLOBAL-Depot/Modele_Depot.ods")

    ' -----------------------------
    ' ouverture modèle
    ' -----------------------------

    oModele = StarDesktop.loadComponentFromURL( _
        cheminModele, "_blank", 0, Array())

    oSheet = oModele.Sheets.getByName("Facture")
 '   oSheet = oModele.Sheets.getByName("Depots")
    ' -----------------------------
    ' remplissage données
    ' -----------------------------

    oSheet.getCellRangeByName("D11").String = _
        Sheet("FactBls").getCellRangeByName("D11").String
    oSheet.getCellRangeByName("B11").String = _
        Sheet("FactBls").getCellRangeByName("B11").String        
        

    oSheet.getCellRangeByName("G4").String = _
        Sheet("FactBls").getCellRangeByName("G4").String
  oSheet.getCellRangeByName("G5").String = _
        Sheet("FactBls").getCellRangeByName("G5").String       
  oSheet.getCellRangeByName("G6").String = _
        Sheet("FactBls").getCellRangeByName("G6").String       

    oSheet.getCellRangeByName("I11").String = _
        Sheet("FactBls").getCellRangeByName("I11").String
        
        

    ' -----------------------------
    ' lignes produits
    ' -----------------------------

    Dim i As Integer

    For i = 15 To 27

        oSheet.getCellRangeByName("B" & i).String = _
            Sheet("FactBls").getCellRangeByName("B" & i).String

        oSheet.getCellRangeByName("G" & i).Value = _
            Sheet("FactBls").getCellRangeByName("G" & i).Value

        oSheet.getCellRangeByName("I" & i).Value = _
            Sheet("FactBls").getCellRangeByName("I" & i).Value
            
          oSheet.getCellRangeByName("J" & i).Value = _
            Sheet("FactBls").getCellRangeByName("J" & i).Value                   

    Next i

        oSheet.getCellRangeByName("J29").Value = _
        Sheet("FactBls").getCellRangeByName("J29").Value

    ' -----------------------------
    ' export PDF
    ' -----------------------------
     
     cheminPDF = ConvertToURL( _
  "/home/jmp/DONNEES/a#Rad/Entreprise/GLOBAL-Depot/Dépôts/2026/Bons de dépot/" & _
  oSheet.getCellRangeByName("D11").String & _
  oSheet.getCellRangeByName("G4").String & ".pdf" )
        

    pdfArgs(0) = CreateUnoStruct("com.sun.star.beans.PropertyValue")

    pdfArgs(0).Name = "FilterName"
    pdfArgs(0).Value = "calc_pdf_Export"

    oModele.storeToURL cheminPDF, pdfArgs()

    ' -----------------------------
' contrôle taille PDF
' -----------------------------

Dim cheminLocal As String
Dim taille As Long

'cheminLocal = "/home/jmp/Documents/Facturation/PDF/" & _
'              oSheet.getCellRangeByName("D11").String & ".pdf"
              
'          cheminLocal = ("/home/jmp/DONNEES/a#Rad/Entreprise/GLOBAL-Depot/2026/Bons de dépot/" & _
'  oSheet.getCellRangeByName("D11").String & " -" & _
'  oSheet.getCellRangeByName("G4").String & ".pdf" )
          

'taille = FileLen(cheminLocal)
taille = FileLen(cheminPDF)

'If taille > 102400 Then
'
 '   MsgBox _
 '       "ATTENTION : le PDF dépasse 100 Ko" & Chr(10) & _
 '       "Taille : " & Round(taille / 1024, 1) & " Ko", _
 '       48, _
 '       "Alerte taille PDF"'

'Else

'    MsgBox _
'        "PDF OK : " & Int(taille / 102.4) / 10 & " Ko"

'End If


    ' -----------------------------
    ' fermeture
    ' -----------------------------

   oModele.close(True)

    MsgBox "Bon de dépôt PDF exporté, " & Chr(10) &" taille : " & Int(taille / 102.4) / 10 & " Ko"
    

 '  MsgBox "Dépôt PDF exporté"

End Sub       
Sub NettoyerFactBlsDepot()

	Sheet("FactBls").getCellRangeByName("G4").String =""
	Sheet("FactBls").getCellRangeByName("D11").String =""	
	Sheet("FactBls").getCellRangeByName("I11").String =""
	Sheet("FactBls").getCellRangeByName("B11").String =""
	Sheet("FactBls").getCellRangeByName("B14").String =""

Dim j As Integer
   Dim oRange As Object
    oRange = Sheet("FactBls").getCellRangeByName("B15:J27")
    oRange.clearContents( _
        com.sun.star.sheet.CellFlags.STRING + _
        com.sun.star.sheet.CellFlags.VALUE+ _
    com.sun.star.sheet.CellFlags.FORMULA   )
End Sub
'-------------------------------------------------------------------------------------------------------------------------

Sub BalanceGenerale()

    'Dim shStats As Object
    Dim shDepots As Object
    Dim shVentes As Object
    Dim shPaiements As Object
    
    Dim i As Long
    Dim j As Long
    Dim nom As String
    Dim existe As Boolean
    'Dim ligneStats As Long
    Dim ligneDepots As Long
    Dim sh As Object
    
   		 'sh = Sheet("Stats") 
   		 sh = Sheet("Depots")
    
    'shStats = ThisComponent.Sheets.getByName("Stats")
    shDepots = ThisComponent.Sheets.getByName("Depots")   
   ' shVentes = ThisComponent.Sheets.getByName("Ventes")
    shPaiements = ThisComponent.Sheets.getByName("Paiements")
    
   ' Nettoyage zone affichage
    
        For i = 3 To 100
       ' sh.getCellRangeByName("K" & i).Value = 0
        sh.getCellRangeByName("R" & i).Value = 0
	    Next i
       
    ligneDepots = 3
    
    ' =========================
    ' CONSTRUIRE LISTE UNIQUE
    ' =========================
    
   'For i = 3 To DerniereLigne("Ventes")
    For i = 3 To DerniereLigne("Depots")
        
        'nom = Trim(shVentes.getCellRangeByName("A" & i).String)
		  nom = Trim(shDepots.getCellRangeByName("A" & i).String)
        
        If nom <> "" Then
            
            existe = False
            
            ' Vérifier si déjà affiché
            For j = 3 To ligneDepots - 1
               ' If shStats.getCellRangeByName("K" & j).String = nom Then

                If shDepots.getCellRangeByName("R" & j).String = nom Then
                    existe = True
                    Exit For
                End If
            Next j
            
            If Not existe Then
                
                Dim totalV As Double
                Dim totalP As Double
                
                totalV = SommeClientDepots(nom)
               ' totalP = SommeClientPaiements(nom)
                 totalP = SommeDepotPaiements(nom)
                
                If totalV - totalP <> 0 Then
                    
                   ' shStats.getCellRangeByName("K" & ligneStats).String = nom
                   ' shStats.getCellRangeByName("L" & ligneStats).Value = totalV
                    shDepots.getCellRangeByName("R" & ligneDepots).Value = totalV - totalP
                    
                    ligneDepots = ligneDepots + 1
                    
                End If
                
            End If
            
        End If
        
    Next i

End Sub



'-----------------------------------------------------------------------------------------------------------------------------
Function SommeClientDepots(nomClient As String) As Double

    Dim i As Long
    Dim total As Double
    
    'For i = 3 To DerniereLigne("Ventes")  
    For i = 3 To DerniereLigne("Depots")   
       
       'If Sheet("Ventes").getCellRangeByName("A" & i).String = nomClient Then
		 If Sheet("Depots").getCellRangeByName("A" & i).String = nomClient Then

            'total = total + Sheet("Ventes").getCellRangeByName("Q" & i).Value
            total = total + Sheet("Depots").getCellRangeByName("Q" & i).Value
            'total = total + 'Sheet("Ventes").getCellRangeByName("AC" & i).Value
		    Sheet("Depots").getCellRangeByName("R" & i).Value
		        If Sheet("Depots").getCellRangeByName("R" & i).Value  = "0" Then
   				 Sheet("Depots").getCellRangeByName("R" & i).String = "    Réglé"
				End If
      	End If
    Next i
    
    'SommeClientVentes = total
	 SommeClientDepots = total


End Function
'-----------------------------------------------------------------------------------------------------------------------------
'Function SommeClientPaiements(nomClient As String) As Double
Function SommeDepotPaiements(nomClient As String) As Double
    Dim i As Long
    Dim total As Double
    
    For i = 3 To DerniereLigne("Paiements")
        
        'If Sheet("Paiements").getCellRangeByName("A" & i).String = nomClient Then	 
         If Sheet("Paiements").getCellRangeByName("B" & i).String = nomClient Then
            total = total + Sheet("Paiements").getCellRangeByName("G" & i).Value
        End If
        
    Next i
    
 '   SommeClientPaiements = total
    SommeDepotPaiements = total   
   ' If  total = "0" Then
   ' Sheet("Depots").getCellRangeByName("R" & i).String = "---"
'msgbox (SommeDepotPaiements)
End Function
