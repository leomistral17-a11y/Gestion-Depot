

REM  *****  BASIC  *****

Option Explicit
Dim nomClientFact As String
Dim numFact as String
Sub RemplirFactBlsDepuisJournal(ligneFacture As Long)

    Dim fb As Object
    Dim f As Object
    Dim i As Integer
    Dim lDest As Integer
    Dim qte As Double
    Dim pu As Double
    
    fb = Sheet("FactBls")
    f = Sheet("Factures")
    
    ' Infos générales
    fb.getCellRangeByName("D11").String = f.getCellRangeByName("A" & ligneFacture).String 'n°facture

    fb.getCellRangeByName("G4").String = f.getCellRangeByName("B" & ligneFacture).String 'nom client
    fb.getCellRangeByName("I11").String = f.getCellRangeByName("E" & ligneFacture).String 'date

        Sheet("FactBls").getCellRangeByName("B11").String = "Facture"
        Sheet("FactBls").getCellRangeByName("B14").String = "Matériels"
        
       
    
    lDest = 15
    
    For i = 0 To 11
        
        qte = f.getCellByPosition(6 + i * 2, ligneFacture - 1).Value
        pu = f.getCellByPosition(7 + i * 2, ligneFacture - 1).Value
        
        If qte > 0 Then
            
            fb.getCellRangeByName("B" & lDest).String = _
                Sheet("Produits").getCellByPosition(0, 1 + i).String
                
            fb.getCellRangeByName("G" & lDest).Value = qte
            fb.getCellRangeByName("I" & lDest).Value = pu
            fb.getCellRangeByName("J" & lDest).Value = qte * pu
            
            lDest = lDest + 1
            
        End If
        
    Next i

End Sub



Sub ExporterFacturePDF()
 'Option Explicit

    Dim oModele As Object
    Dim oSheet As Object

    Dim cheminModele As String
    Dim cheminPDF As String

    Dim args()
    Dim pdfArgs(0) As Variant
    
        Dim oDoc As Object
    Dim sURL As String
    Dim oArgs(0) As Object
    Dim nomFacture As String
        
    oDoc = ThisComponent
    
    ' On relit le nom depuis la feuille FactBls
    nomFacture = ThisComponent.Sheets.getByName("FactBls") _
        .getCellRangeByName("G4").String
    
    If nomFacture = "" Then
        MsgBox "Nom Facture vide."
        Exit Sub
    End If
    
  '  msgbox (Sheet("FactBls").getCellRangeByName("B11").String)
    
 	If Sheet("FactBls").getCellRangeByName("B11").String = "Bon  de dépôt" Then
    'MsgBox "Informations Export incomplètes ou erronnées."
    MsgBox "Mauvais bouton !!!"
    Exit Sub
	End If  
     


	' création d'un fichier intermédiaire momentanné pour séparer la partie "gestion" de la partie édition des factures
    ' -----------------------------
    ' chemin modèle
    ' -----------------------------

    cheminModele = ConvertToURL( _
           "/home/jmp/DONNEES/a#Rad/Entreprise/GLOBAL-Depot/Modele_Facture.ods")

    ' -----------------------------
    ' ouverture modèle
    ' -----------------------------

    oModele = StarDesktop.loadComponentFromURL( _
        cheminModele, "_blank", 0, Array())

    oSheet = oModele.Sheets.getByName("Facture")

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

    For i = 15 To 25

        oSheet.getCellRangeByName("B" & i).String = _
            Sheet("FactBls").getCellRangeByName("B" & i).String

        oSheet.getCellRangeByName("G" & i).Value = _
            Sheet("FactBls").getCellRangeByName("G" & i).Value

        oSheet.getCellRangeByName("I" & i).Value = _
            Sheet("FactBls").getCellRangeByName("I" & i).Value
            
          oSheet.getCellRangeByName("J" & i).Value = _
            Sheet("FactBls").getCellRangeByName("J" & i).Value                   

    Next i
    
        oSheet.getCellRangeByName("J27").Value = _
        Sheet("FactBls").getCellRangeByName("J29").Value

    ' -----------------------------
    ' export PDF
    ' -----------------------------
     
     cheminPDF = ConvertToURL( _
  "/home/jmp/DONNEES/a#Rad/Entreprise/GLOBAL-Depot/factures/2026/" & _
  oSheet.getCellRangeByName("D11").String & " -" & _
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
              
'          cheminLocal = ("/home/jmp/DONNEES/a#Rad/Entreprise/GLOBAL-Depot/factures/2026/" & _
'  oSheet.getCellRangeByName("D11").String & " -" & _
'  oSheet.getCellRangeByName("G4").String & ".pdf" )
          

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

    MsgBox "Facture PDF exportée, " & Chr(10) &" taille : " & Int(taille / 102.4) / 10 & " Ko"
    

End Sub       
Sub NettoyerFactBls()

	Sheet("FactBls").getCellRangeByName("G4").String =""
	Sheet("FactBls").getCellRangeByName("D11").String =""	
	Sheet("FactBls").getCellRangeByName("I11").String =""
	Sheet("FactBls").getCellRangeByName("B11").String =""
	Sheet("FactBls").getCellRangeByName("B14").String =""


    Dim oRange As Object
    oRange = Sheet("FactBls").getCellRangeByName("B15:J25")

    oRange.clearContents( _
        com.sun.star.sheet.CellFlags.STRING + _
        com.sun.star.sheet.CellFlags.VALUE )
        
        End Sub



        
