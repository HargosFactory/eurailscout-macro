Option Explicit

' Function to find the start of the Products table
Function FindProductsTableStart(ws As Worksheet) As Integer
    Dim i As Integer
    Dim lastRow As Integer
    
    ' Déterminer la dernière ligne de la feuille dans la colonne A
    lastRow = define_end_sheet(ws, ws.Range("A1"))
    
    ' Chercher la cellule "Produit" suivie de "Prestation"
    For i = 1 To lastRow - 1
        If ws.Cells(i, 1).Value = "Produit" And ws.Cells(i + 1, 1).Value = "Prestation" Then
            ' Le tableau commence à la ligne suivante de "Prestation"
            FindProductsTableStart = i + 2
            Exit Function
        End If
    Next i
    
    ' Si on n'a pas trouvé le début du tableau
    FindProductsTableStart = 0
End Function

' Procedure to parse a row and create a Product object
Sub ParseProductRow(ws As Worksheet, rowNum As Integer, collection As ProductCollection)
    Dim product As Product
    Dim prestation As String
    Dim idEOTP As String
    Dim montantAnnuel As Double
    Dim produitNom As String
    Dim client As String
    Dim natureComptable As String
    Dim domaineFonctionnel As String
    Dim mois As Integer
    
    ' Récupérer les valeurs des colonnes
    prestation = ws.Cells(rowNum, 1).Value
    idEOTP = ws.Cells(rowNum, 2).Value
    montantAnnuel = CDbl(Replace(ws.Cells(rowNum, 3).Value, "€", ""))
    produitNom = ws.Cells(rowNum, 4).Value
    client = ws.Cells(rowNum, 5).Value
    natureComptable = ws.Cells(rowNum, 6).Value
    domaineFonctionnel = ws.Cells(rowNum, 21).Value
    
    ' Créer un objet Product pour chaque mois avec des montants (colonnes 9 à 20)
    For mois = 1 To 12
        Dim montantMois As Double
        montantMois = 0
        
        ' Vérifier si la cellule contient une valeur
        If Not IsEmpty(ws.Cells(rowNum, 8 + mois).Value) Then
            ' Gérer les formats différents (avec ou sans €)
            If InStr(ws.Cells(rowNum, 8 + mois).Value, "€") > 0 Then
                montantMois = CDbl(Replace(Replace(ws.Cells(rowNum, 8 + mois).Value, "€", ""), " ", ""))
            Else
                montantMois = CDbl(ws.Cells(rowNum, 8 + mois).Value)
            End If
            
            ' Ne créer un objet que si des montants sont enregistrés pour ce mois
            If montantMois > 0 Then
                Set product = New Product
                product.Initialize prestation, idEOTP, montantAnnuel, produitNom, client, _
                                 natureComptable, montantMois, mois, 2025, domaineFonctionnel
                collection.Add product
            End If
        End If
    Next mois
End Sub

' Function to extract the data and return a collection
Function ExtractProductsToCollection() As ProductCollection
    Dim ws As Worksheet
    Dim lastRow As Integer
    Dim startRow As Integer
    Dim collection As New ProductCollection
    Dim i As Integer
    
    ' Utiliser la feuille active
    Set ws = ActiveSheet
    
    ' Trouver le début du tableau
    startRow = FindProductsTableStart(ws)
    
    If startRow = 0 Then
        MsgBox "Impossible de trouver le début du tableau des produits.", vbExclamation
        Set ExtractProductsToCollection = collection
        Exit Function
    End If
    
    ' Déterminer la dernière ligne du tableau
    lastRow = search_last_row(ws, ws.Range("A" & startRow), startRow)
    
    ' Analyser chaque ligne du tableau
    For i = startRow To lastRow
        ' Vérifier que nous avons une ligne valide
        If Not IsEmpty(ws.Cells(i, 1).Value) Then
            ParseProductRow ws, i, collection
        End If
    Next i
    
    Set ExtractProductsToCollection = collection
End Function

Sub ExportProductsToSAP(collection As ProductCollection, Optional targetWorksheet As Worksheet = Nothing)
    Dim ws As Worksheet
    Dim i As Integer
    Dim row As Integer
    Dim product As Product

    ' Déterminer la feuille cible
    If targetWorksheet Is Nothing Then
        ' Créer une nouvelle feuille si aucune n'est spécifiée
        Set ws = Worksheets.Add
        ws.Name = "Export SAP P " & Format(Now(), "yyyy-mm-dd")
    Else
        Set ws = targetWorksheet
    End If

    ' Ajouter les en-têtes
    AddSAPHeaders ws
    
    ' Commencer à partir de la ligne 4 (après les en-têtes)
    row = 4
    
    ' Parcourir chaque produit et remplir les lignes SAP
    For i = 1 To collection.Count
        Set product = collection.Item(i)
        ' Pour les produits, nous n'avons besoin que d'une seule ligne par entrée
        AddSAPLineProduct ws, row, product
        row = row + 1
    Next i

     ' Ajuster les colonnes pour une meilleure lisibilité
    ws.Columns("A:L").AutoFit
    
    ' Informer l'utilisateur
    MsgBox "Export terminé avec succès. " & (row - 4) & " lignes créées.", vbInformation
End Sub

Sub AddSAPLineProduct(ws As Worksheet, row As Integer, product As Product)
    Dim category As String
    Dim psPspid As String
    Dim montant As Double

    category = GetCategoryValue()
    psPspid = GetPSPSPIDValue()

    montant = InvertDoubleValue(product.MontantMois)
    
    ' Remplir les colonnes
    ws.Cells(row, 1).Value = category ' CATEGORY
    ws.Cells(row, 2).Value = product.Annee ' RYEAR
    ws.Cells(row, 3).Value = product.Mois ' POPER
    ws.Cells(row, 4).Value = 1000 ' RBUKRS (valeur fixe)
    ws.Cells(row, 5).Value = psPspid ' PS_PSPID
    ws.Cells(row, 6).Value = product.IdEOTP ' PS_POSID
    ws.Cells(row, 7).Value = product.NatureComptable ' RACCT
    ws.Cells(row, 8).Value = montant ' HSL
    ws.Cells(row, 9).Value = "EUR" ' RHCUR (valeur fixe)
    ws.Cells(row, 10).Value = product.Client ' YY1_NatureDeDepense_JEI (produit dans ce cas)
    ws.Cells(row, 11).Value = product.DomaineFonctionnel ' RFAREA
End Sub

' Procédure principale visible pour les utilisateurs
Sub ExtractAndExportProductsToSAP()
    Dim collection As ProductCollection

    Set collection = ExtractProductsToCollection()

    ' Vérifier si la collection n'est pas vide avant d'exporter
    If Not collection Is Nothing And collection.Count > 0 Then
        ExportProductsToSAP collection
    Else
        MsgBox "Aucun produit n'a été trouvé à exporter.", vbExclamation
    End If
End Sub