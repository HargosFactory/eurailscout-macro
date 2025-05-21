Option Explicit

' Function to find the start of the External Charges table
Function FindExternalChargesTableStart(ws As Worksheet) As Integer
    Dim i As Integer
    Dim lastRow As Integer
    
    ' Déterminer la dernière ligne de la feuille dans la colonne A
    lastRow = define_end_sheet(ws, ws.Range("A1"))
    
    ' Chercher la cellule "Charges externes"
    For i = 1 To lastRow - 1
        If ws.Cells(i, 1).value = "Charges externes" And ws.Cells(i + 1, 1).value = "Prestation" Then
            ' Le tableau commence à la ligne suivante de "Prestation"
            FindExternalChargesTableStart = i + 2
            Exit Function
        End If
    Next i
    
    ' Si on n'a pas trouvé le début du tableau
    FindExternalChargesTableStart = 0
End Function

' Procedure to parse a row and create an ExternalCharge object
Sub ParseExternalChargeRow(ws As Worksheet, rowNum As Integer, collection As ExternalChargeCollection)
    Dim externalCharge As externalCharge
    Dim prestation As String
    Dim idEOTP As String
    Dim montantAnnuel As Double
    Dim groupeMarchandise As String
    Dim fournisseur As String
    Dim natureComptable As String
    Dim description As String
    Dim tauxFR As Double
    Dim mois As Integer
    
    ' Récupérer les valeurs des colonnes
    prestation = ws.Cells(rowNum, 1).value
    idEOTP = ws.Cells(rowNum, 2).value
    montantAnnuel = CDbl(Replace(ws.Cells(rowNum, 3).value, "€", ""))
    groupeMarchandise = ws.Cells(rowNum, 4).value
    fournisseur = ws.Cells(rowNum, 5).value
    natureComptable = ws.Cells(rowNum, 6).value
    description = ws.Cells(rowNum, 7).value
    
    ' Convertir le taux FR de format % à nombre décimal
    tauxFR = 0
    If Not IsEmpty(ws.Cells(rowNum, 8).value) Then
        If InStr(ws.Cells(rowNum, 8).value, "%") > 0 Then
            tauxFR = CDbl(Replace(ws.Cells(rowNum, 8).value, "%", ""))
        Else
            tauxFR = CDbl(ws.Cells(rowNum, 8).value)
        End If
    End If
    
    Dim domaineFonctionnel As String
    domaineFonctionnel = ws.Cells(rowNum, 21).value
    
    ' Créer un objet ExternalCharge pour chaque mois avec des montants (colonnes 9 à 20)
    For mois = 1 To 12
        Dim montantMois As Double
        montantMois = 0
        
        ' Vérifier si la cellule contient une valeur
        If Not IsEmpty(ws.Cells(rowNum, 8 + mois).value) Then
            ' Gérer les formats différents (avec ou sans €)
            If InStr(ws.Cells(rowNum, 8 + mois).value, "€") > 0 Then
                montantMois = CDbl(Replace(Replace(ws.Cells(rowNum, 8 + mois).value, "€", ""), " ", ""))
            Else
                montantMois = CDbl(ws.Cells(rowNum, 8 + mois).value)
            End If
            
            ' Ne créer un objet que si des montants sont enregistrés pour ce mois
            If montantMois > 0 Then
                Set externalCharge = New externalCharge
                externalCharge.Initialize prestation, idEOTP, montantAnnuel, groupeMarchandise, _
                                      fournisseur, natureComptable, description, tauxFR, _
                                      montantMois, mois, 2025, domaineFonctionnel
                collection.Add externalCharge
            End If
        End If
    Next mois
End Sub

' Function to extract the data and return a collection
Function ExtractExternalChargesToCollection() As ExternalChargeCollection
    Dim ws As Worksheet
    Dim lastRow As Integer
    Dim startRow As Integer
    Dim collection As New ExternalChargeCollection
    Dim i As Integer
    
    ' Utiliser la feuille active
    Set ws = ActiveSheet
    
    ' Trouver le début du tableau
    startRow = FindExternalChargesTableStart(ws)
    
    If startRow = 0 Then
        MsgBox "Impossible de trouver le début du tableau des charges externes.", vbExclamation
        Set ExtractExternalChargesToCollection = collection
        Exit Function
    End If
    
    ' Déterminer la dernière ligne du tableau
    lastRow = search_last_row(ws, ws.Range("A" & startRow), startRow)
    
    ' Analyser chaque ligne du tableau
    For i = startRow To lastRow
        ' Vérifier que nous avons une ligne valide (avec une prestation)
        If Not IsEmpty(ws.Cells(i, 1).value) Then
            ParseExternalChargeRow ws, i, collection
        End If
    Next i
    
    Set ExtractExternalChargesToCollection = collection
End Function

Sub ExportExternalChargesToSAP(collection As ExternalChargeCollection, Optional targetWorksheet As Worksheet = Nothing)
    Dim ws As Worksheet
    Dim i As Integer
    Dim row As Integer
    Dim externalCharge As externalCharge

    ' Déterminer la feuille cible
    If targetWorksheet Is Nothing Then
        ' Créer une nouvelle feuille si aucune n'est spécifiée
        Set ws = Worksheets.Add
        ws.Name = "Export SAP CE " & Format(Now(), "yyyy-mm-dd")
    Else
        Set ws = targetWorksheet
    End If

    ' Ajouter les en-têtes
    AddSAPHeaders ws
    
    ' Commencer à partir de la ligne 4 (après les en-têtes)
    row = 4
    
    ' Parcourir chaque charge externe et remplir les lignes SAP
    For i = 1 To collection.Count
        Set externalCharge = collection.Item(i)
        AddSAPLineExternalCharge ws, row, externalCharge, "99991540"
        row = row + 1

        AddSAPLineExternalCharge ws, row, externalCharge, "99991517"
        row = row + 1
    Next i

     ' Ajuster les colonnes pour une meilleure lisibilité
    ws.Columns("A:L").AutoFit
    
    ' Informer l'utilisateur
    MsgBox "Export terminé avec succès. " & (row - 4) & " lignes créées.", vbInformation
End Sub

Sub AddSAPLineExternalCharge(ws As Worksheet, row As Integer, externalCharge As externalCharge, accountNumber As String)
    Dim montant As Double
    Dim category As String
    Dim psPspid As String

    If accountNumber = "99991540" Then
        montant = externalCharge.montantMois
    Else
        montant = externalCharge.montantMois * externalCharge.tauxFR
    End If

    category = GetCategoryValue()
    psPspid = GetPSPSPIDValue()
    
    ' Remplir les colonnes
    ws.Cells(row, 1).value = category ' CATEGORY
    ws.Cells(row, 2).value = externalCharge.annee ' RYEAR
    ws.Cells(row, 3).value = externalCharge.mois ' POPER
    ws.Cells(row, 4).value = 1000 ' RBUKRS (valeur fixe)
    ws.Cells(row, 5).value = psPspid ' PS_PSPID
    ws.Cells(row, 6).value = externalCharge.idEOTP ' PS_POSID
    ws.Cells(row, 7).value = externalCharge.natureComptable ' RACCT
    ws.Cells(row, 8).value = montant ' HSL
    ws.Cells(row, 9).value = "EUR" ' RHCUR (valeur fixe)
    ws.Cells(row, 10).value = externalCharge.fournisseur ' YY1_NatureDeDepense_JEI
    ws.Cells(row, 11).value = externalCharge.domaineFonctionnel ' RFAREA
End Sub

' Procédure principale pour extraire et exporter les charges externes
Sub ExtractAndExportExternalChargeToSAP()
    Dim collection As ExternalChargeCollection

    Set collection = ExtractExternalChargesToCollection()

    ' Vérifier si la collection n'est pas vide avant d'exporter
    If Not collection Is Nothing Then
        ExportExternalChargesToSAP collection
    End If
End Sub

