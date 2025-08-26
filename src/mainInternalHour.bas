Option Explicit

' Function to find the start of the table
Function FindTableStart(ws As Worksheet) As Integer
    Dim i As Integer
    Dim lastRow As Integer
    
    ' Déterminer la dernière ligne de la feuille dans la colonne A
    lastRow = define_end_sheet(ws, ws.Range("A1"))
    
    ' Chercher les cellules "Heures internes" suivies de "Prestation"
    For i = 1 To lastRow - 1
        If ws.Cells(i, 1).value = "Heures internes" And ws.Cells(i + 1, 1).value = "Prestation" Then
            ' Le tableau commence à la ligne suivante de "Prestation"
            FindTableStart = i + 2
            Exit Function
        End If
    Next i
    
    ' Si on n'a pas trouvé le début du tableau
    FindTableStart = 0
End Function

' Procedure to parse a row and create an InternalHour object
Sub ParseRow(ws As Worksheet, rowNum As Integer, collection As InternalHourCollection)
    Dim internalHour As internalHour
    Dim prestation As String
    Dim idEOTP As String
    Dim fonction As String
    Dim niveauService As String
    Dim nom As String
    Dim tjmH As Double
    Dim tjmJ As Double
    Dim tauxFG As Double
    Dim domaineFonctionnel As String
    Dim mois As Integer
    
    ' Récupérer les valeurs des colonnes
    prestation = ws.Cells(rowNum, 1).value
    idEOTP = ws.Cells(rowNum, 3).value
    fonction = ws.Cells(rowNum, 4).value
    niveauService = ws.Cells(rowNum, 5).value
    nom = ws.Cells(rowNum, 6).value
    tjmH = CDbl(ws.Cells(rowNum, 7).value)
    tjmJ = CDbl(ws.Cells(rowNum, 8).value)
    tauxFG = CDbl(ws.Cells(rowNum, 9).value)
    domaineFonctionnel = ws.Cells(rowNum, 22).value
    
    ' Créer un objet InternalHour pour chaque mois avec des heures (colonnes 9 à 20)
    For mois = 1 To 12
        Dim heuresMois As Double
        heuresMois = 0
        
        ' Vérifier si la cellule contient une valeur
        If Not IsEmpty(ws.Cells(rowNum, 9 + mois).value) Then
            heuresMois = CDbl(ws.Cells(rowNum, 9 + mois).value)
            
            ' Ne créer un objet que si des heures sont enregistrées pour ce mois
            If heuresMois > 0 Then
                Set internalHour = New internalHour
                internalHour.Initialize prestation, idEOTP, fonction, niveauService, nom, _
                                     tjmH, tjmJ, tauxFG, heuresMois, mois, 2025, domaineFonctionnel
                collection.Add internalHour
            End If
        End If
    Next mois
End Sub

' Function to extract the data and return a collection
Function ExtractInternalHoursToCollection() As InternalHourCollection
    Dim ws As Worksheet
    Dim lastRow As Integer
    Dim startRow As Integer
    Dim collection As New InternalHourCollection
    Dim i As Integer
    
    ' Utiliser la feuille active
    Set ws = ActiveSheet
    
    ' Trouver le début du tableau
    startRow = FindTableStart(ws)
    If startRow = 0 Then
        MsgBox "Impossible de trouver le début du tableau des heures internes.", vbExclamation
        Exit Function
    End If
    
    ' Déterminer la dernière ligne du tableau
    lastRow = search_last_row(ws, ws.Range("A" & startRow), startRow)
    
    ' Analyser chaque ligne du tableau
    For i = startRow To lastRow
        ' Vérifier que nous avons une ligne valide (avec une prestation)
        If Not IsEmpty(ws.Cells(i, 1).value) Then
            ParseRow ws, i, collection
        End If
    Next i
    
    Set ExtractInternalHoursToCollection = collection
End Function

' Procédure pour exporter une collection d'heures internes vers une feuille Excel au format SAP
Sub ExportInternalHoursToSAP(collection As InternalHourCollection, Optional targetWorksheet As Worksheet = Nothing)
    Dim ws As Worksheet
    Dim row As Integer
    Dim i As Integer
    Dim internalHour As internalHour

    ' Initialiser les comptes de transcodification
    InitAccountTransco ActiveSheet
    
    ' Déterminer la feuille cible
    If targetWorksheet Is Nothing Then
        ' Créer une nouvelle feuille si aucune n'est spécifiée
        Set ws = Worksheets.Add
        ws.Name = "Export SAP HI " & Format(Now(), "yyyy-mm-dd")
    Else
        Set ws = targetWorksheet
    End If
    
    ' Ajouter les en-têtes
    AddSAPHeaders ws
    
    ' Commencer à partir de la ligne 4 (après les en-têtes)
    row = 4
    
    ' Parcourir la collection et ajouter les lignes
    For i = 1 To collection.Count
        Set internalHour = collection.Item(i)

        ' Créer deux lignes pour chaque entrée (une pour chaque compte - personnel et FG)
        AddSAPLine ws, row, internalHour, globaleAccountTranscoInstance.CompteHeuresDuPersonnel
        row = row + 1

        AddSAPLine ws, row, internalHour, globaleAccountTranscoInstance.CompteFGHeuresInternes
        row = row + 1
    Next i
    
    ' Ajuster les colonnes pour une meilleure lisibilité
    ws.Columns("A:L").AutoFit
    
    ' Informer l'utilisateur
    MsgBox "Export terminé avec succès. " & (row - 4) & " lignes créées.", vbInformation
End Sub

' Ajouter les en-têtes SAP à la feuille
Sub AddSAPHeaders(ws As Worksheet)
    ' Ligne 1: Noms des colonnes
    ws.Cells(1, 1).value = "CATEGORY"
    ws.Cells(1, 2).value = "RYEAR"
    ws.Cells(1, 3).value = "POPER"
    ws.Cells(1, 4).value = "RBUKRS"
    ws.Cells(1, 5).value = "PS_PSPID"
    ws.Cells(1, 6).value = "PS_POSID"
    ws.Cells(1, 7).value = "RACCT"
    ws.Cells(1, 8).value = "HSL"
    ws.Cells(1, 9).value = "RHCUR"
    ws.Cells(1, 10).value = "YY1_NatureDeDepense_JEI"
    ws.Cells(1, 11).value = "RFAREA"
    
    ' Ligne 2: Description des colonnes
    ws.Cells(2, 1).value = "Catégorie de budget"
    ws.Cells(2, 2).value = "Exercice du grand livre"
    ws.Cells(2, 3).value = "Période comptable"
    ws.Cells(2, 4).value = "Société"
    ws.Cells(2, 5).value = "Définition de projet"
    ws.Cells(2, 6).value = "Élément d'organigramme technique de projet (élément d'OTP)"
    ws.Cells(2, 7).value = "Numéro de compte"
    ws.Cells(2, 8).value = "Montant en devise globale"
    ws.Cells(2, 9).value = "Devise globale"
    ws.Cells(2, 10).value = "Nature de dépenses"
    ws.Cells(2, 11).value = "Domaine fonctionnel"
    
    ' Ligne 3: Marqueurs X pour les colonnes concernées
    ws.Cells(3, 1).value = "X"
    ws.Cells(3, 2).value = "X"
    ws.Cells(3, 5).value = "X"
    ws.Cells(3, 10).value = "X"
    
    ' Formater les en-têtes
    ws.Range("A1:K2").Font.Bold = True
    ws.Range("A3:K3").Font.Italic = True
End Sub

' Ajouter une ligne de données SAP
Sub AddSAPLine(ws As Worksheet, row As Integer, internalHour As internalHour, accountNumber As String)
    Dim montant As Double
    Dim category As String
    Dim psPspid As String

    If accountNumber = globaleAccountTranscoInstance.CompteFGHeuresInternes Then
        montant = internalHour.heuresMois * internalHour.TJM_H
    Else
        montant = internalHour.heuresMois * internalHour.TJM_H * internalHour.tauxFG
    End If

    montant = RoundToDigits(montant, 2)
    category = GetCategoryValue()
    psPspid = GetPSPSPIDValue()
    
    ' Remplir les colonnes
    ws.Cells(row, 1).value = category ' CATEGORY
    ws.Cells(row, 2).value = internalHour.annee ' RYEAR
    ws.Cells(row, 3).value = internalHour.mois ' POPER
    ws.Cells(row, 4).value = 1000 ' RBUKRS (valeur fixe)
    ws.Cells(row, 5).value = psPspid ' PS_PSPID
    ws.Cells(row, 6).value = internalHour.idEOTP ' PS_POSID
    ws.Cells(row, 7).value = accountNumber ' RACCT
    ws.Cells(row, 8).value = montant ' HSL
    ws.Cells(row, 9).value = "EUR" ' RHCUR (valeur fixe)
    ws.Cells(row, 10).value = internalHour.nom ' YY1_NatureDeDepense_JEI
    ws.Cells(row, 11).value = internalHour.domaineFonctionnel ' RFAREA
End Sub

Function GetCategoryValue() As String
    ' Récupère la valeur de la cellule C1 de la feuille contenant "Prévisionnel"
    Dim categoryValue As String
    Dim prevSheet As Worksheet
    
    ' Trouver la feuille contenant "Prévisionnel" dans son nom
    Set prevSheet = FindSheetByNameContaining("Prévisionnel")
    
    If Not prevSheet Is Nothing Then
        ' Vérifier si la cellule contient une valeur
        If Not IsEmpty(prevSheet.Range("C1").Value) Then
            categoryValue = Trim(prevSheet.Range("C1").Value)
        Else
            ' Valeur par défaut si la cellule est vide
            categoryValue = "ERROR"
        End If
    Else
        ' Aucune feuille correspondante trouvée, utiliser la valeur par défaut
        categoryValue = "ERROR"
        Debug.Print "Avertissement: Aucune feuille contenant 'Prévisionnel' n'a été trouvée."
    End If
    
    GetCategoryValue = categoryValue
End Function

Function GetPSPSPIDValue() As String
    ' Récupère la valeur de la cellule C2 de la feuille contenant "Prévisionnel"
    Dim pspspidValue As String
    Dim prevSheet As Worksheet
    
    ' Trouver la feuille contenant "Prévisionnel" dans son nom
    Set prevSheet = FindSheetByNameContaining("Prévisionnel")
    
    If Not prevSheet Is Nothing Then
        ' Vérifier si la cellule contient une valeur
        If Not IsEmpty(prevSheet.Range("C2").Value) Then
            pspspidValue = Trim(prevSheet.Range("C2").Value)
        Else
            ' Valeur par défaut si la cellule est vide
            pspspidValue = "ERROR"
        End If
    Else
        ' Aucune feuille correspondante trouvée, utiliser la valeur par défaut
        pspspidValue = "ERROR"
        Debug.Print "Avertissement: Aucune feuille contenant 'Prévisionnel' n'a été trouvée."
    End If
    
    GetPSPSPIDValue = pspspidValue
End Function

' Procédure principale pour extraire et exporter les données
Sub ExtractAndExportInternalHourToSAP()
    Dim collection As InternalHourCollection
    
    ' Extraire les données
    Set collection = ExtractInternalHoursToCollection()
    
    ' Si l'extraction a réussi, exporter les données
    If Not collection Is Nothing Then
        ExportInternalHoursToSAP collection
    End If
End Sub


