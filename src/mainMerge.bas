Function ExportToSAP(productCollection As ProductCollection, internalHourCollection As InternalHourCollection, externalChargeCollection As ExternalChargeCollection, Optional targetWorksheet As Worksheet = Nothing)
    Dim ws As Worksheet
    Dim i As Integer
    Dim row As Integer
    Dim product As Product
    Dim internalHour As InternalHour
    Dim externalCharge As ExternalCharge

    ' Initialiser les comptes de transcodification
    InitAccountTransco ActiveSheet

    ' Déterminer la feuille cible
    If targetWorksheet Is Nothing Then
        ' Créer une nouvelle feuille si aucune n'est spécifiée
        Set ws = Worksheets.Add
        ws.Name = "Export SAP " & Format(Now(), "yyyy-mm-dd")
    Else
        Set ws = targetWorksheet
    End If

    ' Ajouter les en-têtes
    AddSAPHeaders ws
    
    ' Commencer à partir de la ligne 4 (après les en-têtes)
    row = 4

    ' Parcourir chaque produit et remplir les lignes SAP
    For i = 1 To productCollection.Count
        Set product = productCollection.Item(i)
        AddSAPLineProduct ws, row, product
        row = row + 1
    Next i
    ' Parcourir chaque heure interne et remplir les lignes SAP
    For i = 1 To internalHourCollection.Count
        Set internalHour = internalHourCollection.Item(i)

        ' Créer deux lignes pour chaque entrée (une pour chaque compte - personnel et FG)
        AddSAPLine ws, row, internalHour, globaleAccountTranscoInstance.CompteHeuresDuPersonnel
        row = row + 1

        AddSAPLine ws, row, internalHour, globaleAccountTranscoInstance.CompteFGHeuresInternes
        row = row + 1
    Next i
    ' Parcourir chaque charge externe et remplir les lignes SAP
    For i = 1 To externalChargeCollection.Count
        Set externalCharge = externalChargeCollection.Item(i)

        ' Créer deux lignes pour chaque entrée (une pour chaque compte - nature comptable et FR)
        AddSAPLineExternalCharge ws, row, externalCharge, externalCharge.NatureComptable
        row = row + 1

        AddSAPLineExternalCharge ws, row, externalCharge, globaleAccountTranscoInstance.CompteFRChargesExternes
        row = row + 1
    Next i

    ' Ajuster les colonnes pour une meilleure lisibilité
    ws.Columns("A:L").AutoFit

    ' Informer l'utilisateur
    MsgBox "Export terminé avec succès. " & (row - 4) & " lignes créées.", vbInformation
End Function

' Procedure principal pour l'exportation vers SAP
Sub ExtractAndExportToSAP()
    Dim productCollection As ProductCollection
    Dim internalHourCollection As InternalHourCollection
    Dim externalChargeCollection As ExternalChargeCollection

    ' Extraire les données des feuilles
    Set productCollection = ExtractProductsToCollection()
    Set internalHourCollection = ExtractInternalHoursToCollection()
    Set externalChargeCollection = ExtractExternalChargesToCollection()

    If productCollection.Count = 0 Then
        MsgBox "Aucun produit trouvé à exporter.", vbExclamation
    End If

    If internalHourCollection.Count = 0 Then
        MsgBox "Aucune heure interne trouvée à exporter.", vbExclamation
    End If

    If externalChargeCollection.Count = 0 Then
        MsgBox "Aucune charge externe trouvée à exporter.", vbExclamation
    End If

    ' Exporter vers SAP
    ExportToSAP productCollection, internalHourCollection, externalChargeCollection
End Sub