Option Explicit

' Classe InternalHourCollection - Collection d'heures internes
Private pHeures As Collection

' Initialisation
Private Sub Class_Initialize()
    Set pHeures = New Collection
End Sub

' Nettoyage
Private Sub Class_Terminate()
    Set pHeures = Nothing
End Sub

' Ajouter une entrée d'heures internes
Public Sub Add(internalHour As InternalHour)
    pHeures.Add internalHour
End Sub

' Obtenir une entrée par index
Public Function Item(index As Integer) As InternalHour
    Set Item = pHeures.Item(index)
End Function

' Obtenir le nombre d'entrées
Public Property Get Count() As Long
    Count = pHeures.Count
End Property

' Récupérer les heures par employé et par mois
Public Function GetHeuresByEmployeeMois(nomEmploye As String, mois As Integer, annee As Integer) As Collection
    Dim result As New Collection
    Dim heure As InternalHour
    
    For Each heure In pHeures
        If heure.Nom = nomEmploye And heure.Mois = mois And heure.Annee = annee Then
            result.Add heure
        End If
    Next
    
    Set GetHeuresByEmployeeMois = result
End Function

' Calculer les heures totales par mois pour tous les employés
Public Function GetHeuresMonthTotal(mois As Integer, annee As Integer) As Double
    Dim total As Double
    Dim heure As InternalHour
    
    total = 0
    
    For Each heure In pHeures
        If heure.Mois = mois And heure.Annee = annee Then
            total = total + heure.HeuresMois
        End If
    Next
    
    GetHeuresMonthTotal = total
End Function

' Calculer le montant total par mois
Public Function GetMontantMonthTotal(mois As Integer, annee As Integer) As Double
    Dim total As Double
    Dim heure As InternalHour
    
    total = 0
    
    For Each heure In pHeures
        If heure.Mois = mois And heure.Annee = annee Then
            total = total + heure.CalculerMontantTotal()
        End If
    Next
    
    GetMontantMonthTotal = total
End Function

' Récupérer toutes les entrées pour un domaine fonctionnel
Public Function GetByDomaineFonctionnel(domaine As String) As Collection
    Dim result As New Collection
    Dim heure As InternalHour
    
    For Each heure In pHeures
        If heure.DomaineFonctionnel = domaine Then
            result.Add heure
        End If
    Next
    
    Set GetByDomaineFonctionnel = result
End Function