Option Explicit

' Classe ExternalChargeCollection - Collection de charges externes
Private pCharges As Collection

' Initialisation
Private Sub Class_Initialize()
    Set pCharges = New Collection
End Sub

' Nettoyage
Private Sub Class_Terminate()
    Set pCharges = Nothing
End Sub

' Ajouter une entrée de charges externes
Public Sub Add(externalCharge As ExternalCharge)
    pCharges.Add externalCharge
End Sub

' Obtenir une entrée par index
Public Function Item(index As Integer) As ExternalCharge
    Set Item = pCharges.Item(index)
End Function

' Obtenir le nombre d'entrées
Public Property Get Count() As Long
    Count = pCharges.Count
End Property

' Récupérer les charges par fournisseur et par mois
Public Function GetChargesByFournisseurMois(fournisseur As String, mois As Integer, annee As Integer) As Collection
    Dim result As New Collection
    Dim charge As ExternalCharge
    
    For Each charge In pCharges
        If charge.Fournisseur = fournisseur And charge.Mois = mois And charge.Annee = annee Then
            result.Add charge
        End If
    Next
    
    Set GetChargesByFournisseurMois = result
End Function

' Calculer les montants totaux par mois pour tous les fournisseurs
Public Function GetMontantMonthTotal(mois As Integer, annee As Integer) As Double
    Dim total As Double
    Dim charge As ExternalCharge
    
    total = 0
    
    For Each charge In pCharges
        If charge.Mois = mois And charge.Annee = annee Then
            total = total + charge.MontantMois
        End If
    Next
    
    GetMontantMonthTotal = total
End Function

' Calculer le montant total avec frais par mois
Public Function GetMontantTotalAvecFrais(mois As Integer, annee As Integer) As Double
    Dim total As Double
    Dim charge As ExternalCharge
    
    total = 0
    
    For Each charge In pCharges
        If charge.Mois = mois And charge.Annee = annee Then
            total = total + charge.CalculerMontantTotal()
        End If
    Next
    
    GetMontantTotalAvecFrais = total
End Function

' Récupérer toutes les entrées pour un domaine fonctionnel
Public Function GetByDomaineFonctionnel(domaine As String) As Collection
    Dim result As New Collection
    Dim charge As ExternalCharge
    
    For Each charge In pCharges
        If charge.DomaineFonctionnel = domaine Then
            result.Add charge
        End If
    Next
    
    Set GetByDomaineFonctionnel = result
End Function

' Récupérer toutes les entrées pour une nature comptable
Public Function GetByNatureComptable(nature As String) As Collection
    Dim result As New Collection
    Dim charge As ExternalCharge
    
    For Each charge In pCharges
        If charge.NatureComptable = nature Then
            result.Add charge
        End If
    Next
    
    Set GetByNatureComptable = result
End Function

' Récupérer toutes les entrées pour un ID eOTP
Public Function GetByIdEOTP(idEOTP As String) As Collection
    Dim result As New Collection
    Dim charge As ExternalCharge
    
    For Each charge In pCharges
        If charge.IdEOTP = idEOTP Then
            result.Add charge
        End If
    Next
    
    Set GetByIdEOTP = result
End Function