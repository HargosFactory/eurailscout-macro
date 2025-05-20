Option Explicit

' Classe ProductCollection - Collection de produits
Private pProducts As Collection

' Initialisation
Private Sub Class_Initialize()
    Set pProducts = New Collection
End Sub

' Nettoyage
Private Sub Class_Terminate()
    Set pProducts = Nothing
End Sub

' Ajouter une entrée de produit
Public Sub Add(product As Product)
    pProducts.Add product
End Sub

' Obtenir une entrée par index
Public Function Item(index As Integer) As Product
    Set Item = pProducts.Item(index)
End Function

' Obtenir le nombre d'entrées
Public Property Get Count() As Long
    Count = pProducts.Count
End Property

' Récupérer les produits par client et par mois
Public Function GetProductsByClientMois(client As String, mois As Integer, annee As Integer) As Collection
    Dim result As New Collection
    Dim product As Product
    
    For Each product In pProducts
        If product.Client = client And product.Mois = mois And product.Annee = annee Then
            result.Add product
        End If
    Next
    
    Set GetProductsByClientMois = result
End Function

' Calculer les montants totaux par mois pour tous les clients
Public Function GetMontantMonthTotal(mois As Integer, annee As Integer) As Double
    Dim total As Double
    Dim product As Product
    
    total = 0
    
    For Each product In pProducts
        If product.Mois = mois And product.Annee = annee Then
            total = total + product.MontantMois
        End If
    Next
    
    GetMontantMonthTotal = total
End Function

' Récupérer toutes les entrées pour un domaine fonctionnel
Public Function GetByDomaineFonctionnel(domaine As String) As Collection
    Dim result As New Collection
    Dim product As Product
    
    For Each product In pProducts
        If product.DomaineFonctionnel = domaine Then
            result.Add product
        End If
    Next
    
    Set GetByDomaineFonctionnel = result
End Function

' Récupérer toutes les entrées pour une nature comptable
Public Function GetByNatureComptable(nature As String) As Collection
    Dim result As New Collection
    Dim product As Product
    
    For Each product In pProducts
        If product.NatureComptable = nature Then
            result.Add product
        End If
    Next
    
    Set GetByNatureComptable = result
End Function

' Récupérer toutes les entrées pour un ID eOTP
Public Function GetByIdEOTP(idEOTP As String) As Collection
    Dim result As New Collection
    Dim product As Product
    
    For Each product In pProducts
        If product.IdEOTP = idEOTP Then
            result.Add product
        End If
    Next
    
    Set GetByIdEOTP = result
End Function