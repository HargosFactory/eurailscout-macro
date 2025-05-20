Option Explicit

' Classe Product - Représente les produits par mois
Private pPrestation As String
Private pIdEOTP As String
Private pMontantAnnuel As Double
Private pProduit As String
Private pClient As String
Private pNatureComptable As String
Private pMontantMois As Double
Private pMois As Integer
Private pAnnee As Integer
Private pDomaineFonctionnel As String

' Propriétés
Public Property Get Prestation() As String
    Prestation = pPrestation
End Property

Public Property Let Prestation(value As String)
    pPrestation = value
End Property

Public Property Get IdEOTP() As String
    IdEOTP = pIdEOTP
End Property

Public Property Let IdEOTP(value As String)
    pIdEOTP = value
End Property

Public Property Get MontantAnnuel() As Double
    MontantAnnuel = pMontantAnnuel
End Property

Public Property Let MontantAnnuel(value As Double)
    pMontantAnnuel = value
End Property

Public Property Get Produit() As String
    Produit = pProduit
End Property

Public Property Let Produit(value As String)
    pProduit = value
End Property

Public Property Get Client() As String
    Client = pClient
End Property

Public Property Let Client(value As String)
    pClient = value
End Property

Public Property Get NatureComptable() As String
    NatureComptable = pNatureComptable
End Property

Public Property Let NatureComptable(value As String)
    pNatureComptable = value
End Property

Public Property Get MontantMois() As Double
    MontantMois = pMontantMois
End Property

Public Property Let MontantMois(value As Double)
    pMontantMois = value
End Property

Public Property Get Mois() As Integer
    Mois = pMois
End Property

Public Property Let Mois(value As Integer)
    If value >= 1 And value <= 12 Then
        pMois = value
    Else
        Err.Raise 1000, "Product", "Le mois doit être entre 1 et 12"
    End If
End Property

Public Property Get Annee() As Integer
    Annee = pAnnee
End Property

Public Property Let Annee(value As Integer)
    pAnnee = value
End Property

Public Property Get DomaineFonctionnel() As String
    DomaineFonctionnel = pDomaineFonctionnel
End Property

Public Property Let DomaineFonctionnel(value As String)
    pDomaineFonctionnel = value
End Property

' Méthode pour initialiser un objet Product
Public Sub Initialize(ByVal prestation As String, ByVal idEOTP As String, ByVal montantAnnuel As Double, _
                     ByVal produit As String, ByVal client As String, _
                     ByVal natureComptable As String, ByVal montantMois As Double, _
                     ByVal mois As Integer, ByVal annee As Integer, ByVal domaineFonctionnel As String)
    
    Me.Prestation = prestation
    Me.IdEOTP = idEOTP
    Me.MontantAnnuel = montantAnnuel
    Me.Produit = produit
    Me.Client = client
    Me.NatureComptable = natureComptable
    Me.MontantMois = montantMois
    Me.Mois = mois
    Me.Annee = annee
    Me.DomaineFonctionnel = domaineFonctionnel
End Sub

' Retourne le nom du mois (en français)
Public Function GetNomMois() As String
    Dim nomsMois As Variant
    nomsMois = Array("Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre")
    
    If Me.Mois >= 1 And Me.Mois <= 12 Then
        GetNomMois = nomsMois(Me.Mois - 1)
    Else
        GetNomMois = "Mois inconnu"
    End If
End Function

' Méthode pour créer une représentation en chaîne de l'objet
Public Function ToString() As String
    ToString = "Prestation: " & Me.Prestation & vbCrLf & _
               "ID eOTP (SAP): " & Me.IdEOTP & vbCrLf & _
               "Montant annuel: " & Format(Me.MontantAnnuel, "#,##0.00 €") & vbCrLf & _
               "Produit: " & Me.Produit & vbCrLf & _
               "Client: " & Me.Client & vbCrLf & _
               "Nature comptable: " & Me.NatureComptable & vbCrLf & _
               "Montant: " & Format(Me.MontantMois, "#,##0.00 €") & " (" & Me.GetNomMois() & " " & Me.Annee & ")" & vbCrLf & _
               "Domaine fonctionnel: " & Me.DomaineFonctionnel
End Function