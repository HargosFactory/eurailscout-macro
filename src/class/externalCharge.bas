Option Explicit

' Classe ExternalCharge - Représente les charges externes par mois
Private pPrestation As String
Private pIdEOTP As String
Private pMontantAnnuel As Double
Private pGroupeMarchandise As String
Private pFournisseur As String
Private pNatureComptable As String
Private pDescription As String
Private pTauxFR As Double
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

Public Property Get GroupeMarchandise() As String
    GroupeMarchandise = pGroupeMarchandise
End Property

Public Property Let GroupeMarchandise(value As String)
    pGroupeMarchandise = value
End Property

Public Property Get Fournisseur() As String
    Fournisseur = pFournisseur
End Property

Public Property Let Fournisseur(value As String)
    pFournisseur = value
End Property

Public Property Get NatureComptable() As String
    NatureComptable = pNatureComptable
End Property

Public Property Let NatureComptable(value As String)
    pNatureComptable = value
End Property

Public Property Get Description() As String
    Description = pDescription
End Property

Public Property Let Description(value As String)
    pDescription = value
End Property

Public Property Get TauxFR() As Double
    TauxFR = pTauxFR
End Property

Public Property Let TauxFR(value As Double)
    pTauxFR = value
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
        Err.Raise 1000, "ExternalCharge", "Le mois doit être entre 1 et 12"
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

' Méthode pour initialiser un objet ExternalCharge
Public Sub Initialize(ByVal prestation As String, ByVal idEOTP As String, ByVal montantAnnuel As Double, _
                     ByVal groupeMarchandise As String, ByVal fournisseur As String, _
                     ByVal natureComptable As String, ByVal description As String, _
                     ByVal tauxFR As Double, ByVal montantMois As Double, _
                     ByVal mois As Integer, ByVal annee As Integer, ByVal domaineFonctionnel As String)
    
    Me.Prestation = prestation
    Me.IdEOTP = idEOTP
    Me.MontantAnnuel = montantAnnuel
    Me.GroupeMarchandise = groupeMarchandise
    Me.Fournisseur = fournisseur
    Me.NatureComptable = natureComptable
    Me.Description = description
    Me.TauxFR = tauxFR
    Me.MontantMois = montantMois
    Me.Mois = mois
    Me.Annee = annee
    Me.DomaineFonctionnel = domaineFonctionnel
End Sub

' Calcule le montant total pour ce mois avec le taux de frais
Public Function CalculerMontantTotal() As Double
    CalculerMontantTotal = Me.MontantMois * (1 + Me.TauxFR / 100)
End Function

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
               "Groupe de marchandise: " & Me.GroupeMarchandise & vbCrLf & _
               "Fournisseur: " & Me.Fournisseur & vbCrLf & _
               "Nature comptable: " & Me.NatureComptable & vbCrLf & _
               "Description: " & Me.Description & vbCrLf & _
               "Taux FR: " & Format(Me.TauxFR, "0.00%") & vbCrLf & _
               "Montant: " & Format(Me.MontantMois, "#,##0.00 €") & " (" & Me.GetNomMois() & " " & Me.Annee & ")" & vbCrLf & _
               "Domaine fonctionnel: " & Me.DomaineFonctionnel
End Function