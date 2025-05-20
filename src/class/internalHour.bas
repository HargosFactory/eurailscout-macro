Option Explicit

' Classe InternalHour - Représente les heures internes pour un employé par mois
Private pPrestation As String
Private pIdEOTP As String
Private pFonction As String
Private pNiveauService As String
Private pNom As String
Private pTJM_H As Double
Private pTJM_J As Double
Private pTauxFG As Double
Private pHeuresMois As Double
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

Public Property Get Fonction() As String
    Fonction = pFonction
End Property

Public Property Let Fonction(value As String)
    pFonction = value
End Property

Public Property Get NiveauService() As String
    NiveauService = pNiveauService
End Property

Public Property Let NiveauService(value As String)
    pNiveauService = value
End Property

Public Property Get Nom() As String
    Nom = pNom
End Property

Public Property Let Nom(value As String)
    pNom = value
End Property

Public Property Get TJM_H() As Double
    TJM_H = pTJM_H
End Property

Public Property Let TJM_H(value As Double)
    pTJM_H = value
End Property

Public Property Get TJM_J() As Double
    TJM_J = pTJM_J
End Property

Public Property Let TJM_J(value As Double)
    pTJM_J = value
End Property

Public Property Get TauxFG() As Double
    TauxFG = pTauxFG
End Property

Public Property Let TauxFG(value As Double)
    pTauxFG = value
End Property

Public Property Get HeuresMois() As Double
    HeuresMois = pHeuresMois
End Property

Public Property Let HeuresMois(value As Double)
    pHeuresMois = value
End Property

Public Property Get Mois() As Integer
    Mois = pMois
End Property

Public Property Let Mois(value As Integer)
    If value >= 1 And value <= 12 Then
        pMois = value
    Else
        Err.Raise 1000, "InternalHour", "Le mois doit être entre 1 et 12"
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

' Méthode pour initialiser un objet InternalHour
Public Sub Initialize(ByVal prestation As String, ByVal idEOTP As String, ByVal fonction As String, _
                     ByVal niveauService As String, ByVal nom As String, ByVal tjmH As Double, _
                     ByVal tjmJ As Double, ByVal tauxFG As Double, ByVal heuresMois As Double, _
                     ByVal mois As Integer, ByVal annee As Integer, ByVal domaineFonctionnel As String)
    
    Me.Prestation = prestation
    Me.IdEOTP = idEOTP
    Me.Fonction = fonction
    Me.NiveauService = niveauService
    Me.Nom = nom
    Me.TJM_H = tjmH
    Me.TJM_J = tjmJ
    Me.TauxFG = tauxFG
    Me.HeuresMois = heuresMois
    Me.Mois = mois
    Me.Annee = annee
    Me.DomaineFonctionnel = domaineFonctionnel
End Sub

' Calcule le montant total pour ce mois
Public Function CalculerMontantTotal() As Double
    CalculerMontantTotal = Me.HeuresMois * Me.TJM_H * (1 + Me.TauxFG / 100)
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
               "Fonction: " & Me.Fonction & vbCrLf & _
               "Niveau Service: " & Me.NiveauService & vbCrLf & _
               "Nom: " & Me.Nom & vbCrLf & _
               "TJM (H): " & Me.TJM_H & vbCrLf & _
               "TJM (J): " & Me.TJM_J & vbCrLf & _
               "Taux FG: " & Format(Me.TauxFG, "0.00%") & vbCrLf & _
               "Heures: " & Me.HeuresMois & " (" & Me.GetNomMois() & " " & Me.Annee & ")" & vbCrLf & _
               "Domaine fonctionnel: " & Me.DomaineFonctionnel
End Function
