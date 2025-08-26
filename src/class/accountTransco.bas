Option Explicit

' Classe AccountTransco - Représente la transcodification des comptes
Private pCompteHeuresDuPersonnel As String
Private pCompteFGHeuresInternes As String
Private pCompteFRChargesExternes As String
Private pCompteFraisFinanciers As String
Private pCompteDotationsAuxAmortissements As String

' Propriétés
Public Property Get CompteHeuresDuPersonnel() As String
    CompteHeuresDuPersonnel = pCompteHeuresDuPersonnel
End Property

Public Property Let CompteHeuresDuPersonnel(value As String)
    pCompteHeuresDuPersonnel = value
End Property

Public Property Get CompteFGHeuresInternes() As String
    CompteFGHeuresInternes = pCompteFGHeuresInternes
End Property

Public Property Let CompteFGHeuresInternes(value As String)
    pCompteFGHeuresInternes = value
End Property

Public Property Get CompteFRChargesExternes() As String
    CompteFRChargesExternes = pCompteFRChargesExternes
End Property

Public Property Let CompteFRChargesExternes(value As String)
    pCompteFRChargesExternes = value
End Property

Public Property Get CompteFraisFinanciers() As String
    CompteFraisFinanciers = pCompteFraisFinanciers
End Property

Public Property Let CompteFraisFinanciers(value As String)
    pCompteFraisFinanciers = value
End Property

Public Property Get CompteDotationsAuxAmortissements() As String
    CompteDotationsAuxAmortissements = pCompteDotationsAuxAmortissements
End Property

Public Property Let CompteDotationsAuxAmortissements(value As String)
    pCompteDotationsAuxAmortissements = value
End Property

' Méthode pour initialiser les comptes de transcodification
Public Sub Initialize(compteHeures As String, compteFG As String, compteFR As String, _
                      compteFraisFinanciers As String, compteDotations As String)
    pCompteHeuresDuPersonnel = compteHeures
    pCompteFGHeuresInternes = compteFG
    pCompteFRChargesExternes = compteFR
    pCompteFraisFinanciers = compteFraisFinanciers
    pCompteDotationsAuxAmortissements = compteDotations
End Sub

Public Function ToString() As String
    ToString = "Compte Heures du Personnel: " & pCompteHeuresDuPersonnel & vbCrLf & _
               "Compte FG Heures Internes: " & pCompteFGHeuresInternes & vbCrLf & _
               "Compte FR Charges Externes: " & pCompteFRChargesExternes & vbCrLf & _
               "Compte Frais Financiers: " & pCompteFraisFinanciers & vbCrLf & _
               "Compte Dotations aux Amortissements: " & pCompteDotationsAuxAmortissements
End Function