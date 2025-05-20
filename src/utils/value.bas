Option Explicit

' Description: Cette fonction convertit un pourcentage en son complément à 100.
' Par exemple: 15 -> 85, 30 -> 70, etc.
' Parameters:
'   - percentValue: La valeur du pourcentage (ex: 15 pour 15%)
' Returns: Le complément à 100 du pourcentage (ex: 85 pour un input de 15)
Function GetComplementPercentage(percentValue As Double) As Double
    GetComplementPercentage = 100 - percentValue
End Function

' Description: Cette fonction convertit un pourcentage en facteur décimal.
' Par exemple: 15 -> 0.15, 30 -> 0.30, etc.
' Parameters:
'   - percentValue: La valeur du pourcentage (ex: 15 pour 15%)
' Returns: Le facteur décimal correspondant (ex: 0.15 pour un input de 15)
Function GetDecimalFactor(percentValue As Double) As Double
    GetDecimalFactor = percentValue / 100
End Function

' Description: Cette fonction convertit un pourcentage en son complément en facteur décimal.
' Par exemple: 15 -> 0.85, 30 -> 0.70, etc.
' Parameters:
'   - percentValue: La valeur du pourcentage (ex: 15 pour 15%)
' Returns: Le complément du facteur décimal (ex: 0.85 pour un input de 15)
Function GetComplementDecimalFactor(percentValue As Double) As Double
    GetComplementDecimalFactor = (100 - percentValue) / 100
End Function

' Description: Cette fonction convertit un facteur décimal en son complément à 1.
' Par exemple: 0.15 -> 0.85, 0.30 -> 0.70, etc.
' Parameters:
'   - decimalFactor: Le facteur décimal (ex: 0.15)
' Returns: Le complément à 1 du facteur décimal (ex: 0.85 pour un input de 0.15)
Function GetComplementDecimal(decimalFactor As Double) As Double
    GetComplementDecimal = 1 - decimalFactor
End Function

' Description: Cette fonction inverse le signe d'un nombre.
' Par exemple: 50 -> -50, -30 -> 30, etc.
' Parameters:
'   - value: La valeur à inverser
' Returns: La valeur avec son signe inversé
Function InvertDoubleValue(value As Double) As Double
    InvertDoubleValue = -value
End Function