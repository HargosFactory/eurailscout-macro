Public globaleAccountTranscoInstance As AccountTransco

Sub InitAccountTransco(ws As Worksheet)
    Dim localAccountTransco As AccountTransco
    Dim compteHeures As String
    Dim compteFG As String
    Dim compteFR As String
    Dim compteFraisFinanciers As String
    Dim compteDotations As String

    compteHeures = ws.Cells(1, 24).Value
    compteFG = ws.Cells(2, 24).Value
    compteFR = ws.Cells(3, 24).Value
    compteFraisFinanciers = ws.Cells(4, 24).Value
    compteDotations = ws.Cells(5, 24).Value

    Set localAccountTransco = New AccountTransco
    localAccountTransco.Initialize compteHeures, compteFG, compteFR, compteFraisFinanciers, compteDotations

    Set globaleAccountTranscoInstance = localAccountTransco
End Sub