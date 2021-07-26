Attribute VB_Name = "mdl_DadosAleatorios"
Option Explicit

Function randomString(Length As Integer)
    'PURPOSE: Create a Randomized String of Characters
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

    Dim CharacterBank As Variant
    Dim X As Long
    Dim str As String

    'Test Length Input
    If Length < 1 Then
        MsgBox "Length variable must be greater than 0"
        Exit Function
    End If

    CharacterBank = Array("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", _
                          "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", _
                          "y", "z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "!", "@", _
                          "#", "$", "%", "^", "&", "*", "A", "B", "C", "D", "E", "F", "G", "H", _
                          "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", _
                          "W", "X", "Y", "Z")
  

    'Randomly Select Characters One-by-One
    For X = 1 To Length
        Randomize
        str = str & CharacterBank(Int((UBound(CharacterBank) - LBound(CharacterBank) + 1) * Rnd + LBound(CharacterBank)))
    Next X

    'Output Randomly Generated String
    randomString = str

End Function

Function randomInteger(pMinimo, pMaximo) As Integer
    randomInteger = Int(pMinimo + Rnd() * (pMaximo - pMinimo))
End Function

Function randomLong(pMinimo, pMaximo) As Long
    randomLong = CLng(pMinimo + Rnd() * (pMaximo - pMinimo))
End Function

Function randomDouble(pMinimo, pMaximo) As Double

    randomDouble = pMinimo + Rnd() * (pMaximo - pMinimo)
    randomDouble = Round(randomDouble, 2)
    
End Function

Function randomDate() As Date

    Dim dia, ano As Integer
    Dim mes As String

    dia = randomInteger(10, 28)
    mes = "0" & randomInteger(1, 9)
    ano = randomInteger(2000, 2020)

    randomDate = dia & "/" & mes & "/" & ano

End Function


