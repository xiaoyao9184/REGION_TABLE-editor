Attribute VB_Name = "Module1"
Option Explicit
Public SelectListNO%
Public R As Byte, G As Byte, B As Byte, T As Byte, TRGB$
Public fontNO As Byte
Public one(3) As Byte
Public icon10(9, 3) As Integer

Public Sub load_data(i)
Dim q%
Select Case i
Case 1
        For q = 0 To 3
            icon10(0, q) = one(q)
        Next q
Case 9
        For q = 0 To 3
            icon10(1, q) = one(q)
        Next q
Case 17
        For q = 0 To 3
            icon10(2, q) = one(q)
        Next q
Case 25
        For q = 0 To 3
            icon10(3, q) = one(q)
        Next q
Case 33
        For q = 0 To 3
            icon10(4, q) = one(q)
        Next q
Case 41
        For q = 0 To 3
            icon10(5, q) = one(q)
        Next q
Case 49
        For q = 0 To 3
            icon10(6, q) = one(q)
        Next q
Case 57
        For q = 0 To 3
            icon10(7, q) = one(q)
        Next q
Case 65
        For q = 0 To 3
            icon10(8, q) = one(q)
        Next q
Case 73
        For q = 0 To 3
            icon10(9, q) = one(q)
        Next q
End Select
End Sub
Public Function HEXtoDEC(NO1, NO2)
If NO1 = "A" Then
    NO1 = 10
    ElseIf NO1 = "B" Then
    NO1 = 11
    ElseIf NO1 = "C" Then
    NO1 = 12
    ElseIf NO1 = "D" Then
    NO1 = 13
    ElseIf NO1 = "E" Then
    NO1 = 14
    ElseIf NO1 = "F" Then
    NO1 = 15
    Else
    NO1 = NO1
End If
If NO2 = "A" Then
    NO2 = 10
    ElseIf NO2 = "B" Then
    NO2 = 11
    ElseIf NO2 = "C" Then
    NO2 = 12
    ElseIf NO2 = "D" Then
    NO2 = 13
    ElseIf NO2 = "E" Then
    NO2 = 14
    ElseIf NO2 = "F" Then
    NO2 = 15
    Else
    NO2 = NO2
End If
HEXtoDEC = NO1 * 16 + NO2
End Function
