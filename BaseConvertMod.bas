Attribute VB_Name = "BaseConvertMod"
'******************************************************************************
'                   BaseConvertMod
'
'=======================================
' Twister of Twisted Media
' E -mail: vincent_gw_lewis@ hotmail.com
' ICQ: 12674360
' URL: http://www.twistedmedia.f2s.com
'=======================================
'******************************************************************************

Option Explicit

Public Function HexToDec(ByVal HexStr As String) As Double
    Dim mult As Double
    Dim DecNum As Double
    Dim ch As String
    mult = 1
    DecNum = 0

    Dim i As Integer
    For i = Len(HexStr) To 1 Step -1
        ch = Mid(HexStr, i, 1)
        If (ch >= "0") And (ch <= "9") Then
            DecNum = DecNum + (Val(ch) * mult)
        Else
            If (ch >= "A") And (ch <= "F") Then
                DecNum = DecNum + ((Asc(ch) - Asc("A") + 10) * mult)
            Else
                If (ch >= "a") And (ch <= "f") Then
                    DecNum = DecNum + ((Asc(ch) - Asc("a") + 10) * mult)
                Else
                    HexToDec = 0
                    Exit Function
                End If
            End If
        End If
        mult = mult * 16
    Next i
    HexToDec = DecNum
End Function

Public Function DecToHex(ByVal DecNum As Double) As String
    Dim remainder As Integer
    Dim HexStr As String
    HexStr = ""
    Do While DecNum <> 0
        remainder = DecNum Mod 16   'This method Blows!!!!!! Causes Overflow
        If remainder <= 9 Then
            HexStr = Chr(Asc(remainder)) & HexStr
        Else
            HexStr = Chr(Asc("A") + remainder - 10) & HexStr
        End If
        DecNum = DecNum \ 16
    Loop
    If HexStr = "" Then HexStr = "0"
    DecToHex = HexStr
End Function

Public Function DecToBin(ByVal DecNum As Double) As String
    Dim BinStr As String
    BinStr = ""
    Do While DecNum <> 0
        If (DecNum Mod 2) = 1 Then   'This method Blows!!!!!! Causes Overflow
            BinStr = "1" & BinStr
        Else
            BinStr = "0" & BinStr
        End If
        DecNum = DecNum \ 2
    Loop
    If BinStr = "" Then BinStr = "0000"
    DecToBin = BinStr
End Function

Public Function BinToDec(ByVal BinStr As String) As Double
    Dim mult As Double
    Dim DecNum As Double
    mult = 1
    DecNum = 0
    
    Dim i As Integer
    For i = Len(BinStr) To 1 Step -1
        If Mid(BinStr, i, 1) = "1" Then
            DecNum = DecNum + mult
        End If
        mult = mult * 2
    Next i
    BinToDec = DecNum
End Function

Public Function HexToBin(ByVal HexStr As String) As String
    Dim BinStr As String
    BinStr = ""
    Dim i As Integer
    For i = 1 To Len(HexStr)
        BinStr = BinStr & DecToBin(HexToDec(Mid(HexStr, i, 1)))
    Next i
    HexToBin = BinStr
End Function

Public Function BinToHex(ByVal BinStr As String) As String
    Dim HexStr As String
    HexStr = ""
    Dim i As Integer
    For i = 1 To Len(BinStr) Step 4
        HexStr = HexStr & DecToHex(BinToDec(Mid(BinStr, i, 4)))
    Next i
    BinToHex = HexStr
End Function

Public Function ConvertNum(ByVal numToConvert As Variant, ByVal fromBase As Integer, ByVal toBase As Integer) As Variant
    If fromBase = toBase Then
        ConvertNum = numToConvert
        Exit Function
    End If
    
    Select Case fromBase
    Case 2
        Select Case toBase
        Case 10
            ConvertNum = BinToDec(numToConvert)
            Exit Function
        Case 16
            ConvertNum = BinToHex(numToConvert)
            Exit Function
        End Select
    Case 10
        Select Case toBase
        Case 2
            ConvertNum = DecToBin(numToConvert)
            Exit Function
        Case 16
            ConvertNum = DecToHex(numToConvert)
            Exit Function
        End Select
    Case 16
        Select Case toBase
        Case 2
            ConvertNum = HexToBin(numToConvert)
            Exit Function
        Case 10
            ConvertNum = HexToDec(numToConvert)
            Exit Function
        End Select
    End Select
End Function


