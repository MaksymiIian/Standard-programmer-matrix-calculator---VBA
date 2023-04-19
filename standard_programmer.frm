VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} prosty_programisty 
   Caption         =   "Kalkulator zaliczeniowy"
   ClientHeight    =   7368
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5244
   OleObjectBlob   =   "prosty_programisty.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "prosty_programisty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim operator As String
Dim firstNumber As Double
Dim secondNumber As Double
Dim result As Double
Dim f As String
Dim conversion As String
Dim s As String
Dim l As String
Dim k As String
Dim m As String
Dim a As String
Dim val As Double
Dim valS As String

Private Sub CommandButton1_Click()
If InStr(1, TextBox1.Text, ",", vbTextCompare) Then
TextBox1.Text = TextBox1.Text
Else
TextBox1.Text = TextBox1.Text + ","
End If
End Sub

Private Sub CommandButton10_Click()
If TextBox1.Text = "0" Then
TextBox1.Text = "7"
Else
TextBox1.Text = TextBox1.Text + "7"
End If

End Sub

Private Sub CommandButton11_Click()
If TextBox1.Text = "0" Then
TextBox1.Text = "8"
Else
TextBox1.Text = TextBox1.Text + "8"
End If
End Sub

Private Sub CommandButton12_Click()
If TextBox1.Text = "0" Then
TextBox1.Text = "9"
Else
TextBox1.Text = TextBox1.Text + "9"
End If
End Sub

Private Sub CommandButton13_Click()
firstNumber = TextBox1.Text

TextBox1.Text = ""
operator = "+"
End Sub

Private Sub CommandButton14_Click()
firstNumber = TextBox1.Text
TextBox1.Text = ""
operator = "-"
End Sub

Private Sub CommandButton15_Click()
firstNumber = TextBox1.Text
TextBox1.Text = ""
operator = "*"
End Sub

Private Sub CommandButton16_Click()
firstNumber = TextBox1.Text
TextBox1.Text = ""
operator = "/"
End Sub

Private Sub CommandButton17_Click()
firstNumber = TextBox1.Text
TextBox1.Text = ""
operator = "%"
End Sub

Private Sub CommandButton18_Click()
firstNumber = TextBox1.Text
TextBox1.Text = ""
operator = "pow"
result = firstNumber * firstNumber
TextBox1.Text = result
End Sub

Private Sub CommandButton19_Click()
firstNumber = TextBox1.Text
TextBox1.Text = ""
operator = "sqrt"
result = Sqr(firstNumber)
TextBox1.Text = result
End Sub

Private Sub CommandButton2_Click()
If TextBox1.Text = "0" Then
TextBox1.Text = "0"
Else
TextBox1.Text = TextBox1.Text + "0"
End If
End Sub

Private Sub CommandButton20_Click()
TextBox1.Text = "0"
End Sub

Private Sub CommandButton21_Click()
If TextBox2.Text = "0" Then
TextBox2.Text = "A"
Else
TextBox2.Text = TextBox2.Text + "A"
End If
End Sub

Private Sub CommandButton22_Click()
If TextBox2.Text = "0" Then
TextBox2.Text = "B"
Else
TextBox2.Text = TextBox2.Text + "B"
End If
End Sub

Private Sub CommandButton23_Click()
If TextBox2.Text = "0" Then
TextBox2.Text = "C"
Else
TextBox2.Text = TextBox2.Text + "C"
End If
End Sub

Private Sub CommandButton24_Click()
If TextBox2.Text = "0" Then
TextBox2.Text = "D"
Else
TextBox2.Text = TextBox2.Text + "D"
End If
End Sub

Private Sub CommandButton25_Click()
If TextBox2.Text = "0" Then
TextBox2.Text = "E"
Else
TextBox2.Text = TextBox2.Text + "E"
End If
End Sub

Private Sub CommandButton26_Click()
If TextBox2.Text = "0" Then
TextBox2.Text = "F"
Else
TextBox2.Text = TextBox2.Text + "F"
End If
End Sub
Function DecToBin(ByVal DecimalIn As Variant) As String
  DecToBin = ""
  DecimalIn = CDec(DecimalIn)
  Do While DecimalIn <> 0
    DecToBin = Trim$(Str$(DecimalIn - 2 * Int(DecimalIn / 2))) & DecToBin
    DecimalIn = Int(DecimalIn / 2)
  Loop
End Function

Function BinToDec(sMyBin As String) As Long
    Dim x As Integer
    Dim iLen As Integer
    iLen = Len(sMyBin) - 1
    For x = 0 To iLen
        BinToDec = BinToDec + _
        Mid(sMyBin, iLen - x + 1, 1) * 2 ^ x
    Next
End Function

Private Sub CommandButton27_Click()
s = TextBox2.Text
operator = "+"
If conversion = "hex" Then
        TextBox3.Text = "HEX " + TextBox2.Text
        l = WorksheetFunction.Hex2Dec(s)
        TextBox4.Text = "DEC " + l
        TextBox5.Text = "OCT " + Oct(l)
        k = DecToBin(l)
        TextBox6.Text = "BIN " + k
        TextBox2.Text = ""
    ElseIf conversion = "dec" Then
        l = TextBox2.Text
        TextBox3.Text = "HEX " + Hex(l)
        TextBox4.Text = "DEC " + TextBox2.Text
        TextBox5.Text = "OCT " + Oct(l)
        l = DecToBin(s)
        TextBox6.Text = "BIN " + l
        TextBox2.Text = ""
    ElseIf conversion = "oct" Then
        l = WorksheetFunction.Oct2Dec(s)
        TextBox3.Text = "HEX " + Hex(l)
        TextBox4.Text = "DEC " + l
        k = DecToBin(l)
        TextBox5.Text = "OCT " + TextBox2.Text
        TextBox6.Text = "BIN " + k
        TextBox2.Text = ""
    ElseIf conversion = "bin" Then
        l = BinToDec(s)
        TextBox3.Text = "HEX " + Hex(l)
        TextBox4.Text = "DEC " + l
        TextBox5.Text = "OCT " + Oct(l)
        TextBox6.Text = "BIN " + TextBox2.Text
        TextBox2.Text = ""
    ElseIf conversion = "" Then
        TextBox3.Text = "HEX " + TextBox2.Text
        l = WorksheetFunction.Hex2Dec(s)
        TextBox4.Text = "DEC " + l
        TextBox5.Text = "OCT " + Oct(l)
        k = DecToBin(l)
        TextBox6.Text = "BIN " + k
        TextBox2.Text = ""
End If
End Sub

Private Sub CommandButton28_Click()
If TextBox2.Text = "0" Then
TextBox2.Text = "0"
Else
TextBox2.Text = TextBox2.Text + "0"
End If
End Sub

Private Sub CommandButton29_Click()
s = TextBox2.Text
operator = "-"
If conversion = "hex" Then
        TextBox3.Text = "HEX " + TextBox2.Text
        l = WorksheetFunction.Hex2Dec(s)
        TextBox4.Text = "DEC " + l
        TextBox5.Text = "OCT " + Oct(l)
        k = DecToBin(l)
        TextBox6.Text = "BIN " + k
        TextBox2.Text = ""
    ElseIf conversion = "dec" Then
        l = TextBox2.Text
        TextBox3.Text = "HEX " + Hex(l)
        TextBox4.Text = "DEC " + TextBox2.Text
        TextBox5.Text = "OCT " + Oct(l)
        l = DecToBin(s)
        TextBox6.Text = "BIN " + l
        TextBox2.Text = ""
    ElseIf conversion = "oct" Then
        l = WorksheetFunction.Oct2Dec(s)
        TextBox3.Text = "HEX " + Hex(l)
        TextBox4.Text = "DEC " + l
        k = DecToBin(l)
        TextBox5.Text = "OCT " + TextBox2.Text
        TextBox6.Text = "BIN " + k
        TextBox2.Text = ""
    ElseIf conversion = "bin" Then
        l = BinToDec(s)
        TextBox3.Text = "HEX " + Hex(l)
        TextBox4.Text = "DEC " + l
        TextBox5.Text = "OCT " + Oct(l)
        TextBox6.Text = "BIN " + TextBox2.Text
        TextBox2.Text = ""
    ElseIf conversion = "" Then
        TextBox3.Text = "HEX " + TextBox2.Text
        l = WorksheetFunction.Hex2Dec(s)
        TextBox4.Text = "DEC " + l
        TextBox5.Text = "OCT " + Oct(l)
        k = DecToBin(l)
        TextBox6.Text = "BIN " + k
        TextBox2.Text = ""
End If
End Sub

Private Sub CommandButton3_Click()
secondNumber = TextBox1.Text
If operator = "+" Then
result = firstNumber + secondNumber
TextBox1.Text = result
ElseIf operator = "-" Then
result = firstNumber - secondNumber
TextBox1.Text = result
ElseIf operator = "/" Then
result = firstNumber / secondNumber
TextBox1.Text = result
ElseIf operator = "*" Then
result = firstNumber * secondNumber
TextBox1.Text = result
ElseIf operator = "%" Then
result = firstNumber Mod secondNumber
TextBox1.Text = result
ElseIf operator = "pow" Then
result = firstNumber * firstNumber
TextBox1.Text = result
ElseIf operator = "sqrt" Then
result = Sqr(firstNumber)
TextBox1.Text = result
End If
operator = ""
End Sub


Private Sub CommandButton30_Click()
If TextBox2.Text = "0" Then
TextBox2.Text = "1"
Else
TextBox2.Text = TextBox2.Text + "1"
End If
End Sub

Private Sub CommandButton31_Click()
If TextBox2.Text = "0" Then
TextBox2.Text = "2"
Else
TextBox2.Text = TextBox2.Text + "2"
End If
End Sub

Private Sub CommandButton32_Click()
If TextBox2.Text = "0" Then
TextBox2.Text = "3"
Else
TextBox2.Text = TextBox2.Text + "3"
End If
End Sub

Private Sub CommandButton33_Click()
If TextBox2.Text = "0" Then
TextBox2.Text = "4"
Else
TextBox2.Text = TextBox2.Text + "4"
End If
End Sub

Private Sub CommandButton34_Click()
If TextBox2.Text = "0" Then
TextBox2.Text = "5"
Else
TextBox2.Text = TextBox2.Text + "5"
End If
End Sub

Private Sub CommandButton35_Click()
If TextBox2.Text = "0" Then
TextBox2.Text = "6"
Else
TextBox2.Text = TextBox2.Text + "6"
End If
End Sub

Private Sub CommandButton36_Click()
If TextBox2.Text = "0" Then
TextBox2.Text = "7"
Else
TextBox2.Text = TextBox2.Text + "7"
End If
End Sub

Private Sub CommandButton37_Click()
If TextBox2.Text = "0" Then
TextBox2.Text = "8"
Else
TextBox2.Text = TextBox2.Text + "8"
End If
End Sub

Private Sub CommandButton38_Click()
If TextBox2.Text = "0" Then
TextBox2.Text = "9"
Else
TextBox2.Text = TextBox2.Text + "9"
End If
End Sub

Private Sub CommandButton39_Click()
If operator = "" Then
    s = TextBox2.Text
    If conversion = "hex" Then
        TextBox3.Text = "HEX " + TextBox2.Text
        l = WorksheetFunction.Hex2Dec(s)
        TextBox4.Text = "DEC " + l
        TextBox5.Text = "OCT " + Oct(l)
        k = DecToBin(l)
        TextBox6.Text = "BIN " + k
    
    ElseIf conversion = "dec" Then
        l = TextBox2.Text
        TextBox3.Text = "HEX " + Hex(l)
        TextBox4.Text = "DEC " + TextBox2.Text
        TextBox5.Text = "OCT " + Oct(l)
        l = DecToBin(s)
        TextBox6.Text = "BIN " + l
    ElseIf conversion = "oct" Then
        l = WorksheetFunction.Oct2Dec(s)
        TextBox3.Text = "HEX " + Hex(l)
        TextBox4.Text = "DEC " + l
        k = DecToBin(l)
        TextBox5.Text = "OCT " + TextBox2.Text
        TextBox6.Text = "BIN " + k
    ElseIf conversion = "bin" Then
        l = BinToDec(s)
        TextBox3.Text = "HEX " + Hex(l)
        TextBox4.Text = "DEC " + l
        TextBox5.Text = "OCT " + Oct(l)
        TextBox6.Text = "BIN " + TextBox2.Text
    ElseIf conversion = "" Then
        TextBox3.Text = "HEX " + TextBox2.Text
        l = WorksheetFunction.Hex2Dec(s)
        TextBox4.Text = "DEC " + l
        TextBox5.Text = "OCT " + Oct(l)
        k = DecToBin(l)
        TextBox6.Text = "BIN " + k
    End If
    
ElseIf operator = "+" Then
    f = TextBox2.Text
    If conversion = "hex" Then
        l = WorksheetFunction.Hex2Dec(s)
        m = WorksheetFunction.Hex2Dec(f)
        val = CDec(l) + CDec(m)
        valS = CStr(val)
        TextBox4.Text = "DEC " + valS
        TextBox3.Text = "HEX " + Hex(valS)
        TextBox5.Text = "OCT " + Oct(valS)
        TextBox6.Text = "BIN " + DecToBin(valS)
        TextBox2.Text = Hex(valS)
        valS = 0
    ElseIf conversion = "dec" Then
        val = CDec(s) + CDec(f)
        valS = CStr(val)
        TextBox4.Text = "DEC " + valS
        TextBox3.Text = "HEX " + Hex(valS)
        TextBox5.Text = "OCT " + Oct(valS)
        TextBox6.Text = "BIN " + DecToBin(valS)
        TextBox2.Text = valS
        valS = 0
    ElseIf conversion = "oct" Then
        l = WorksheetFunction.Oct2Dec(s)
        m = WorksheetFunction.Oct2Dec(f)
        val = CDec(l) + CDec(m)
        valS = CStr(val)
        TextBox3.Text = "HEX " + Hex(valS)
        TextBox4.Text = "DEC " + valS
        k = DecToBin(valS)
        TextBox5.Text = "OCT " + Oct(valS)
        TextBox6.Text = "BIN " + k
        TextBox2.Text = Oct(valS)
        valS = 0
    ElseIf conversion = "bin" Then
        l = BinToDec(s)
        m = BinToDec(f)
        val = CDec(l) + CDec(m)
        valS = CStr(val)
        TextBox3.Text = "HEX " + Hex(valS)
        TextBox4.Text = "DEC " + valS
        TextBox5.Text = "OCT " + Oct(valS)
        TextBox6.Text = "BIN " + DecToBin(valS)
        TextBox2.Text = DecToBin(valS)
        valS = 0
    ElseIf conversion = "" Then
        l = WorksheetFunction.Hex2Dec(s)
        m = WorksheetFunction.Hex2Dec(f)
        val = CDec(l) + CDec(m)
        valS = CStr(val)
        TextBox4.Text = "DEC " + valS
        TextBox3.Text = "HEX " + Hex(valS)
        TextBox5.Text = "OCT " + Oct(valS)
        TextBox6.Text = "BIN " + DecToBin(valS)
        TextBox2.Text = Hex(valS)
        valS = 0
    End If
ElseIf operator = "-" Then
    f = TextBox2.Text
    If conversion = "hex" Then
        l = WorksheetFunction.Hex2Dec(s)
        m = WorksheetFunction.Hex2Dec(f)
        val = CDec(l) - CDec(m)
        valS = CStr(val)
        TextBox4.Text = "DEC " + valS
        TextBox3.Text = "HEX " + Hex(valS)
        TextBox5.Text = "OCT " + Oct(valS)
        TextBox6.Text = "BIN " + DecToBin(valS)
        TextBox2.Text = Hex(valS)
        valS = 0
    ElseIf conversion = "dec" Then
        val = CDec(s) - CDec(f)
        valS = CStr(val)
        TextBox4.Text = "DEC " + valS
        TextBox3.Text = "HEX " + Hex(valS)
        TextBox5.Text = "OCT " + Oct(valS)
        TextBox6.Text = "BIN " + DecToBin(valS)
        TextBox2.Text = valS
        valS = 0
    ElseIf conversion = "oct" Then
        l = WorksheetFunction.Oct2Dec(s)
        m = WorksheetFunction.Oct2Dec(f)
        val = CDec(l) - CDec(m)
        valS = CStr(val)
        TextBox3.Text = "HEX " + Hex(valS)
        TextBox4.Text = "DEC " + valS
        k = DecToBin(valS)
        TextBox5.Text = "OCT " + Oct(valS)
        TextBox6.Text = "BIN " + k
        TextBox2.Text = Oct(valS)
        valS = 0
    ElseIf conversion = "bin" Then
        l = BinToDec(s)
        m = BinToDec(f)
        val = CDec(l) - CDec(m)
        valS = CStr(val)
        TextBox3.Text = "HEX " + Hex(valS)
        TextBox4.Text = "DEC " + valS
        TextBox5.Text = "OCT " + Oct(valS)
        TextBox6.Text = "BIN " + DecToBin(valS)
        TextBox2.Text = DecToBin(valS)
        valS = 0
    ElseIf conversion = "" Then
        l = WorksheetFunction.Hex2Dec(s)
        m = WorksheetFunction.Hex2Dec(f)
        val = CDec(l) - CDec(m)
        valS = CStr(val)
        TextBox4.Text = "DEC " + valS
        TextBox3.Text = "HEX " + Hex(valS)
        TextBox5.Text = "OCT " + Oct(valS)
        TextBox6.Text = "BIN " + DecToBin(valS)
        TextBox2.Text = Hex(valS)
        valS = 0
    End If
ElseIf operator = "*" Then
    f = TextBox2.Text
    If conversion = "hex" Then
        l = WorksheetFunction.Hex2Dec(s)
        m = WorksheetFunction.Hex2Dec(f)
        val = CDec(l) * CDec(m)
        valS = CStr(val)
        TextBox4.Text = "DEC " + valS
        TextBox3.Text = "HEX " + Hex(valS)
        TextBox5.Text = "OCT " + Oct(valS)
        TextBox6.Text = "BIN " + DecToBin(valS)
        TextBox2.Text = Hex(valS)
        valS = 0
    ElseIf conversion = "dec" Then
        val = CDec(s) * CDec(f)
        valS = CStr(val)
        TextBox4.Text = "DEC " + valS
        TextBox3.Text = "HEX " + Hex(valS)
        TextBox5.Text = "OCT " + Oct(valS)
        TextBox6.Text = "BIN " + DecToBin(valS)
        TextBox2.Text = valS
        valS = 0
    ElseIf conversion = "oct" Then
        l = WorksheetFunction.Oct2Dec(s)
        m = WorksheetFunction.Oct2Dec(f)
        val = CDec(l) * CDec(m)
        valS = CStr(val)
        TextBox3.Text = "HEX " + Hex(valS)
        TextBox4.Text = "DEC " + valS
        k = DecToBin(valS)
        TextBox5.Text = "OCT " + Oct(valS)
        TextBox6.Text = "BIN " + k
        TextBox2.Text = Oct(valS)
        valS = 0
    ElseIf conversion = "bin" Then
        l = BinToDec(s)
        m = BinToDec(f)
        val = CDec(l) * CDec(m)
        valS = CStr(val)
        TextBox3.Text = "HEX " + Hex(valS)
        TextBox4.Text = "DEC " + valS
        TextBox5.Text = "OCT " + Oct(valS)
        TextBox6.Text = "BIN " + DecToBin(valS)
        TextBox2.Text = DecToBin(valS)
        valS = 0
    ElseIf conversion = "" Then
        l = WorksheetFunction.Hex2Dec(s)
        m = WorksheetFunction.Hex2Dec(f)
        val = CDec(l) * CDec(m)
        valS = CStr(val)
        TextBox4.Text = "DEC " + valS
        TextBox3.Text = "HEX " + Hex(valS)
        TextBox5.Text = "OCT " + Oct(valS)
        TextBox6.Text = "BIN " + DecToBin(valS)
        TextBox2.Text = Hex(valS)
        valS = 0
    End If
ElseIf operator = "/" Then
    f = TextBox2.Text
    If conversion = "hex" Then
        l = WorksheetFunction.Hex2Dec(s)
        m = WorksheetFunction.Hex2Dec(f)
        val = CDec(l) / CDec(m)
        valS = Application.WorksheetFunction.RoundDown(CStr(val), 0)
        TextBox4.Text = "DEC " + valS
        TextBox3.Text = "HEX " + Hex(valS)
        TextBox5.Text = "OCT " + Oct(valS)
        TextBox6.Text = "BIN " + DecToBin(valS)
        TextBox2.Text = Hex(valS)
        valS = 0
    ElseIf conversion = "dec" Then
        val = CDec(s) / CDec(f)
        valS = Application.WorksheetFunction.RoundDown(CStr(val), 0)
        TextBox4.Text = "DEC " + valS
        TextBox3.Text = "HEX " + Hex(valS)
        TextBox5.Text = "OCT " + Oct(valS)
        TextBox6.Text = "BIN " + DecToBin(valS)
        TextBox2.Text = valS
        valS = 0
    ElseIf conversion = "oct" Then
        l = WorksheetFunction.Oct2Dec(s)
        m = WorksheetFunction.Oct2Dec(f)
        val = CDec(l) / CDec(m)
        valS = Application.WorksheetFunction.RoundDown(CStr(val), 0)
        TextBox3.Text = "HEX " + Hex(valS)
        TextBox4.Text = "DEC " + valS
        k = DecToBin(valS)
        TextBox5.Text = "OCT " + Oct(valS)
        TextBox6.Text = "BIN " + k
        TextBox2.Text = Oct(valS)
        valS = 0
    ElseIf conversion = "bin" Then
        l = BinToDec(s)
        m = BinToDec(f)
        val = CDec(l) / CDec(m)
        valS = Application.WorksheetFunction.RoundDown(CStr(val), 0)
        TextBox3.Text = "HEX " + Hex(valS)
        TextBox4.Text = "DEC " + valS
        TextBox5.Text = "OCT " + Oct(valS)
        TextBox6.Text = "BIN " + DecToBin(valS)
        TextBox2.Text = DecToBin(valS)
        valS = 0
    ElseIf conversion = "" Then
        l = WorksheetFunction.Hex2Dec(s)
        m = WorksheetFunction.Hex2Dec(f)
        val = Application.WorksheetFunction.RoundDown(CStr(val), 0)
        valS = CStr(val)
        TextBox4.Text = "DEC " + valS
        TextBox3.Text = "HEX " + Hex(valS)
        TextBox5.Text = "OCT " + Oct(valS)
        TextBox6.Text = "BIN " + DecToBin(valS)
        TextBox2.Text = Hex(valS)
        valS = 0
    End If
ElseIf operator = "%" Then
    f = TextBox2.Text
    If conversion = "hex" Then
        l = WorksheetFunction.Hex2Dec(s)
        m = WorksheetFunction.Hex2Dec(f)
        val = CDec(l) Mod CDec(m)
        valS = CStr(val)
        TextBox4.Text = "DEC " + valS
        TextBox3.Text = "HEX " + Hex(valS)
        TextBox5.Text = "OCT " + Oct(valS)
        TextBox6.Text = "BIN " + DecToBin(valS)
        TextBox2.Text = Hex(valS)
        valS = 0
    ElseIf conversion = "dec" Then
        val = CDec(s) Mod CDec(f)
        valS = CStr(val)
        TextBox4.Text = "DEC " + valS
        TextBox3.Text = "HEX " + Hex(valS)
        TextBox5.Text = "OCT " + Oct(valS)
        TextBox6.Text = "BIN " + DecToBin(valS)
        TextBox2.Text = valS
        valS = 0
    ElseIf conversion = "oct" Then
        l = WorksheetFunction.Oct2Dec(s)
        m = WorksheetFunction.Oct2Dec(f)
        val = CDec(l) Mod CDec(m)
        valS = CStr(val)
        TextBox3.Text = "HEX " + Hex(valS)
        TextBox4.Text = "DEC " + valS
        k = DecToBin(valS)
        TextBox5.Text = "OCT " + Oct(valS)
        TextBox6.Text = "BIN " + k
        TextBox2.Text = Oct(valS)
        valS = 0
    ElseIf conversion = "bin" Then
        l = BinToDec(s)
        m = BinToDec(f)
        val = CDec(l) Mod CDec(m)
        valS = CStr(val)
        TextBox3.Text = "HEX " + Hex(valS)
        TextBox4.Text = "DEC " + valS
        TextBox5.Text = "OCT " + Oct(valS)
        TextBox6.Text = "BIN " + DecToBin(valS)
        TextBox2.Text = DecToBin(valS)
        valS = 0
    ElseIf conversion = "" Then
        l = WorksheetFunction.Hex2Dec(s)
        m = WorksheetFunction.Hex2Dec(f)
        val = CDec(l) Mod CDec(m)
        valS = CStr(val)
        TextBox4.Text = "DEC " + valS
        TextBox3.Text = "HEX " + Hex(valS)
        TextBox5.Text = "OCT " + Oct(valS)
        TextBox6.Text = "BIN " + DecToBin(valS)
        TextBox2.Text = Hex(valS)
        valS = 0
    End If
Else
    If conversion = "hex" Then
        TextBox3.Text = "HEX " + TextBox2.Text
        l = WorksheetFunction.Hex2Dec(s)
        TextBox4.Text = "DEC " + l
        TextBox5.Text = "OCT " + Oct(l)
        k = DecToBin(l)
        TextBox6.Text = "BIN " + k
    
    ElseIf conversion = "dec" Then
        l = TextBox2.Text
        TextBox3.Text = "HEX " + Hex(l)
        TextBox4.Text = "DEC " + TextBox2.Text
        TextBox5.Text = "OCT " + Oct(l)
        l = DecToBin(s)
        TextBox6.Text = "BIN " + l
    ElseIf conversion = "oct" Then
        l = WorksheetFunction.Oct2Dec(s)
        TextBox3.Text = "HEX " + Hex(l)
        TextBox4.Text = "DEC " + l
        k = DecToBin(l)
        TextBox5.Text = "OCT " + TextBox2.Text
        TextBox6.Text = "BIN " + k
    ElseIf conversion = "bin" Then
        l = BinToDec(s)
        TextBox3.Text = "HEX " + Hex(l)
        TextBox4.Text = "DEC " + l
        TextBox5.Text = "OCT " + Oct(l)
        TextBox6.Text = "BIN " + TextBox2.Text
    ElseIf conversion = "" Then
        TextBox3.Text = "HEX " + TextBox2.Text
        l = WorksheetFunction.Hex2Dec(s)
        TextBox4.Text = "DEC " + l
        TextBox5.Text = "OCT " + Oct(l)
        k = DecToBin(l)
        TextBox6.Text = "BIN " + k
    End If
End If
operator = ""
End Sub

Private Sub CommandButton4_Click()
If TextBox1.Text = "0" Then
TextBox1.Text = "1"
Else
TextBox1.Text = TextBox1.Text + "1"
End If
End Sub

Private Sub CommandButton40_Click()
s = TextBox2.Text
operator = "*"
If conversion = "hex" Then
        TextBox3.Text = "HEX " + TextBox2.Text
        l = WorksheetFunction.Hex2Dec(s)
        TextBox4.Text = "DEC " + l
        TextBox5.Text = "OCT " + Oct(l)
        k = DecToBin(l)
        TextBox6.Text = "BIN " + k
        TextBox2.Text = ""
    ElseIf conversion = "dec" Then
        l = TextBox2.Text
        TextBox3.Text = "HEX " + Hex(l)
        TextBox4.Text = "DEC " + TextBox2.Text
        TextBox5.Text = "OCT " + Oct(l)
        l = DecToBin(s)
        TextBox6.Text = "BIN " + l
        TextBox2.Text = ""
    ElseIf conversion = "oct" Then
        l = WorksheetFunction.Oct2Dec(s)
        TextBox3.Text = "HEX " + Hex(l)
        TextBox4.Text = "DEC " + l
        k = DecToBin(l)
        TextBox5.Text = "OCT " + TextBox2.Text
        TextBox6.Text = "BIN " + k
        TextBox2.Text = ""
    ElseIf conversion = "bin" Then
        l = BinToDec(s)
        TextBox3.Text = "HEX " + Hex(l)
        TextBox4.Text = "DEC " + l
        TextBox5.Text = "OCT " + Oct(l)
        TextBox6.Text = "BIN " + TextBox2.Text
        TextBox2.Text = ""
    ElseIf conversion = "" Then
        TextBox3.Text = "HEX " + TextBox2.Text
        l = WorksheetFunction.Hex2Dec(s)
        TextBox4.Text = "DEC " + l
        TextBox5.Text = "OCT " + Oct(l)
        k = DecToBin(l)
        TextBox6.Text = "BIN " + k
        TextBox2.Text = ""
End If
End Sub

Private Sub CommandButton41_Click()
s = TextBox2.Text
operator = "/"
If conversion = "hex" Then
        TextBox3.Text = "HEX " + TextBox2.Text
        l = WorksheetFunction.Hex2Dec(s)
        TextBox4.Text = "DEC " + l
        TextBox5.Text = "OCT " + Oct(l)
        k = DecToBin(l)
        TextBox6.Text = "BIN " + k
        TextBox2.Text = ""
    ElseIf conversion = "dec" Then
        l = TextBox2.Text
        TextBox3.Text = "HEX " + Hex(l)
        TextBox4.Text = "DEC " + TextBox2.Text
        TextBox5.Text = "OCT " + Oct(l)
        l = DecToBin(s)
        TextBox6.Text = "BIN " + l
        TextBox2.Text = ""
    ElseIf conversion = "oct" Then
        l = WorksheetFunction.Oct2Dec(s)
        TextBox3.Text = "HEX " + Hex(l)
        TextBox4.Text = "DEC " + l
        k = DecToBin(l)
        TextBox5.Text = "OCT " + TextBox2.Text
        TextBox6.Text = "BIN " + k
        TextBox2.Text = ""
    ElseIf conversion = "bin" Then
        l = BinToDec(s)
        TextBox3.Text = "HEX " + Hex(l)
        TextBox4.Text = "DEC " + l
        TextBox5.Text = "OCT " + Oct(l)
        TextBox6.Text = "BIN " + TextBox2.Text
        TextBox2.Text = ""
    ElseIf conversion = "" Then
        TextBox3.Text = "HEX " + TextBox2.Text
        l = WorksheetFunction.Hex2Dec(s)
        TextBox4.Text = "DEC " + l
        TextBox5.Text = "OCT " + Oct(l)
        k = DecToBin(l)
        TextBox6.Text = "BIN " + k
        TextBox2.Text = ""
End If
End Sub

Private Sub CommandButton42_Click()
s = TextBox2.Text
operator = "%"
If conversion = "hex" Then
        TextBox3.Text = "HEX " + TextBox2.Text
        l = WorksheetFunction.Hex2Dec(s)
        TextBox4.Text = "DEC " + l
        TextBox5.Text = "OCT " + Oct(l)
        k = DecToBin(l)
        TextBox6.Text = "BIN " + k
        TextBox2.Text = ""
    ElseIf conversion = "dec" Then
        l = TextBox2.Text
        TextBox3.Text = "HEX " + Hex(l)
        TextBox4.Text = "DEC " + TextBox2.Text
        TextBox5.Text = "OCT " + Oct(l)
        l = DecToBin(s)
        TextBox6.Text = "BIN " + l
        TextBox2.Text = ""
    ElseIf conversion = "oct" Then
        l = WorksheetFunction.Oct2Dec(s)
        TextBox3.Text = "HEX " + Hex(l)
        TextBox4.Text = "DEC " + l
        k = DecToBin(l)
        TextBox5.Text = "OCT " + TextBox2.Text
        TextBox6.Text = "BIN " + k
        TextBox2.Text = ""
    ElseIf conversion = "bin" Then
        l = BinToDec(s)
        TextBox3.Text = "HEX " + Hex(l)
        TextBox4.Text = "DEC " + l
        TextBox5.Text = "OCT " + Oct(l)
        TextBox6.Text = "BIN " + TextBox2.Text
        TextBox2.Text = ""
    ElseIf conversion = "" Then
        TextBox3.Text = "HEX " + TextBox2.Text
        l = WorksheetFunction.Hex2Dec(s)
        TextBox4.Text = "DEC " + l
        TextBox5.Text = "OCT " + Oct(l)
        k = DecToBin(l)
        TextBox6.Text = "BIN " + k
        TextBox2.Text = ""
End If
End Sub

Private Sub CommandButton43_Click()
TextBox2.Text = "0"
TextBox3.Text = "HEX 0"
TextBox4.Text = "DEC 0"
TextBox5.Text = "OCT 0"
TextBox6.Text = "BIN 0"
conversion = "oct"
CommandButton31.Enabled = True
CommandButton32.Enabled = True
CommandButton33.Enabled = True
CommandButton34.Enabled = True
CommandButton35.Enabled = True
CommandButton36.Enabled = True
CommandButton21.Enabled = False
CommandButton22.Enabled = False
CommandButton23.Enabled = False
CommandButton24.Enabled = False
CommandButton25.Enabled = False
CommandButton26.Enabled = False
CommandButton37.Enabled = False
CommandButton38.Enabled = False

End Sub

Private Sub CommandButton44_Click()
TextBox2.Text = "0"
TextBox3.Text = "HEX 0"
TextBox4.Text = "DEC 0"
TextBox5.Text = "OCT 0"
TextBox6.Text = "BIN 0"
conversion = "hex"
CommandButton21.Enabled = True
CommandButton22.Enabled = True
CommandButton23.Enabled = True
CommandButton24.Enabled = True
CommandButton25.Enabled = True
CommandButton26.Enabled = True
CommandButton31.Enabled = True
CommandButton32.Enabled = True
CommandButton33.Enabled = True
CommandButton34.Enabled = True
CommandButton35.Enabled = True
CommandButton36.Enabled = True
CommandButton37.Enabled = True
CommandButton38.Enabled = True
End Sub

Private Sub CommandButton45_Click()
TextBox2.Text = "0"
TextBox3.Text = "HEX 0"
TextBox4.Text = "DEC 0"
TextBox5.Text = "OCT 0"
TextBox6.Text = "BIN 0"
conversion = "bin"
CommandButton21.Enabled = False
CommandButton22.Enabled = False
CommandButton23.Enabled = False
CommandButton24.Enabled = False
CommandButton25.Enabled = False
CommandButton26.Enabled = False
CommandButton31.Enabled = False
CommandButton32.Enabled = False
CommandButton33.Enabled = False
CommandButton34.Enabled = False
CommandButton35.Enabled = False
CommandButton36.Enabled = False
CommandButton37.Enabled = False
CommandButton38.Enabled = False
End Sub

Private Sub CommandButton46_Click()
TextBox2.Text = "0"
TextBox3.Text = "HEX 0"
TextBox4.Text = "DEC 0"
TextBox5.Text = "OCT 0"
TextBox6.Text = "BIN 0"
conversion = "dec"
CommandButton31.Enabled = True
CommandButton32.Enabled = True
CommandButton33.Enabled = True
CommandButton34.Enabled = True
CommandButton35.Enabled = True
CommandButton36.Enabled = True
CommandButton37.Enabled = True
CommandButton38.Enabled = True
CommandButton21.Enabled = False
CommandButton22.Enabled = False
CommandButton23.Enabled = False
CommandButton24.Enabled = False
CommandButton25.Enabled = False
CommandButton26.Enabled = False

End Sub

Private Sub CommandButton48_Click()
TextBox2.Text = "0"
TextBox3.Text = "HEX 0"
TextBox4.Text = "DEC 0"
TextBox5.Text = "OCT 0"
TextBox6.Text = "BIN 0"
End Sub

Private Sub CommandButton5_Click()
If TextBox1.Text = "0" Then
TextBox1.Text = "2"
Else
TextBox1.Text = TextBox1.Text + "2"
End If
End Sub

Private Sub CommandButton6_Click()
If TextBox1.Text = "0" Then
TextBox1.Text = "3"
Else
TextBox1.Text = TextBox1.Text + "3"
End If
End Sub

Private Sub CommandButton7_Click()
If TextBox1.Text = "0" Then
TextBox1.Text = "4"
Else
TextBox1.Text = TextBox1.Text + "4"
End If
End Sub

Private Sub CommandButton8_Click()
If TextBox1.Text = "0" Then
TextBox1.Text = "5"
Else
TextBox1.Text = TextBox1.Text + "5"
End If
End Sub

Private Sub CommandButton9_Click()
If TextBox1.Text = "0" Then
TextBox1.Text = "6"
Else
TextBox1.Text = TextBox1.Text + "6"
End If
End Sub


Private Sub TextBox2_Change()

End Sub

Private Sub UserForm_Click()

End Sub

