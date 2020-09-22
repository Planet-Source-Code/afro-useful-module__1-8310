Attribute VB_Name = "Module"
Option Explicit
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Const EWX_FORCE = 4
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1
Public Declare Function FindWindow% Lib "user32" Alias "FindWindowA" (ByVal lpClassName$, ByVal lpWindowName$)
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_SCREENSAVERRUNNING = 97


Public Sub AllCAPS(CAPS As Object)
CAPS = Format(CAPS, ">")
'For Example: "HeLLo It'S Me" would become "HELLO IT'S ME"
End Sub
Public Sub lowercase(lcase As Object)
lcase = Format(lcase, "<")
'For Example: "HeLLo It'S Me" would become "hello it's me"
End Sub
Public Sub DateTime(DT As Label)
DT.Caption = Now ' The "Now" command places the date and current time into an object's caption.
End Sub

Public Sub SaveListBox(Path As String, List As ListBox)
'Ex: Call SaveListBox("c:\windows\desktop\list.lst", list1)

    Dim Listz As Long
    On Error Resume Next

    Open Path$ For Output As #1
    For Listz& = 0 To List.ListCount - 1
        Print #1, List.List(Listz&)
        Next Listz&
    Close #1
End Sub
Public Function LstExtract(LstBox As ListBox, Txtbox As textbox)
' This takes everything from a ListBox, puts it into a TextBox and separates it with a comma (",")
Dim a As Long
Dim b As String
    
        For a = 1 To LstBox.ListCount - LstBox.ListCount
            LstBox.AddItem ", " & a
        Next

        For a = 0 To (LstBox.ListCount - 1)
            b = b & LstBox.List(a) & ", "
            
    Next


        Txtbox.Text = Mid(b, 1, Len(b) - 2)
        
End Function
Public Sub LoadComboBox(Path As String, Combo As ComboBox)
'For Example: LoadComboBox("c:\MyCombo.txt", Combo1)

    Dim What As String
    On Error Resume Next
    Open Path$ For Input As #1
    While Not EOF(1)
        Input #1, What$
        DoEvents
        Combo.AddItem What$
    Wend
    Close #1
End Sub
Public Sub LoadText(Txt As textbox, FilePath As String)
'For Example: LoadText(List1,"c:\MyText.txt")

    Dim mystr As String, FilePath2 As String, textz As String, a As String
    
    Open FilePath2$ For Input As #1
    Do While Not EOF(1)
    Line Input #1, a$
        textz$ = textz$ + a$ + Chr$(13) + Chr$(10)
        Loop
        Txt = textz$
    Close #1
End Sub
Public Sub LoadListBox(Path As String, Lst As ListBox)
'For Example: LoadListBox("c:\MyList.txt", List1)

    Dim What As String
    On Error Resume Next

    Open Path$ For Input As #1
    While Not EOF(1)
        Input #1, What$
        DoEvents
        Lst.AddItem What$
    Wend
    Close #1
End Sub
Public Function ListBoxSearch(Search As String, LB As ListBox)
'This is possibly the fastest ListBox search there is
Call SendMessageByString(LB.hwnd, LB_SELECTSTRING, 0&, Search$)
End Function
Public Sub SaveTextBox(Txt As textbox, FilePath As String)
'For Example: SaveTextBox(Text1,"c:\MyText.txt")
    Dim FilePath3 As String
    
    Open FilePath3$ For Output As #1
        Print #1, Txt
    Close 1
End Sub

Public Sub SaveComboBox(Path As String, Combo As ComboBox)
'For Example: SaveComboBox("c:\MyCombo.txt", Combo1)

    Dim Saves As Long
    On Error Resume Next

    Open Path$ For Output As #1
    For Saves& = 0 To Combo.ListCount - 1
        Print #1, Combo.List(Saves&)
    Next Saves&
    Close #1
End Sub

Public Sub Average(Total As Integer, MySum As Integer, Answer As textbox)
Dim MyAnswer
MyAnswer = Total / MySum
Answer = MyAnswer
End Sub


Public Sub Cm2Inches(Cms As Integer, Inches As textbox)
Inches.Text = Cms / 2.54
End Sub

Public Sub Inches2Cm(Inches As Integer, Cms As textbox)
Cms.Text = Inches * 2.54
End Sub

Public Sub PokemonAnagrams(Text1 As textbox, Text2 As textbox)
'Call me sad but I had to sit there and write out over 150 Pokemon!
Randomize
Dim MyValue
MyValue = Int((153 * Rnd) + 1)

If MyValue = 1 Then
    Text1.Text = "bulbasaur"
ElseIf MyValue = 2 Then
    Text1.Text = "ivysaur"
ElseIf MyValue = 3 Then
    Text1.Text = "venusaur"
ElseIf MyValue = 4 Then
    Text1.Text = "charmander"
ElseIf MyValue = 5 Then
    Text1.Text = "charmeleon"
ElseIf MyValue = 6 Then
    Text1.Text = "charizard"
ElseIf MyValue = 7 Then
    Text1.Text = "squirtle"
ElseIf MyValue = 8 Then
    Text1.Text = "wartortle"
ElseIf MyValue = 9 Then
    Text1.Text = "blastoise"
ElseIf MyValue = 10 Then
    Text1.Text = "caterpie"
ElseIf MyValue = 11 Then
    Text1.Text = "metapod"
ElseIf MyValue = 12 Then
    Text1.Text = "butterfree"
ElseIf MyValue = 13 Then
    Text1.Text = "weedle"
ElseIf MyValue = 14 Then
    Text1.Text = "kakuna"
ElseIf MyValue = 15 Then
    Text1.Text = "beedrill"
ElseIf MyValue = 16 Then
    Text1.Text = "pidgey"
ElseIf MyValue = 17 Then
    Text1.Text = "pidgeotto"
ElseIf MyValue = 18 Then
    Text1.Text = "pidgeot"
ElseIf MyValue = 19 Then
    Text1.Text = "rattata"
ElseIf MyValue = 20 Then
    Text1.Text = "raticate"
ElseIf MyValue = 21 Then
    Text1.Text = "spearow"
ElseIf MyValue = 22 Then
    Text1.Text = "fearow"
ElseIf MyValue = 23 Then
    Text1.Text = "ekans"
ElseIf MyValue = 24 Then
    Text1.Text = "arbok"
ElseIf MyValue = 25 Then
    Text1.Text = "pikachu"
ElseIf MyValue = 26 Then
    Text1.Text = "raichu"
ElseIf MyValue = 27 Then
    Text1.Text = "sandshrew"
ElseIf MyValue = 28 Then
    Text1.Text = "sandslash"
ElseIf MyValue = 29 Then
    Text1.Text = "nidoran"
ElseIf MyValue = 30 Then
    Text1.Text = "nidorina"
ElseIf MyValue = 31 Then
    Text1.Text = "nidoqueen"
ElseIf MyValue = 32 Then
    Text1.Text = "nidoran"
ElseIf MyValue = 33 Then
    Text1.Text = "nidorino"
ElseIf MyValue = 34 Then
    Text1.Text = "nidoking"
ElseIf MyValue = 35 Then
    Text1.Text = "clefairy"
ElseIf MyValue = 36 Then
    Text1.Text = "clefable"
ElseIf MyValue = 37 Then
    Text1.Text = "vulpix"
ElseIf MyValue = 38 Then
    Text1.Text = "ninetales"
ElseIf MyValue = 39 Then
    Text1.Text = "jigglypuff"
ElseIf MyValue = 40 Then
    Text1.Text = "wigglytuff"
ElseIf MyValue = 41 Then
    Text1.Text = "zubat"
ElseIf MyValue = 42 Then
    Text1.Text = "golbat"
ElseIf MyValue = 43 Then
    Text1.Text = "oddish"
ElseIf MyValue = 44 Then
    Text1.Text = "gloom"
ElseIf MyValue = 45 Then
    Text1.Text = "vileplume"
ElseIf MyValue = 46 Then
    Text1.Text = "paras"
ElseIf MyValue = 47 Then
    Text1.Text = "parasect"
ElseIf MyValue = 48 Then
    Text1.Text = "venonat"
ElseIf MyValue = 49 Then
    Text1.Text = "venomoth"
ElseIf MyValue = 50 Then
    Text1.Text = "diglett"
ElseIf MyValue = 51 Then
    Text1.Text = "dugtrio"
ElseIf MyValue = 52 Then
    Text1.Text = "meowth"
ElseIf MyValue = 53 Then
    Text1.Text = "persian"
ElseIf MyValue = 54 Then
    Text1.Text = "psyduck"
ElseIf MyValue = 55 Then
    Text1.Text = "golduck"
ElseIf MyValue = 56 Then
    Text1.Text = "mankey"
ElseIf MyValue = 57 Then
    Text1.Text = "primeape"
ElseIf MyValue = 58 Then
    Text1.Text = "growlithe"
ElseIf MyValue = 59 Then
    Text1.Text = "arcanine"
ElseIf MyValue = 60 Then
    Text1.Text = "poliwag"
ElseIf MyValue = 61 Then
    Text1.Text = "poliwhirl"
ElseIf MyValue = 62 Then
    Text1.Text = "poliwrath"
ElseIf MyValue = 63 Then
    Text1.Text = "abra"
ElseIf MyValue = 64 Then
    Text1.Text = "kadabra"
ElseIf MyValue = 65 Then
    Text1.Text = "alakazam"
ElseIf MyValue = 66 Then
    Text1.Text = "machop"
ElseIf MyValue = 67 Then
    Text1.Text = "machoke"
ElseIf MyValue = 68 Then
    Text1.Text = "machamp"
ElseIf MyValue = 69 Then
    Text1.Text = "bellsprout"
ElseIf MyValue = 70 Then
    Text1.Text = "weepinbell"
ElseIf MyValue = 71 Then
    Text1.Text = "victreebel"
ElseIf MyValue = 72 Then
    Text1.Text = "tentacool"
ElseIf MyValue = 73 Then
    Text1.Text = "tentacruel"
ElseIf MyValue = 74 Then
    Text1.Text = "geodude"
ElseIf MyValue = 75 Then
    Text1.Text = "graveler"
ElseIf MyValue = 76 Then
    Text1.Text = "golem"
ElseIf MyValue = 77 Then
    Text1.Text = "ponyta"
ElseIf MyValue = 78 Then
    Text1.Text = "rapidash"
ElseIf MyValue = 79 Then
    Text1.Text = "slowpoke"
ElseIf MyValue = 80 Then
    Text1.Text = "slowbro"
ElseIf MyValue = 81 Then
    Text1.Text = "magnemite"
ElseIf MyValue = 82 Then
    Text1.Text = "magneton"
ElseIf MyValue = 83 Then
    Text1.Text = "farfetch'd"
ElseIf MyValue = 84 Then
    Text1.Text = "doduo"
ElseIf MyValue = 85 Then
    Text1.Text = "dodrio"
ElseIf MyValue = 86 Then
    Text1.Text = "seel"
ElseIf MyValue = 87 Then
    Text1.Text = "dewgong"
ElseIf MyValue = 88 Then
    Text1.Text = "grimer"
ElseIf MyValue = 89 Then
    Text1.Text = "muk"
ElseIf MyValue = 90 Then
    Text1.Text = "shellder"
ElseIf MyValue = 91 Then
    Text1.Text = "cloyster"
ElseIf MyValue = 92 Then
    Text1.Text = "gastly"
ElseIf MyValue = 93 Then
    Text1.Text = "haunter"
ElseIf MyValue = 94 Then
    Text1.Text = "gengar"
ElseIf MyValue = 95 Then
    Text1.Text = "onix"
ElseIf MyValue = 96 Then
    Text1.Text = "drowzee"
ElseIf MyValue = 97 Then
    Text1.Text = "hypno"
ElseIf MyValue = 98 Then
    Text1.Text = "krabby"
ElseIf MyValue = 99 Then
    Text1.Text = "kingler"
ElseIf MyValue = 100 Then
    Text1.Text = "voltorb"
ElseIf MyValue = 101 Then
    Text1.Text = "electrode"
ElseIf MyValue = 102 Then
    Text1.Text = "exeggcute"
ElseIf MyValue = 103 Then
    Text1.Text = "exeggutor"
ElseIf MyValue = 104 Then
    Text1.Text = "kubone"
ElseIf MyValue = 105 Then
    Text1.Text = "marowak"
ElseIf MyValue = 106 Then
    Text1.Text = "hitmonlee"
ElseIf MyValue = 107 Then
    Text1.Text = "hitmonchan"
ElseIf MyValue = 108 Then
    Text1.Text = "lickitung"
ElseIf MyValue = 109 Then
    Text1.Text = "koffing"
ElseIf MyValue = 110 Then
    Text1.Text = "weezing"
ElseIf MyValue = 111 Then
    Text1.Text = "rhyhorn"
ElseIf MyValue = 112 Then
    Text1.Text = "rhydon"
ElseIf MyValue = 113 Then
    Text1.Text = "chansey"
ElseIf MyValue = 114 Then
    Text1.Text = "tangela"
ElseIf MyValue = 115 Then
    Text1.Text = "kangaskhan"
ElseIf MyValue = 116 Then
    Text1.Text = "horsea"
ElseIf MyValue = 117 Then
    Text1.Text = "seadra"
ElseIf MyValue = 118 Then
    Text1.Text = "goldeen"
ElseIf MyValue = 119 Then
    Text1.Text = "seaking"
ElseIf MyValue = 120 Then
    Text1.Text = "staryu"
ElseIf MyValue = 121 Then
    Text1.Text = "starmie"
ElseIf MyValue = 122 Then
    Text1.Text = "mr mime"
ElseIf MyValue = 123 Then
    Text1.Text = "scyther"
ElseIf MyValue = 124 Then
    Text1.Text = "jynx"
ElseIf MyValue = 125 Then
    Text1.Text = "electabuzz"
ElseIf MyValue = 126 Then
    Text1.Text = "magmar"
ElseIf MyValue = 127 Then
    Text1.Text = "pinsir"
ElseIf MyValue = 128 Then
    Text1.Text = "tauros"
ElseIf MyValue = 129 Then
    Text1.Text = "magikarp"
ElseIf MyValue = 130 Then
    Text1.Text = "gyrados"
ElseIf MyValue = 131 Then
    Text1.Text = "lapras"
ElseIf MyValue = 132 Then
    Text1.Text = "ditto"
ElseIf MyValue = 133 Then
    Text1.Text = "eevee"
ElseIf MyValue = 134 Then
    Text1.Text = "vaporeon"
ElseIf MyValue = 135 Then
    Text1.Text = "jolteon"
ElseIf MyValue = 136 Then
    Text1.Text = "flareon"
ElseIf MyValue = 137 Then
    Text1.Text = "porygon"
ElseIf MyValue = 138 Then
    Text1.Text = "omanyte"
ElseIf MyValue = 139 Then
    Text1.Text = "omastar"
ElseIf MyValue = 140 Then
    Text1.Text = "kabuto"
ElseIf MyValue = 141 Then
    Text1.Text = "kabutops"
ElseIf MyValue = 142 Then
    Text1.Text = "aerodactyl"
ElseIf MyValue = 143 Then
    Text1.Text = "snorlax"
ElseIf MyValue = 144 Then
    Text1.Text = "articuno"
ElseIf MyValue = 145 Then
    Text1.Text = "zapdos"
ElseIf MyValue = 146 Then
    Text1.Text = "moltres"
ElseIf MyValue = 147 Then
    Text1.Text = "dratini"
ElseIf MyValue = 148 Then
    Text1.Text = "dragonair"
ElseIf MyValue = 149 Then
    Text1.Text = "dragonite"
ElseIf MyValue = 150 Then
    Text1.Text = "mewtwo"
ElseIf MyValue = 151 Then
    Text1.Text = "mew"
ElseIf MyValue = 152 Then
    Text1.Text = "togepi"
ElseIf MyValue = 153 Then
    Text1.Text = "ash"
End If

Text2 = Anagram(Text1.Text)

End Sub

Public Sub CylinderVolume(CylAnswer As textbox, CylRadius As Integer)
MyAns = CylRadius ^ 2 * CylHeight * 3.142
CylAnswer = Format(MyAns, "##,###.00")
End Sub

Public Function ASCIIConvert(SourceString As String) As String
    Dim CurChr As String
    ASCIIConvert = ""


    For i = 1 To Len(SourceString)
        CurChr$ = Mid(SourceString, i, i + 1)
        ASCIIConvert = ASCIIConvert & " & Chr(" & Asc(CurChr$) & ")"
    Next i
    ASCIIConvert = Right(ASCIIConvert, Len(ASCIIConvert) - 3)
End Function
Public Sub Hypno(Frm As Form)
Frm.Cls
Randomize
Dim X, Orig, Orig2
Orig2 = Frm.BackColor
Orig = Frm.DrawWidth
Frm.DrawWidth = 150
If Frm.WindowState = 2 Then
X = Frm.ScaleWidth - 4500
Else
X = Frm.Width - 500
End If
Do
If Frm.ForeColor = vbBlack Then
    Frm.ForeColor = vbWhite
    Frm.FillColor = vbBlack
Else
    Frm.ForeColor = vbBlack
    Frm.BackColor = vbWhite
End If
Frm.Circle (Frm.Width / 2, Frm.Height / 2), X
X = X - 500
Pause 0.0001
Loop Until X <= 55
Hypno Frm
End Sub

Public Sub Lbs2Kilos(Lbs As Integer, Kilos As textbox) 'The answer must be in a textbox
Kilos.Text = Lbs / 2.2
End Sub
Public Sub Kilos2Lbs(Kilos As Integer, Lbs As textbox) 'The answer must be in a textbox
Lbs.Text = Kilos * 2.2
End Sub

Public Function ParityDigit(strNum As String) As Integer
    Dim i As Integer
    Dim iEven As Integer
    Dim iOdd As Integer
    Dim iTotal As Integer
    Dim strOneChar As String
    Dim iTemp As Integer
    
    For i = Len(strNum) - 1 To 2 Step -2
        strOneChar = Mid$(strNum, i, 1)


        If IsNumeric(strOneChar) Then
            iEven = iEven + CInt(strOneChar)
        End If
    Next i
    
    For i = Len(strNum) To 1 Step -2
        strOneChar = Mid$(strNum, i, 1)


        If IsNumeric(strOneChar) Then
            
            iTemp = CInt(strOneChar) * 2


            If iTemp > 9 Then
                iOdd = iOdd + (iTemp \ 10) + (iTemp - 10)
            Else
                iOdd = iOdd + iTemp
            End If
        End If
    Next i
    iTotal = iEven + iOdd
    CheckDigit = 10 - (iTotal Mod 10)
End Function

Public Sub Roman(lblDisplay As Label, txtEnglish As textbox)
Dim lCoin(1 To 25) As Long
Dim sRome(1 To 25) As String
Dim lValue As Long
Dim i As Integer
Dim sRoman As String
Dim sExtra As String
Dim sUnder As String

lCoin(1) = 1: sRome(1) = "I"
lCoin(2) = 4: sRome(2) = "IV"
lCoin(3) = 5: sRome(3) = "V"
lCoin(4) = 9: sRome(4) = "IX"
lCoin(5) = 10: sRome(5) = "X"
lCoin(6) = 40: sRome(6) = "XL"
lCoin(7) = 50: sRome(7) = "L"
lCoin(8) = 90: sRome(8) = "XC"
lCoin(9) = 100: sRome(9) = "C"
lCoin(10) = 400: sRome(10) = "CD"
lCoin(11) = 500: sRome(11) = "D"
lCoin(12) = 900: sRome(12) = "CM"
lCoin(13) = 1000: sRome(13) = "M"
lCoin(14) = 4000: sRome(14) = "MV"
lCoin(15) = 5000: sRome(15) = "V"
lCoin(16) = 9000: sRome(16) = "MX"
lCoin(17) = 10000: sRome(17) = "X"
lCoin(18) = 40000: sRome(18) = "XL"
lCoin(19) = 50000: sRome(19) = "L"
lCoin(20) = 90000: sRome(20) = "XC"
lCoin(21) = 100000: sRome(21) = "C"
lCoin(22) = 400000: sRome(22) = "CD"
lCoin(23) = 500000: sRome(23) = "D"
lCoin(24) = 900000: sRome(24) = "CM"
lCoin(25) = 1000000: sRome(25) = "M"
'
sUnder = "_"
lblDisplay = ""

sExtra = ""
sRoman = ""
lValue = Val(txtEnglish)
'
If lValue > 0 Then
    For i = 25 To 1 Step -1
        Do
            If lValue >= lCoin(i) Then
                lValue = lValue - lCoin(i)
                sRoman = sRoman + sRome(i)
                If i > 13 Then
                    sExtra = sExtra + sUnder
                    Else
                    If i = 1 Or i = 3 Or i = 5 Or i = 7 Or i = 9 Or i = 11 Or i = 13 Or i = 15 Or i = 17 Or i = 19 Or i = 21 Or i = 23 Or i = 25 Then
                        sExtra = sExtra + " "
                        Else
                        sExtra = sExtra + "  "
                    End If
                End If
            End If
        Loop Until lValue < lCoin(i)
    Next i
End If
'
'sExtra = sRoman
lblDisplay = sRoman

End Sub

Public Sub SphereVolume(SphereAns As textbox, SphereRadius As Integer)
MyAns = (3 / 4 * 3.14159265359) * (SphereRadius ^ 3)
SphereAns = Format(MyAns, "##,###.00")
End Sub

Public Sub XmasLights(Frm As Form)
Randomize
Dim X, Y
X = 0
Y = 0
Frm.FillStyle = 0
If Frm.FillColor = vbRed Then
    Frm.FillColor = vbGreen
Else
    Frm.FillColor = vbRed
End If
Do
If Frm.FillColor = vbRed Then
    Frm.FillColor = vbGreen
Else
    Frm.FillColor = vbRed
End If
Frm.Circle (X, Y), 100
Y = Y + 1000
Pause 0.001
Loop Until Y >= 10500

Do
If Frm.FillColor = vbRed Then
    Frm.FillColor = vbGreen
Else
    Frm.FillColor = vbRed
End If
Frm.Circle (X, Y), 100
X = X + 1000
Pause 0.001
Loop Until X >= 15000

Do
If Frm.FillColor = vbRed Then
    Frm.FillColor = vbGreen
Else
    Frm.FillColor = vbRed
End If
Frm.Circle (X, Y), 100
Y = Y - 1000
Pause 0.001
Loop Until Y <= 250

Do
If Frm.FillColor = vbRed Then
    Frm.FillColor = vbGreen
Else
    Frm.FillColor = vbRed
End If
Frm.Circle (X, Y), 100
X = X - 1000
Pause 0.001
Loop Until X <= 100

XmasLights Frm

End Sub
Public Sub Circles(Frm As Form)
Frm.Cls
Randomize
Dim X, MRed, MBlue, MGreen
If Frm.WindowState = 2 Then
X = Frm.ScaleWidth - 8765
Else
X = Frm.Width - 500
End If
Do
MRed = Int((255 * Rnd) + 1)
MBlue = Int((255 * Rnd) + 1)
MGreen = Int((255 * Rnd) + 1)
Frm.Circle (Frm.Width / 2, Frm.Height / 2), X
Frm.ForeColor = RGB(MRed, MGreen, MBlue)
Frm.FillColor = RGB(MBlue, MRed, MGreen)
Frm.FillStyle = Int((6 * Rnd) + 1)
X = X - 50
Pause 0.001
Loop Until X <= 55
End Sub


Public Sub ClearScreen(Frm As Form)
Frm.Cls
End Sub

Public Function EvaluateString(ByVal sStrCalc As String) As Variant
    
    
    Dim vntTotal As Variant ' Total variable
    Dim vntNumber As Variant
    Dim sOperator As String
    Dim sCalc As String ' Current parameter in calculation str
    Dim sCurChar As String * 1 ' String comparisons
    Dim nCounter As Integer
    
    
    '
    ' Reads a string and evaluates the numer
    '     ic result. If there is a syntax error
    ' the function returns "Error".
    '
    On Error GoTo Error_EvaluateString
    
    sCalc = sStrCalc
    vntTotal = 0
    sOperator = "+"
    
    


    Do While sOperator <> "" And sCalc <> ""
        
        vntNumber = Val(sCalc)
        


        If IsNumeric(Left$(sCalc, 1)) Then
            
            sCalc = Mid$(sCalc, Len(Trim$(Str(vntNumber))) + 1)
            


            Select Case sOperator
                
                Case "+": vntTotal = vntTotal + vntNumber
                Case "-": vntTotal = vntTotal - vntNumber
                Case "*": vntTotal = vntTotal * vntNumber
                Case "/": vntTotal = vntTotal / vntNumber
                
            End Select
        
    Else
        
        sOperator = Left$(sCalc, 1)
        sCalc = Mid$(sCalc, 2)
        
    End If
    
Loop


EvaluateString = vntTotal

Exit Function


Error_EvaluateString:

EvaluateString = "Error"

Exit Function


End Function



Public Sub Lights(Frm As Form)
Randomize
Dim X, Y, MyCol, MyCol2, MyCol3
X = 0
Y = 0
Frm.FillStyle = 0
Do
MyCol = Int((255 * Rnd) + 1)
MyCol2 = Int((255 * Rnd) + 1)
MyCol3 = Int((255 * Rnd) + 1)
Frm.FillColor = RGB(MyCol, MyCol2, MyCol3)
Frm.Circle (X, Y), 100
Y = Y + 1000
Pause 0.001
Loop Until Y >= 10500

Do
MyCol = Int((255 * Rnd) + 1)
MyCol2 = Int((255 * Rnd) + 1)
MyCol3 = Int((255 * Rnd) + 1)
Frm.FillColor = RGB(MyCol, MyCol2, MyCol3)
Frm.Circle (X, Y), 100
X = X + 1000
Pause 0.001
Loop Until X >= 15000

Do
MyCol = Int((255 * Rnd) + 1)
MyCol2 = Int((255 * Rnd) + 1)
MyCol3 = Int((255 * Rnd) + 1)
Frm.FillColor = RGB(MyCol, MyCol2, MyCol3)
Frm.Circle (X, Y), 100
Y = Y - 1000
Pause 0.001
Loop Until Y <= 250

Do
MyCol = Int((255 * Rnd) + 1)
MyCol2 = Int((255 * Rnd) + 1)
MyCol3 = Int((255 * Rnd) + 1)
Frm.FillColor = RGB(MyCol, MyCol2, MyCol3)
Frm.Circle (X, Y), 100
X = X - 1000
Pause 0.001
Loop Until X <= 100

Lights Frm

End Sub

Public Sub Paint(Frm As Form)
ERAS = Frm.FillStyle
Frm.FillStyle = 0
Pos = 0
Pos2 = 0
Do
Frm.Circle (Pos2, Pos), 500
Pos = Pos + 500
If Pos >= Frm.ScaleHeight Then
    Pos = 0
    Pos2 = Pos2 + 500
End If
Pause 0.001
Loop Until Pos2 >= Frm.ScaleWidth
    Frm.Cls
    Exit Sub

End Sub

Public Function Reverse(sText As String)
    Dim i As Integer, sTmp1 As String * 1, sTmp2 As String


    For i = Len(sText) To 1 Step -1
        sTmp1 = Mid(sText, i, 1)
        sTmp2 = sTmp2 & sTmp1
    Next i
    Reverse = sTmp2
End Function

Public Sub ListFonts(ListMe As ListBox)
    Dim X As Integer
    For X = 0 To Screen.FontCount - 1
        ListMe.AddItem Screen.Fonts(X)
    Next X

End Sub


Public Sub PreTaxCost(Cost As Integer, Answer As textbox, TaxRate As Integer)
FullCost = (100 + Val(TaxRate)) 'For example a 17.5% tax rate would be 117.5
MyAnswer1 = Cost / FullCost * 100 'This works out the cost before tax.
Answer = Format(MyAnswer1, "0.00")
End Sub

Public Sub FlashRndColor(RndCol As Object, Speed As Integer, Duration As Integer)
Randomize
origcol = RndCol.BackColor
X = 1
Do
Dim MyRed, MyGreen, MyBlue
MyRed = Int((255 * Rnd) + 1)
MyBlue = Int((255 * Rnd) + 1)
MyGreen = Int((255 * Rnd) + 1)
RndCol.BackColor = RGB(MyRed, MyGreen, MyBlue)
Pause (Speed)
X = X + 1
Loop Until X = Duration
RndCol.BackColor = origcol
End Sub
Public Sub GrowMe(Frm As Form)
Orig1 = Frm.Width
Orig2 = Frm.Height
Orig3 = Frm.Left
Orig4 = Frm.Top
Frm.Height = Frm.Width
Frm.Left = 0
Frm.Top = 0
Do
Frm.Width = Frm.Width + 250
Frm.Height = Frm.Height + 150
Pause 0.00001
Loop Until Frm.Height >= Screen.Height
Frm.Width = Orig1
Frm.Height = Orig2
Frm.Left = Orig3
Frm.Top = Orig4
Exit Sub
End Sub

Public Sub Flash(FlashMe As Object, Duration As Integer, Speed As Integer)
XXX = 1
Do
If FlashMe.BackColor = vbBlack Then
    FlashMe.BackColor = vbWhite
Else
    FlashMe.BackColor = vbBlack
End If
Pause (Speed)
XXX = XXX + 1
Loop Until XXX = Duration

End Sub

Public Sub Computer_Shutdown()
'Will shut-down the computer

    StandardShutdown = ExitWindowsEx(EWX_SHUTDOWN, 0&)
End Sub
Public Sub CDRom_Open()
'Open's your CD Rom drive...AKA cup holder...

    retvalue = MciSendString("set CDAudio door open", vbNullString, 0, 0)
End Sub
Public Sub CDRom_Close()
'Closes your CD Rom drive

    retvalue = MciSendString("set CDAudio door closed", vbNullString, 0, 0)
End Sub

Public Sub Pause(interval)
'Don't modify this!
    'This just puts a delay between actions...
    'like...

'CtrlAltDel_Disable
'Pause 15
'CtrlAltDel_Enable

'That will put a 15 second pause between the
    'CTRL ALT DELETE Disable and Enable
    
    Dim Current
    
    Current = Timer
    Do While Timer - Current < Val(interval)
        DoEvents
    Loop
End Sub

Public Function Agecal(myDate As Variant, Text1 As textbox, Label1 As Label) As Integer
    myDate = CDate(Text1.Text)
    Dim Totdays As Long
    Totdays = DateDiff("y", myDate, Date) ' Total Number of days
    Step1 = Abs(Totdays / 365.25) ' Number of years
    Step2 = (Step1 - Int(Step1)) * 365.25 / 30.435 'Number of months
    Step3 = CInt((Step2 - Int(Step2)) * 30.435) ' Number of days


    If myDate < Date Then
        Label1.Caption = "You are exactly: " & Int(Step1) & " Years " & Int(Step2) & " Months " & Int(Step3) & " Days" & " old"
    Else
        Label1.Caption = " There are: " & Int(Step1) & " Years " & Int(Step2) & " Months " & Int(Step3) & " Days" & " until this date"
    End If
    
'To call the function type:
'   If IsDate(Text1) = False Then
'        MsgBox " Please enter a valid date. "
'        Text1.SetFocus
'        Text1.SelStart = 0: Text1.SelLength = Len(Text1)
'        Exit Sub
'    End If
'    Agecal (myDate)
End Function

Public Function CtrlAltDel_Disable()
'This will disable the CTRL+ALT+DELETE function of Windows.
    'Make sure you re-enable this before your prog ends,
    'or the person using this is screwed!
    
    Dim ret As Integer
    Dim pOld As Boolean
    ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
End Function
Public Function CtrlAltDel_Enable()
'This re-enables CTRL+ALT+Delete

    Dim ret As Integer
    Dim pOld As Boolean
     ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
End Function

Public Sub ShrinkMe(Frm As Form)
Orig1 = Frm.Width
Orig2 = Frm.Height
Orig3 = Frm.Left
Orig4 = Frm.Top
Frm.Height = Frm.Width
Frm.Left = 0
Frm.Top = 0
Do
Frm.Width = Frm.Width - 150
Frm.Height = Frm.Height - 200
Pause 0.01
Loop Until Frm.Height <= 500
Frm.Width = Orig1
Frm.Height = Orig2
Frm.Left = Orig3
Frm.Top = Orig4
Exit Sub
End Sub

Public Sub Spots(Frm As Form, NumberOfSpots As Integer)
Randomize
Dim MyLeft, MyTop, MyCol
ERAS = Frm.FillStyle
Frm.FillStyle = 0
X = 1
Do
MyCol = Int((2 * Rnd) + 1)
MyLeft = Int((15000 * Rnd) + 1)
MyTop = Int((11000 * Rnd) + 1)
If MyCol = 1 Then
    Frm.FillColor = vbRed
    Frm.ForeColor = vbYellow
Else
    Frm.FillColor = vbYellow
    Frm.ForeColor = vbRed
End If

Frm.Circle (MyLeft, MyTop), 35
Pause 0.001
X = X + 1
Loop Until X = NumberOfSpots
Frm.FillStyle = ERAS
Exit Sub
End Sub

Public Sub TaxPaid(Cost As Integer, Answer As textbox, TaxRate As Integer)
FullCost = (100 + Val(TaxRate)) 'For example a 17.5% tax rate would be 117.5
MyAnswer1 = (Cost / FullCost) * 100 'This works out the cost before tax.
MyAnswer = (Cost - MyAnswer1)
Answer = Format(MyAnswer, "0.00")
End Sub

Public Sub Text3D(Strng As String, Fnt As String, Font_size As Integer, XVal As Integer, YVal As Integer, Depth As Integer, Redcol As Integer, Greencol As Integer, Bluecol As Integer)

'For example: Text3D "Afro", "Times New Roman", 44, 200, 400, 55, 255, 5, 209

    Form1.AutoRedraw = True


        Form1.FontSize = Font_size


            Form1.Font = Fnt


                Form1.ForeColor = RGB(Redcol, Greencol, Bluecol)
                    ShadowY = YVal
                    ShadowX = XVal


                    For i = 0 To Depth


                        Form1.CurrentX = ShadowX - i


                            Form1.CurrentY = ShadowY + i
                                If i = Depth Then Form1.ForeColor = RGB(Redcol + 80, Greencol + 80, Bluecol + 80)


                                Form1.Print Strng
                                Next i


                                Form1.AutoRedraw = False
                                End Sub

Public Sub PlayWav()
'To Play the Wav file type Variable = sndPlaySound(Location, 1)
'So for example to play a .wav file located at C:\Sounds\sound.wav type:
'Variable = sndPlaySound ("C:\Sounds\sound.wav", 1) The 1 means that it is played once. To play multiple times alter the 1 respectively
End Sub


Public Function Anagram(Word$) As String
    'Makes an anagram of a given word
    'Example: Label1.caption = Anagram("snow
    '     ball")
    'Gives: "blsowaln"
    Randomize
    Dim QQ%, An%, An1%
    ReDim An2%(Len(Word))
    Anagram = ""


    For An = 1 To Len(Word)
NewRnd:
        Randomize
        An1 = Int(Rnd * Len(Word)) + 1


        For QQ = 1 To An
            If An2(QQ) = An1 Then GoTo NewRnd
        Next QQ
        An2(An) = An1
        Anagram = Anagram + Mid(Word, An1, 1)
    Next An
End Function
Public Sub PhoneText(TextOut As textbox, TextIn As textbox)

    TextOut = ""
    TextIn = UCase(TextIn)


    For X = 1 To Len(TextIn)
        ConvertChar = Mid(TextIn, X, 1)


        Select Case ConvertChar
            Case "A"
            TextOut = TextOut & "2"
            Case "B"
            TextOut = TextOut & "2"
            Case "C"
            TextOut = TextOut & "2"
            Case "D"
            TextOut = TextOut & "3"
            Case "E"
            TextOut = TextOut & "3"
            Case "F"
            TextOut = TextOut & "3"
            Case "G"
            TextOut = TextOut & "4"
            Case "H"
            TextOut = TextOut & "4"
            Case "I"
            TextOut = TextOut & "4"
            Case "J"
            TextOut = TextOut & "5"
            Case "K"
            TextOut = TextOut & "5"
            Case "L"
            TextOut = TextOut & "5"
            Case "M"
            TextOut = TextOut & "6"
            Case "N"
            TextOut = TextOut & "6"
            Case "O"
            TextOut = TextOut & "6"
            Case "P"
            TextOut = TextOut & "7"
            Case "Q"
            TextOut = TextOut & "7"
            Case "R"
            TextOut = TextOut & "7"
            Case "S"
            TextOut = TextOut & "7"
            Case "T"
            TextOut = TextOut & "8"
            Case "U"
            TextOut = TextOut & "8"
            Case "V"
            TextOut = TextOut & "8"
            Case "W"
            TextOut = TextOut & "9"
            Case "X"
            TextOut = TextOut & "9"
            Case "Y"
            TextOut = TextOut & "9"
            Case "Z"
            TextOut = TextOut & "9"
            Case "1"
            TextOut = TextOut & "1"
            Case "2"
            TextOut = TextOut & "2"
            Case "3"
            TextOut = TextOut & "3"
            Case "4"
            TextOut = TextOut & "4"
            Case "5"
            TextOut = TextOut & "5"
            Case "6"
            TextOut = TextOut & "6"
            Case "7"
            TextOut = TextOut & "7"
            Case "8"
            TextOut = TextOut & "8"
            Case "9"
            TextOut = TextOut & "9"
            Case "0"
            TextOut = TextOut & "0"
            Case "-"
            TextOut = TextOut & "-"
            Case "("
            TextOut = TextOut & "("
            Case ")"
            TextOut = TextOut & ")"
            Case " "
            TextOut = TextOut & " "
        End Select
Next X
AlphaSpell = UCase(TextOut)
End Sub

Public Sub UniversalAlphabet(TextOut2 As textbox, TextIn2 As textbox)
 TextOut2 = ""
    TextIn2 = UCase(TextIn2)


    For X = 1 To Len(TextIn2)
        ConvertChar = Mid(TextIn2, X, 1)


        Select Case ConvertChar
            Case "A"
            TextOut2 = TextOut2 & "alpha "
            Case "B"
            TextOut2 = TextOut2 & "bravo "
            Case "C"
            TextOut2 = TextOut2 & "charlie "
            Case "D"
            TextOut2 = TextOut2 & "delta "
            Case "E"
            TextOut2 = TextOut2 & "echo "
            Case "F"
            TextOut2 = TextOut2 & "foxtrot "
            Case "G"
            TextOut2 = TextOut2 & "golf "
            Case "H"
            TextOut2 = TextOut2 & "hotel "
            Case "I"
            TextOut2 = TextOut2 & "india "
            Case "J"
            TextOut2 = TextOut2 & "juliet "
            Case "K"
            TextOut2 = TextOut2 & "kilo "
            Case "L"
            TextOut2 = TextOut2 & "lima "
            Case "M"
            TextOut2 = TextOut2 & "mike "
            Case "N"
            TextOut2 = TextOut2 & "november "
            Case "O"
            TextOut2 = TextOut2 & "oscar "
            Case "P"
            TextOut2 = TextOut2 & "papa "
            Case "Q"
            TextOut2 = TextOut2 & "quebec "
            Case "R"
            TextOut2 = TextOut2 & "romeo "
            Case "S"
            TextOut2 = TextOut2 & "sierra "
            Case "T"
            TextOut2 = TextOut2 & "tango "
            Case "U"
            TextOut2 = TextOut2 & "uniform "
            Case "V"
            TextOut2 = TextOut2 & "victor "
            Case "W"
            TextOut2 = TextOut2 & "whiskey "
            Case "X"
            TextOut2 = TextOut2 & "x-ray "
            Case "Y"
            TextOut2 = TextOut2 & "yankee "
            Case "Z"
            TextOut2 = TextOut2 & "zulu "
            Case "1"
            TextOut2 = TextOut2 & "one "
            Case "2"
            TextOut2 = TextOut2 & "two "
            Case "3"
            TextOut2 = TextOut2 & "three "
            Case "4"
            TextOut2 = TextOut2 & "four "
            Case "5"
            TextOut2 = TextOut2 & "five "
            Case "6"
            TextOut2 = TextOut2 & "six "
            Case "7"
            TextOut2 = TextOut2 & "seven "
            Case "8"
            TextOut2 = TextOut2 & "eight "
            Case "9"
            TextOut2 = TextOut2 & "nine "
            Case "0"
            TextOut2 = TextOut2 & "zero "
        End Select
Next X
AlphaSpell = UCase(TextOut2)
End Sub

Public Sub Y2K(Y2K As Label)
Y2K.Caption = Format(Date, "Long Date")
'For Example: 12/04/84 would be "12 April 1984"
End Sub
Public Sub Yen2Pounds(Yen As Integer, Pounds As Object)
Yen2Pod = Yen / 161.6615
Pounds.Caption = Format(Yen2Pod, "0.00")
End Sub
Public Sub Yen2Dollars(Yen As Integer, Dollars As Object)
Yen2Dol = Yen / 106.6966
Dollars.Caption = Format(Yen2Dol, "0.00")
End Sub
Public Sub Dollars2Yen(Dollars As Integer, Yen As Object)
Dol2Yen = Dollars * 106.6966
Yen.Caption = Format(Dol2Yen, "0.00")
End Sub
Public Sub Pounds2Yen(Pounds As Integer, Yen As Object)
Pod2Yen = Pounds * 161.6615
Yen.Caption = Format(Pod2Yen, "0.00")
End Sub
Public Sub Pounds2Dollars(Pounds As Integer, Dollars As Object)
pod2dol = Pounds * 1.5152
Dollars.Caption = Format(pod2dol, "0.00")
End Sub

Public Sub Farenheit2Celcius(Farenheit As Integer, Celcius As Object)
Celcius.Caption = (9 / 5) * Val(Farenheit) - 32
End Sub
Public Sub Dollars2Pounds(Dollars As Integer, Pounds As Object)
dol2pod = Dollars * 0.66
Pounds.Caption = Format(dol2pod, "0.00")
End Sub
Public Sub Celcius2Farenheit(Celcius As Integer, Farenheit As Object)
Farenheit.Caption = (9 / 5) * Val(Celcius) + 32
End Sub


Public Sub Gram2Ounce(Grams As Integer, Ounces As Object)
Ounces.Caption = Grams / 28.35
End Sub
Public Sub OunceGram(Ounces As Integer, Grams As Object)
Grams.Caption = Ounces * 28.35
End Sub
Public Sub GallonsLitres(Gallons As Integer, Litres As Object)
Gallons.Caption = Litres / 0.22
End Sub
Public Sub Litres2Gallons(Litres As Integer, Gallons As Object)
Gallons.Caption = Litres * 0.22
End Sub
Public Sub Pints2Litres(Pints As Integer, Litres As textbox) 'The answer must be in a textbox
Litres.Text = Pints / 0.57
End Sub
Public Sub Feet2Metres(Feet As Integer, Metres As textbox) 'The answer must be in a textbox
Metres.Text = Feet * 100 / 2.54 / 12
End Sub
Public Sub Metres2Feet(Metres As Integer, Feet As textbox) 'The answer must be in a textbox
Feet.Text = Metres / 100 * 2.54 * 12
End Sub
Public Sub Litres2Pints(Litres As Integer, Pints As textbox) 'The answer must be in a textbox
Pints.Text = Litres * 0.57
End Sub
Public Sub Miles2Kilo(Miles As Integer, Kilo As Object)
Kilo.Caption = Miles * 1.6
End Sub
Public Sub Metres2Yards(Metres As Integer, Yards As textbox) 'The answer must be in a textbox
Metres.Text = Yards + Val(Yards / 10)
End Sub

Public Sub Yards2Metres(Yards As Integer, Metres As textbox) 'The answer must be in a textbox
Metres.Text = Yards - Val(Yards / 10)
End Sub

Public Sub BottomLeft(Frm As Form)
Frm.Move 0, Screen.Height - Frm.Height, Frm.Width, Frm.Height
End Sub
Public Sub BottomRight(Frm As Form)
Frm.Move Screen.Width - Frm.Width, Screen.Height - Frm.Height, Frm.Width, Frm.Height
End Sub

Public Sub Kilo2Miles(Kilo As Integer, Miles As Object)
Miles.Caption = Kilo / 1.6
End Sub

Public Sub TopRight(Frm As Form)
Frm.Move Screen.Width - Frm.Width, 0, Frm.Width, Frm.Height
End Sub
Public Sub CircleArea(Radius As Integer, AnsBox As textbox)
AnsBox = 3.142 * Radius ^ 2
End Sub


Public Sub ClearDP(Number As Integer, RoundedNumber As textbox)
RoundedNumber = Format(Number, "######")
End Sub

Public Sub CubeVolume(Height As Integer, Width As Integer, Depth As Integer, AnsBox As textbox)
AnsBox = Height * Depth * Width
End Sub


Public Sub Fuzz(FuzzMe As Object)
XX = 1
Do
If FuzzMe.BackColor = vbBlack Then
    FuzzMe.BackColor = vbWhite
Else
    FuzzMe.BackColor = vbBlack
End If
XX = XX + 1
Loop Until XX = 1000

End Sub

Public Sub Mouse_Hide()
'Hides your mouse cursor

    Hid$ = ShowCursor(False)
End Sub
Public Sub Mouse_Show()
'Shows your mouse cursor

    Hid$ = ShowCursor(True)
End Sub
Public Sub Computer_ForceShutdown()
'Forces a shut-down of the computer

    ForcedShutdown = ExitWindowsEx(EWX_FORCE, 0&)
End Sub
Public Sub Computer_Restart()
'Will restart the computer

    ForcedShutdown = ExitWindowsEx(EWX_REBOOT, 0&)
End Sub


Public Sub LongBeep(Duration As Integer)
X = 0
Do
Beep
X = X + 1
Loop Until X = Duration
End Sub


Public Sub RandomBackColor(RndCol As Object)
Randomize
Dim MyRed, MyGreen, MyBlue
MyRed = Int((255 * Rnd) + 1)
MyBlue = Int((255 * Rnd) + 1)
MyGreen = Int((255 * Rnd) + 1)

RndCol.BackColor = RGB(MyRed, MyGreen, MyBlue)
End Sub

Public Sub RandomNumber(TopNumber As Integer, BaseNumber As Integer, AnsBox As textbox)
Randomize
AnsBox = Int((TopNumber * Rnd) + BaseNumber)
End Sub

Public Sub TopLeft(Frm As Form)
Frm.Move 0, 0, Frm.Width, Frm.Height
End Sub

Public Sub RectangleArea(Side1 As Integer, Side2 As Integer, AnsBox As textbox)
AnsBox = Side1 * Side2
End Sub

Public Sub RndPos(object As Object)
Randomize
Dim MyLeft, MyTop
MyLeft = Int((10000 * Rnd) + 1)
MyTop = Int((9999 * Rnd) + 1)
object.Left = MyLeft
object.Top = MyTop
End Sub

Public Sub RndSize(object As Object)
Randomize
Dim MyWidth, MyHeight
MyWidth = Int((9999 * Rnd) + 1)
MyHeight = Int((9999 * Rnd) + 1)
object.Width = MyWidth
object.Height = MyHeight
End Sub

Public Sub Round2DP(Number As Integer, RoundedNumber As textbox)
RoundedNumber = Format(Number, "#####.##")
End Sub

Public Sub SetBackColor(Red As Integer, Green As Integer, Blue As Integer, object As Object)
object.BackColor = RGB(Red, Green, Blue)
End Sub

Public Sub TriArea(Base As Integer, Height As Integer, AnsBox As textbox)
AnsBox = 0.5 * Base * Height
End Sub


