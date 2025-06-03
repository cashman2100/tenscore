Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function ReleaseCapture Lib "user32" () As Long ' для чекбоксов
Public ChangeTo As Integer ' обе для ЧекБоксов в ряд
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lplplplpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1 ' эта и предыдущая (обе) для ссылок
Public Linklabel As String
Public CheckSingl As Integer
Public CheckMonth As Integer
Public CheckTUR As Integer
Public hard As Integer
Public soil As Integer
Public corpet As Integer
Public grass As Integer
Public noncover As Integer
Public Player1 As String
Public Player2 As String
Public GOD As Integer
Public txt As String
Public alltxt As String
Public p As Integer    ' универсальная переменная, применяется во многих функциях
Public d As Integer       ' допол. универсальная переменная
Public File As Integer
Public Raskladka As Integer
Public Correct As Integer
Public m As Integer
Public x As Integer    ' для функции макс месяц и ПрогрессБар
Public First As Integer ' важно для первого запуска/последующей работы
Public AllGod As Integer ' для поиска по всем годам

Sub OptionsSave()
Player1 = Form1.Text1.Text
Player2 = Form1.Text3.Text
If Form1.Option1.Value = True Then CheckSingl = 1 Else CheckSingl = 0
If Form2.Check6.Value = 1 Then Correct = 1 Else Correct = 0
If Form2.Check5.Value = 1 Then AllGod = 1 Else AllGod = 0
Open "option.ini" For Output As #1
Print #1, Linklabel & vbCrLf & Player1 & vbCrLf & Player2 & vbCrLf & _
    CheckSingl & vbCrLf & CheckMonth & vbCrLf & CheckTUR & vbCrLf & _
    hard & vbCrLf & soil & vbCrLf & corpet & vbCrLf & grass & vbCrLf & noncover _
    & vbCrLf & Raskladka & vbCrLf & Correct & vbCrLf & _
    GOD & vbCrLf & AllGod
Close #1
SaveGOD
End Sub


Sub OptionsLoad()
On Error Resume Next

Open "option.ini" For Input As #1
Line Input #1, txt
     Linklabel = txt
Line Input #1, txt
     Form1.Text1.Text = txt
Line Input #1, txt
     Form1.Text3.Text = txt

Line Input #1, txt
     If txt = 1 Then CheckSingl = 1 Else CheckSingl = 0
Line Input #1, txt
     If txt = 1 Then CheckMonth = 1 Else CheckMonth = 0
Line Input #1, txt
     If txt = "1" Then CheckTUR = 1 Else CheckTUR = 0
Line Input #1, txt
     If txt = 1 Then hard = 1 ' Else hard = 0
Line Input #1, txt
     If txt = 1 Then soil = 1 'Else soil = 0
Line Input #1, txt
     If txt = 1 Then corpet = 1 ' Else corpet = 0
Line Input #1, txt
     If txt = 1 Then grass = 1 'Else grass = 0
Line Input #1, txt
     If txt = 1 Then noncover = 1 'Else noncover = 0
Line Input #1, txt
     If txt = "1" Then Raskladka = 1 Else Raskladka = 0
Line Input #1, txt
     If txt = "1" Then Correct = 1 Else Correct = 0

Line Input #1, txt
    GOD = txt
Line Input #1, txt
     If txt = "1" Then AllGod = 1 Else AllGod = 0
Close #1

If Linklabel = "" Then Linklabel = "http://www.marathonbet.com/odds.shtml"
If CheckSingl <> 0 Then CheckSingl = 1
If CheckMonth <> 0 Then CheckMonth = 1
If CheckTUR <> 1 Then CheckTUR = 0
If Correct <> 1 Then Correct = 0
If AllGod <> 1 Then AllGod = 0

If Form1.Text1.Text = "" Then Form1.Text1.Text = "Теннисист1"
If Form1.Text3.Text = "" Then Form1.Text3.Text = "Теннисист2"
End Sub

Sub SaveGOD()
On Error GoTo Kones
Open App.Path + "\" + CStr(GOD) + "\" + CStr(GOD) + ".txt" For Output As #1
Print #1, Form1.Check1(1).Value & vbCrLf & Form1.Check1(2).Value & vbCrLf & _
    Form1.Check1(3).Value & vbCrLf & Form1.Check1(4).Value & vbCrLf & _
    Form1.Check1(5).Value & vbCrLf & Form1.Check1(6).Value & vbCrLf & _
    Form1.Check1(7).Value & vbCrLf & Form1.Check1(8).Value & vbCrLf & _
    Form1.Check1(9).Value & vbCrLf & Form1.Check1(10).Value & vbCrLf & _
    Form1.Check1(11).Value & vbCrLf & Form1.Check1(12).Value
Close #1
Kones:
End Sub


Sub LoadGOD()
On Error Resume Next
Open App.Path + "\" + CStr(GOD) + "\" + CStr(GOD) + ".txt" For Input As #1
Line Input #1, txt
    If txt = 1 Then Form1.Check1(1).Value = 1 Else Form1.Check1(1).Value = 0
Line Input #1, txt
    If txt = 1 Then Form1.Check1(2).Value = 1 Else Form1.Check1(2).Value = 0
Line Input #1, txt
    If txt = 1 Then Form1.Check1(3).Value = 1 Else Form1.Check1(3).Value = 0
Line Input #1, txt
    If txt = 1 Then Form1.Check1(4).Value = 1 Else Form1.Check1(4).Value = 0
Line Input #1, txt
    If txt = 1 Then Form1.Check1(5).Value = 1 Else Form1.Check1(5).Value = 0
Line Input #1, txt
    If txt = 1 Then Form1.Check1(6).Value = 1 Else Form1.Check1(6).Value = 0
Line Input #1, txt
    If txt = 1 Then Form1.Check1(7).Value = 1 Else Form1.Check1(7).Value = 0
Line Input #1, txt
    If txt = 1 Then Form1.Check1(8).Value = 1 Else Form1.Check1(8).Value = 0
Line Input #1, txt
    If txt = 1 Then Form1.Check1(9).Value = 1 Else Form1.Check1(9).Value = 0
Line Input #1, txt
    If txt = 1 Then Form1.Check1(10).Value = 1 Else Form1.Check1(10).Value = 0
Line Input #1, txt
    If txt = 1 Then Form1.Check1(11).Value = 1 Else Form1.Check1(11).Value = 0
Line Input #1, txt
    If txt = 1 Then Form1.Check1(12).Value = 1 Else Form1.Check1(12).Value = 0
Close #1
End Sub

Sub vsecheki()
For File = 1 To 12
If Len(Dir(CStr(GOD) + "\" + CStr(File) + ".txt")) Then
    Form1.Check1(File).Enabled = True
    Else
    Form1.Check1(File).Enabled = False
    Form1.Check1(File).Value = 0
    End If
Next File
End Sub

Sub LastData()
alltxt = ""
On Error Resume Next
For x = 12 To 1 Step -1
Open App.Path + "\" + CStr(GOD) + "\" + CStr(x) + ".txt" For Input As #1
alltxt = Input(LOF(1), 1)
Close #1
    
If InStr(p, alltxt, GOD) <> 0 Then
    For m = 31 To 1 Step -1
    If InStr(p, alltxt, m & "." & x & "." & GOD) <> 0 Then GoTo bla
    Next m
    GoTo errors
bla:
    If x = 1 Then Form1.Label4.Caption = "Последние данные на:" & " " & m & " января " & GOD & "."
    If x = 2 Then Form1.Label4.Caption = "Последние данные на:" & " " & m & " февраля " & GOD & "."
    If x = 3 Then Form1.Label4.Caption = "Последние данные на:" & " " & m & " марта " & GOD & "."
    If x = 4 Then Form1.Label4.Caption = "Последние данные на:" & " " & m & " апреля " & GOD & "."
    If x = 5 Then Form1.Label4.Caption = "Последние данные на:" & " " & m & " мая " & GOD & "."
    If x = 6 Then Form1.Label4.Caption = "Последние данные на:" & " " & m & " июня " & GOD & "."
    If x = 7 Then Form1.Label4.Caption = "Последние данные на:" & " " & m & " июля " & GOD & "."
    If x = 8 Then Form1.Label4.Caption = "Последние данные на:" & " " & m & " августа " & GOD & "."
    If x = 9 Then Form1.Label4.Caption = "Последние данные на:" & " " & m & " сентября " & GOD & "."
    If x = 10 Then Form1.Label4.Caption = "Последние данные на:" & " " & m & " октября " & GOD & "."
    If x = 11 Then Form1.Label4.Caption = "Последние данные на:" & " " & m & " ноября " & GOD & "."
    If x = 12 Then Form1.Label4.Caption = "Последние данные на:" & " " & m & " декабря " & GOD & "."
    Exit Sub
End If
errors:
Next x
End Sub

Sub ForSearchGOD()
If GOD = 2000 Then
    Form1.Option3.Caption = "2000"
    Form1.Option4.Caption = "2001"
    Form1.Option5.Caption = "2002"
    Form1.Option3.Value = True
    Exit Sub
    End If
If GOD = Year(Date) Or GOD < 2000 Then
    Form1.Option3.Caption = Year(Date) - 2 '"2003"
    Form1.Option4.Caption = Year(Date) - 1 '"2004"
    Form1.Option5.Caption = Year(Date) '"2005"
    Form1.Option5.Value = True
    Exit Sub
    End If
Form1.Option3.Caption = GOD - 1
Form1.Option4.Caption = GOD
Form1.Option5.Caption = GOD + 1
Form1.Option4.Value = True
End Sub








