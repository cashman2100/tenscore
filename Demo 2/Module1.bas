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
Open "option.ini" For Output As #1
Print #1, Player1 & vbCrLf & Player2
Close #1

End Sub


Sub OptionsLoad()
On Error Resume Next
Open "option.ini" For Input As #1
Line Input #1, txt
     Form1.Text1.Text = txt
Line Input #1, txt
     Form1.Text3.Text = txt
If Form1.Text1.Text = "" Then Form1.Text1.Text = "Теннисист1"
If Form1.Text3.Text = "" Then Form1.Text3.Text = "Теннисист2"
End Sub

Sub vsecheki()
For File = 1 To 12
If Len(Dir(CStr(GOD) + "\" + CStr(File) + ".txt")) Then
    Form1.Check1(File).Enabled = True
    Form1.Check1(File).Value = 1
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








