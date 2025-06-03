VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TenScore Demo "
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12705
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   12705
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox Text4 
      Height          =   6255
      Left            =   6840
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   2400
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   11033
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":030A
   End
   Begin RichTextLib.RichTextBox Text2 
      Height          =   6255
      Left            =   240
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   2400
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   11033
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":038E
   End
   Begin VB.CommandButton cDemo 
      Caption         =   "Демо версия"
      Height          =   975
      Left            =   6000
      TabIndex        =   36
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton cPersonal 
      Caption         =   "Личные встречи"
      Enabled         =   0   'False
      Height          =   975
      Left            =   6000
      MousePointer    =   1  'Arrow
      TabIndex        =   35
      Top             =   2520
      Width           =   735
   End
   Begin RichTextLib.RichTextBox Text5 
      Height          =   375
      Left            =   6360
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2760
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0412
   End
   Begin VB.Frame Frame4 
      Caption         =   "Рейтинг:"
      Height          =   855
      Left            =   7860
      TabIndex        =   31
      Top             =   0
      Width           =   1050
      Begin VB.Line Line2 
         BorderColor     =   &H80000004&
         X1              =   120
         X2              =   840
         Y1              =   500
         Y2              =   500
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   120
         X2              =   840
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label5 
         Caption         =   "WTA"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   280
         TabIndex        =   33
         Top             =   540
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "ATP"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   280
         TabIndex        =   32
         Top             =   230
         UseMnemonic     =   0   'False
         Width           =   495
      End
   End
   Begin RichTextLib.RichTextBox Text6 
      Height          =   375
      Left            =   6120
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2760
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0496
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   8280
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      RemoteHost      =   "www.marathonbet.com"
      URL             =   "http://www.marathonbet.com/results.shtml"
      Document        =   "/results.shtml"
   End
   Begin VB.Frame Frame3 
      Caption         =   "Загрузка результатов из Интернета:"
      Height          =   1500
      Left            =   9000
      TabIndex        =   6
      Top             =   120
      Width           =   3495
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   300
         Left            =   1920
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   480
         Width           =   1270
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   662831105
         CurrentDate     =   38193
         MaxDate         =   40543
         MinDate         =   36526
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   360
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   480
         Width           =   1270
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         Format          =   662831105
         CurrentDate     =   38353
         MaxDate         =   38411
         MinDate         =   38353
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.CommandButton cStop 
         Caption         =   "...Прервать"
         Enabled         =   0   'False
         Height          =   300
         Left            =   1800
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   860
         Width           =   1095
      End
      Begin VB.CommandButton cLoad 
         Caption         =   "Загрузить..."
         Height          =   300
         Left            =   650
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   860
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Данные отсутствуют..."
         ForeColor       =   &H80000002&
         Height          =   225
         Left            =   120
         TabIndex        =   30
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "по"
         Height          =   255
         Left            =   1670
         TabIndex        =   8
         Top             =   530
         Width           =   265
      End
      Begin VB.Label Label1 
         Caption         =   "C"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   530
         Width           =   255
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Игра:"
      Height          =   855
      Left            =   3000
      TabIndex        =   5
      Top             =   0
      Width           =   1935
      Begin VB.OptionButton Option2 
         Caption         =   "Парная/Микст"
         Height          =   315
         Left            =   240
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Одиночная"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.CommandButton cExit 
      Cancel          =   -1  'True
      Caption         =   "Выход"
      Height          =   975
      Left            =   6000
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4680
      Width           =   735
   End
   Begin VB.CommandButton cSearch 
      Caption         =   "Поиск"
      Default         =   -1  'True
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   7320
      MaxLength       =   37
      TabIndex        =   1
      Text            =   "Срич"
      Top             =   1920
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   240
      MaxLength       =   37
      TabIndex        =   0
      Text            =   "одд"
      Top             =   1920
      Width           =   5175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Месяц:"
      Height          =   1815
      Left            =   240
      TabIndex        =   2
      Top             =   0
      Width           =   2655
      Begin VB.CheckBox Check1 
         Caption         =   "Декабрь"
         Enabled         =   0   'False
         Height          =   255
         Index           =   12
         Left            =   1320
         TabIndex        =   29
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Ноябрь"
         Enabled         =   0   'False
         Height          =   255
         Index           =   11
         Left            =   1320
         TabIndex        =   28
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Октябрь"
         Enabled         =   0   'False
         Height          =   255
         Index           =   10
         Left            =   1320
         TabIndex        =   27
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Сентябрь"
         Enabled         =   0   'False
         Height          =   255
         Index           =   9
         Left            =   1320
         TabIndex        =   26
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Август"
         Enabled         =   0   'False
         Height          =   255
         Index           =   8
         Left            =   1320
         TabIndex        =   25
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Июнь"
         Enabled         =   0   'False
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   24
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Май"
         Enabled         =   0   'False
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   23
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Апрель"
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   22
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Март"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   21
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Февраль"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Июль"
         Enabled         =   0   'False
         Height          =   255
         Index           =   7
         Left            =   1320
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Январь"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Вы используете демонстрационную версию программы TenScore"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   3000
      TabIndex        =   39
      Top             =   960
      Width           =   6015
   End
   Begin VB.Image Image1 
      Height          =   690
      Left            =   5040
      Picture         =   "Form1.frx":051A
      Top             =   120
      Width           =   2670
   End
   Begin VB.Label Label3 
      Caption         =   "producted by &REY "
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   11040
      TabIndex        =   19
      Top             =   1650
      UseMnemonic     =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim txt As String
Dim alltxt As String
Dim alltxt2 As String
Dim alltxt3 As String
Dim maxday As Integer
Dim p As Integer    ' универсальная переменная, применяется во многих функциях
Dim m As Integer
Dim d As Integer
Dim stopp As Integer
Dim X As Integer    ' для функции макс месяц и ПрогрессБар
Dim x2 As Integer   ' для функции макс месяц
Dim ChangeTo As Integer ' обе для ЧекБоксов в ряд
Private Declare Function ReleaseCapture Lib "user32" () As Long
Dim GOD As Integer
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lplplplpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1
Private Declare Function ActivateKeyboardLayout Lib "user32" (ByVal HKL As Long, ByVal Flags As Long) As Long
Const kb_lay_ru As Long = 68748313

Private Sub cDemo_Click()
'Form1.Enabled = False
'Form2.Show
'Form2.SetFocus
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
If State = 12 Then 'если документ полностью получен, то..
    Dim res As String, a As String
    Do
        a = Inet1.GetChunk(1024)
        res = res & a
    Loop While Len(a) > 0
    
    'Open "C:\yandex.txt" For Output As #1
    'Print #1, res
    'Close #1
'***********************************************
    Open App.Path + "\vremen.txt" For Output As #1
    Print #1, res 'Text5.Text
    Close #1
    'Text5.Text = ""

    alltxt = ""
    Open App.Path + "\vremen.txt" For Input As #1
    Do Until EOF(1)
    Line Input #1, txt
        alltxt = alltxt + txt + vbCrLf
    Loop
    Close #1
    Text5.Text = alltxt
'***********************************************
End If
End Sub


Private Sub cStop_Click()
stopp = 1
'Inet1.Cancel
Label4.Caption = "Загрузка прервана!.."
cStop.Enabled = False
End Sub
Private Sub Form_Load()
GOD = 2004 'Year(Date)
Form1.Caption = Form1.Caption & GOD
p = 1
'd = 1
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)

DTPicker1.Year = Year(Date)
DTPicker1.Month = Month(Date)
DTPicker1.Day = Day(Date)
DTPicker2.Year = Year(Date)
DTPicker2.Month = Month(Date)
DTPicker2.Day = Day(Date)
'DTPicker1.MaxDate = DTPicker2.Value    demo
'DTPicker2.MinDate = DTPicker1.Value    demo

'ActivateKeyboardLayout kb_lay_ru, 0 ' смена раскладки клавиатуры на RUS
vsecheki

End Sub

Sub vsecheki()
d = 1
'formcheck12
'formcheck11
'formcheck10
'formcheck9
'formcheck8
'formcheck7
'formcheck6
'formcheck5
'formcheck4
'formcheck3
formcheck2
formcheck1
End Sub
Private Sub cExit_Click()
Form1.WindowState = 1   ' прежде чем выключиться, окно свернется
End
End Sub
Private Sub cSearch_Click()
p = 1
alltxt = ""
alltxt2 = ""
Form1.MousePointer = 13
If Check1(1).Value = 1 Then
    Open App.Path + "\01.txt" For Input As #1
    Search
    Close #1
End If
If Check1(2).Value = 1 Then
    Open App.Path + "\02.txt" For Input As #1
    Search
    Close #1
End If
Text2.Text = alltxt
Text4.Text = alltxt2

Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)
Form1.MousePointer = 0
End Sub
Sub Search()
Dim neprav As Integer
neprav = 1
Do Until EOF(1)
Line Input #1, txt

If InStr(p, txt, ".20") <> 0 Then   ' защита года (дату в поиске не покажет)
    If InStr(p, txt, GOD) <> 0 Then neprav = 0 Else neprav = 1
Else
    If neprav = 0 Then  ' защита года


If Text1.Text <> "" Then
If InStr(p, txt, Text1.Text, vbTextCompare) <> 0 Then
    If Option1.Value = True Then        ' одиночная игра
        If InStr(p, txt, "/") <> 0 Then
        Else
        alltxt = alltxt + txt + vbCrLf
        End If
    End If
    If Option2.Value = True Then        ' парная/микст
        If InStr(p, txt, "/") <> 0 Then
        alltxt = alltxt + txt + vbCrLf
        End If
    End If
End If
End If
If Text3.Text <> "" Then
If InStr(p, txt, Text3.Text, vbTextCompare) <> 0 Then
    If Option1.Value = True Then        ' одиночная игра
        If InStr(p, txt, "/") <> 0 Then
        Else
        alltxt2 = alltxt2 + txt + vbCrLf
        End If
    End If
    If Option2.Value = True Then        ' парная/микст
        If InStr(p, txt, "/") <> 0 Then
        alltxt2 = alltxt2 + txt + vbCrLf
        End If
    End If
End If
End If
'--------------------------------------------------------------

'-------------------------------------------------------------
End If ' защита года
End If ' защита года
Loop
End Sub
'****************************************************************************************
Private Sub cLoad_Click()
Text5.Text = ""
Text6.Text = ""
stopp = 0
Label4.Caption = "Идет загрузка данных..."

'If DTPicker1.Year <= Year(Date) And DTPicker1.Month <= Month(Date) _
'                And DTPicker1.Day <= Day(Date) Then 'скачивает только прошлое
'If DTPicker1.Year <= DTPicker2.Year Then   ' если нет проблемы с годами
cStop.Enabled = True
cLoad.Enabled = False
Form1.MousePointer = 13

'------------------------------------------------------ деление несколько месяцев/один месяц
If DTPicker1.Month = DTPicker2.Month Then

    mday
    LoadOldfile      ' даст данные в alltxt2 и alltxt3

If DTPicker1.Day = DTPicker2.Day Then
    LoadInform       ' загрузка из InterNet
    KillNotTennis    ' ост. только ТЕННИС
    KillTag          ' убирает тэги
    KillBad          ' редактир. (по ставкам...)
    If Text5.Text <> "" Then
    Text6.Text = Text6.Text & DTPicker1.Day & "." & DTPicker1.Month & "." & DTPicker1.Year _
    & vbCrLf & Text5.Text & vbCrLf
    End If
        If stopp = 1 Then GoTo Kones
        'Label4.Caption = "Загрузка прервана!.."
        'GoTo Kones
        'End If
        
    'DTPicker1.Day = DTPicker1.Day + 1  ' счетчик по дням
    




'If DTPicker1.Day = maxday Then GoTo mainerror2
    

'If m = 50 Then
'mainerror2:
Kones:
zapis
Text6.Text = ""
If stopp = 0 Then Label4.Caption = "Загрузка завершена..."

End If
'Else: Label4.Caption = "Дата начала должна быть меньше..."
alltxt = ""
vsecheki
End If


cStop.Enabled = False
cLoad.Enabled = True
Form1.MousePointer = 0


End Sub
Sub mday()      ' максимальный день в каждом месяце
X = DTPicker1.Month
daymonth
End Sub

Sub daymonth()
x2 = DTPicker1.Year
If X = 1 Or X = 3 Or X = 5 Or X = 7 Or X = 8 Or X = 10 Or X = 12 Then maxday = 31
If X = 4 Or X = 6 Or X = 9 Or X = 11 Then maxday = 30

If x2 = 1988 Or x2 = 1992 Or x2 = 1994 Or x2 = 2000 Or x2 = 2004 Or x2 = 2008 _
    Or x2 = 2012 Or x2 = 2016 Or x2 = 2020 Then 'весокосный
    If X = 2 Then maxday = 29
    Else
    If X = 2 Then maxday = 28
End If
End Sub
Sub LoadInform()
Inet1.Execute "http://www.marathonbet.com/results.shtml?day=" _
   & DTPicker1.Day & "&month=" & DTPicker1.Month & "&year=" & DTPicker1.Year & "&tr=-1", "GET"
Do While Inet1.StillExecuting: DoEvents: Loop ' кольцо для нескольких дней запроса (на выделенке глюк)

'Text5.Text = Inet1.OpenURL("http://www.marathonbet.com/results.shtml?day=" _
   & DTPicker1.Day & "&month=" & DTPicker1.Month & "&year=" & DTPicker1.Year & "&tr=-1")

End Sub
Sub KillNotTennis()
Dim q As Integer
p = 1
If InStr(p, Text5.Text, "Теннис") <> 0 Then
p = InStr(p, Text5.Text, "Теннис") + Len("Теннис")
Text5.SetFocus
End If
If InStr(p, Text5.Text, "Теннис") <> 0 Then
q = InStr(p, Text5.Text, "Теннис") - 1
Text5.SelStart = 0
Text5.SelLength = q
Text5.SelText = ""
Else
Text5.Text = ""
End If

'Триатлон, Фехтование, Тяжелая атлетика, Формула-1, Футбол, Хоккей
p = 1
If InStr(p, Text5.Text, "Триатлон") <> 0 Then
Text5.SelStart = InStr(p, Text5.Text, "Триатлон") - 1
Text5.SelLength = Len(Text5.Text) - Text5.SelStart
Text5.SelText = ""
End If

p = 1
If InStr(p, Text5.Text, "Тяжелая атлетика") <> 0 Then
Text5.SelStart = InStr(p, Text5.Text, "Тяжелая атлетика") - 1
Text5.SelLength = Len(Text5.Text) - Text5.SelStart
Text5.SelText = ""
End If

p = 1
If InStr(p, Text5.Text, "Фехтование") <> 0 Then
Text5.SelStart = InStr(p, Text5.Text, "Фехтование") - 1
Text5.SelLength = Len(Text5.Text) - Text5.SelStart
Text5.SelText = ""
End If

p = 1
If InStr(p, Text5.Text, "Формула") <> 0 Then
Text5.SelStart = InStr(p, Text5.Text, "Формула") - 1
Text5.SelLength = Len(Text5.Text) - Text5.SelStart
Text5.SelText = ""
End If

p = 1
If InStr(p, Text5.Text, "Футбол") <> 0 Then
Text5.SelStart = InStr(p, Text5.Text, "Футбол") - 1
Text5.SelLength = Len(Text5.Text) - Text5.SelStart
Text5.SelText = ""
End If

p = 1
If InStr(p, Text5.Text, "Хоккей") <> 0 Then
Text5.SelStart = InStr(p, Text5.Text, "Хоккей") - 1
Text5.SelLength = Len(Text5.Text) - Text5.SelStart
Text5.SelText = ""
End If

p = 1
If InStr(p, Text5.Text, "</pre></div>") <> 0 Then   ' если был только ТЕННИС
Text5.SelStart = InStr(p, Text5.Text, "</pre></div>") - 1
Text5.SelLength = Len(Text5.Text) - Text5.SelStart
Text5.SelText = ""
End If

End Sub
Sub KillTag()
p = 1
Do While InStr(p, Text5.Text, "</span><pre>")
If InStr(p, Text5.Text, "</span><pre>") <> 0 Then
Text5.SelStart = InStr(p, Text5.Text, "</span><pre>") - 1
Text5.SelLength = Len("</span><pre>")
Text5.SelText = ""
End If
If InStr(p, Text5.Text, "</pre><span class=cap>") <> 0 Then
Text5.SelStart = InStr(p, Text5.Text, "</pre><span class=cap>") - 1
Text5.SelLength = Len("</pre><span class=cap>")
Text5.SelText = ""
End If       ' либо переставить в допл экшен (чтобы работало быстрее) VIP
If InStr(p, Text5.Text, "&nbsp;") <> 0 Then
Text5.SelStart = InStr(p, Text5.Text, "&nbsp;") - 1
Text5.SelLength = Len("&nbsp;")
Text5.SelText = ""
End If
Loop
End Sub
Sub KillBad()
Dim n As Integer 'позиция курсора

p = 1
Do While InStr(p, Text5.Text, "Матч отменен", vbTextCompare)
If InStr(p, Text5.Text, "Матч отменен", vbTextCompare) <> 0 Then
Text5.SelStart = InStr(p, Text5.Text, "Матч отменен", vbTextCompare) + 11
End If
p = Text5.SelStart
If InStr(p, Text5.Text, vbCrLf) <> 0 Then
    Text5.SetFocus
    n = InStr(p, Text5.Text, vbCrLf) - 1
    Text5.SelLength = n - Text5.SelStart
    Text5.SelText = ""
End If
Loop

p = 1
Do While InStr(p, Text5.Text, "(в", vbTextCompare) ' (Вы... -  с Большой буквы
If InStr(p, Text5.Text, "(в", vbTextCompare) <> 0 Then
Text5.SelStart = InStr(p, Text5.Text, "(в", vbTextCompare) - 2
End If
p = Text5.SelStart
If InStr(p, Text5.Text, vbCrLf) <> 0 Then
    Text5.SetFocus
    n = InStr(p, Text5.Text, vbCrLf) - 1
    Text5.SelLength = n - Text5.SelStart
    Text5.SelText = ""
End If
Loop

p = 1
Do While InStr(p, Text5.Text, "(По", vbTextCompare)
If InStr(p, Text5.Text, "(По", vbTextCompare) <> 0 Then
Text5.SelStart = InStr(p, Text5.Text, "(По", vbTextCompare) - 2
End If
p = Text5.SelStart
If InStr(p, Text5.Text, vbCrLf) <> 0 Then
    Text5.SetFocus
    n = InStr(p, Text5.Text, vbCrLf) - 1
    Text5.SelLength = n - Text5.SelStart
    Text5.SelText = ""
End If
Loop

p = 1
Do While InStr(p, Text5.Text, "( По", vbTextCompare)   ' ( По... -  с Маленькой буквы
If InStr(p, Text5.Text, "( По", vbTextCompare) <> 0 Then
Text5.SelStart = InStr(p, Text5.Text, "( По", vbTextCompare) - 2
End If
p = Text5.SelStart
If InStr(p, Text5.Text, vbCrLf) <> 0 Then
    Text5.SetFocus
    n = InStr(p, Text5.Text, vbCrLf) - 1
    Text5.SelLength = n - Text5.SelStart
    Text5.SelText = ""
End If
Loop

p = 1
Do While InStr(p, Text5.Text, "(и", vbTextCompare) ' (Изменен формат третьего сета
If InStr(p, Text5.Text, "(и", vbTextCompare) <> 0 Then
Text5.SelStart = InStr(p, Text5.Text, "(и", vbTextCompare) - 2
End If
p = Text5.SelStart
If InStr(p, Text5.Text, vbCrLf) <> 0 Then
    Text5.SetFocus
    n = InStr(p, Text5.Text, vbCrLf) - 1
    Text5.SelLength = n - Text5.SelStart
    Text5.SelText = ""
End If
Loop

End Sub
'******************************************************* пошла муть (кроме "qwe")
Sub LoadOldfile()
If DTPicker1.Month = 1 Then
    On Error GoTo error1
    If p = 50 Then
error1:
        Open App.Path + "\01.txt" For Output As #1
        Close #1
    End If
    Open App.Path + "\01.txt" For Input As #1
    qwe
    Close #1
End If
If DTPicker1.Month = 2 Then
    On Error GoTo error2
    If p = 50 Then
error2:
        Open App.Path + "\02.txt" For Output As #1
        Close #1
    End If
    Open App.Path + "\02.txt" For Input As #1
    qwe
    Close #1
End If
End Sub
Sub qwe()       ' разборка по датам старый файл
Dim q As Integer
q = 0
alltxt2 = ""
alltxt3 = ""

Do Until EOF(1)
Line Input #1, txt
If q = 0 Then ' первая часть alltxt2
If InStr(txt, DTPicker1.Day & "." & DTPicker1.Month & "." & DTPicker1.Year) Then
    q = 1
    Else
    For p = DTPicker1.Day To 31   '
        If InStr(txt, p & "." & DTPicker1.Month & "." & DTPicker1.Year) Then q = 1
    Next p
    If q = 0 Then alltxt2 = alltxt2 + txt + vbCrLf
End If
End If
'---------------------------------------------------------------------
If q = 1 Then ' первая часть alltxt3
If InStr(txt, DTPicker2.Day + 1 & "." & DTPicker2.Month & "." & DTPicker2.Year) Then
    q = 2
    Else
    For p = DTPicker2.Day + 1 To 31 '
        If InStr(txt, p & "." & DTPicker2.Month & "." & DTPicker2.Year) Then q = 2
    Next p
End If
End If
If q = 2 Then
    alltxt3 = alltxt3 + txt + vbCrLf
End If
Loop
'Text6.Text = alltxt2
'Text5.Text = alltxt3
End Sub
Sub zapis()
If DTPicker1.Month = 1 Then
    Open App.Path + "\01.txt" For Output As #1
    Print #1, alltxt2
    Print #1, Text6.Text
    Print #1, alltxt3
    Close #1
End If
If DTPicker1.Month = 2 Then
    Open App.Path + "\02.txt" For Output As #1
    Print #1, alltxt2
    Print #1, Text6.Text
    Print #1, alltxt3
    Close #1
End If
End Sub

Sub formcheck1()
On Error GoTo errors
    Open App.Path + "\01.txt" For Input As #1

    If d = 1 Then
    Do Until EOF(1)
    Line Input #1, txt
    alltxt = alltxt + txt + vbCrLf
    Loop
    Text5.Text = alltxt
    End If
    Close #1
    
    If d = 1 Then
    For m = 31 To 1 Step -1
    If InStr(p, Text5.Text, m & "." & 1 & "." & GOD) <> 0 Then GoTo bla
    Next m
GoTo errors  ' если не найдена ни одна дата в месяце за нужный GOD (kill - may be)
bla:
Label4.Caption = "Последние данные на:" & " " & m & " января."
Text5.Text = ""
    d = 0
    
    End If
    Check1(1).Enabled = True
    Check1(1).Value = 1
If p = 50 Then
errors:
Check1(1).Enabled = False
Check1(1).Value = 0
End If
End Sub
Sub formcheck2()
On Error GoTo errors
    Open App.Path + "\02.txt" For Input As #1

    If d = 1 Then
    Do Until EOF(1)
    Line Input #1, txt
    alltxt = alltxt + txt + vbCrLf
    Loop
    Text5.Text = alltxt
    End If
    Close #1
    
    If d = 1 Then
    For m = 29 To 1 Step -1
    If InStr(p, Text5.Text, m & "." & 2 & "." & GOD) <> 0 Then GoTo bla
    Next m
GoTo errors  ' если не найдена ни одна дата в месяце за нужный GOD (kill - may be)
bla:
Label4.Caption = "Последние данные на:" & " " & m & " февраля."
Text5.Text = ""
    d = 0
    
    End If
    Check1(2).Enabled = True
    Check1(2).Value = 1
If p = 50 Then
errors:
Check1(2).Enabled = False
Check1(2).Value = 0
End If
End Sub


Private Sub Check1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then _
  ChangeTo = 1 - Check1(Index).Value: _
  Check1(Index).Value = ChangeTo: _
  ReleaseCapture
End Sub

Private Sub Check1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then Check1(Index).Value = ChangeTo
End Sub

Private Sub Image1_Click()
ShellExecute Me.hWnd, vbNullString, "http://www.marathonbet.com/odds.shtml", vbNullString, vbNullString, SW_SHOWNORMAL
'ShellExecute Me.hWnd, vbNullString, "mailto:gaidar@vbstreets.ru?subject=test", vbNullString, vbNullString, SW_SHOWNORMAL 'для почты
End Sub

Private Sub Label5_Click()
ShellExecute Me.hWnd, vbNullString, "http://www.marathonbet.com/tab_2t.shtml?razd=-2", vbNullString, vbNullString, SW_SHOWNORMAL
Label5.ForeColor = &HFF0000
Label5.FontUnderline = False
End Sub

Private Sub Label7_Click()
ShellExecute Me.hWnd, vbNullString, "http://www.marathonbet.com/tab_2t.shtml?razd=-1", vbNullString, vbNullString, SW_SHOWNORMAL
Label7.ForeColor = &HFF0000
Label7.FontUnderline = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = &HFF0000
Label7.FontUnderline = False
Label5.ForeColor = &HFF0000
Label5.FontUnderline = False
End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = &HFF0000
Label7.FontUnderline = False
Label5.ForeColor = &HFF0000
Label5.FontUnderline = False
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = &HFF&
Label7.FontUnderline = True
Label5.ForeColor = &HFF0000
Label5.FontUnderline = False
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = &HFF&
Label5.FontUnderline = True
Label7.ForeColor = &HFF0000
Label7.FontUnderline = False
End Sub









