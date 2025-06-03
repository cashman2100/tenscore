VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A51095D7-8D17-11D6-9913-E1D1DF4BFD40}#1.0#0"; "XPButton.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TenScore 2005"
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
   Begin XPButton.UserControl1 HDDdetectByWMI 
      Height          =   255
      Left            =   9000
      TabIndex        =   48
      Top             =   1635
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      Caption         =   "#HDD by WMI (only XP)"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DefCurHand      =   0   'False
   End
   Begin XPButton.UserControl1 cExit 
      Height          =   975
      Left            =   6000
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   4680
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1720
      Caption         =   "Выход"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DefCurHand      =   0   'False
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Command2"
      Height          =   255
      Left            =   6240
      TabIndex        =   45
      Top             =   5040
      Width           =   255
   End
   Begin XPButton.UserControl1 cSearch 
      Height          =   375
      Left            =   5520
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Поиск"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16761024
      DefCurHand      =   0   'False
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   195
      Left            =   5880
      TabIndex        =   43
      Top             =   2040
      Width           =   975
   End
   Begin XPButton.UserControl1 cStop 
      Height          =   255
      Left            =   10800
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      Caption         =   "...Прервать"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DefCurHand      =   0   'False
   End
   Begin XPButton.UserControl1 cLoad 
      Height          =   255
      Left            =   9600
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      Caption         =   "Загрузить..."
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DefCurHand      =   0   'False
   End
   Begin XPButton.UserControl1 cOption 
      Height          =   975
      Left            =   6000
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   3600
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1720
      Caption         =   "Опции"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DefCurHand      =   0   'False
   End
   Begin XPButton.UserControl1 cPersonal 
      Height          =   975
      Left            =   6000
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   2520
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1720
      Caption         =   "Личные встречи"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DefCurHand      =   0   'False
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   195
      Left            =   4200
      TabIndex        =   38
      Top             =   1200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   195
      Left            =   4200
      TabIndex        =   37
      Top             =   1440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   195
      Left            =   4200
      TabIndex        =   36
      Top             =   960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame5 
      Caption         =   "Год:"
      Height          =   975
      Left            =   3000
      TabIndex        =   33
      Top             =   840
      Width           =   975
      Begin VB.OptionButton Option3 
         Caption         =   "2003"
         Height          =   240
         Left            =   120
         TabIndex        =   47
         Top             =   200
         Width           =   735
      End
      Begin VB.OptionButton Option5 
         Caption         =   "2005"
         Height          =   240
         Left            =   120
         TabIndex        =   35
         Top             =   675
         Width           =   735
      End
      Begin VB.OptionButton Option4 
         Caption         =   "2004"
         Height          =   240
         Left            =   120
         TabIndex        =   34
         Top             =   435
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin RichTextLib.RichTextBox Text4 
      Height          =   6255
      Left            =   6840
      TabIndex        =   32
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
      TabIndex        =   31
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
   Begin RichTextLib.RichTextBox Text5 
      Height          =   375
      Left            =   6360
      TabIndex        =   30
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
      TabIndex        =   27
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
         TabIndex        =   29
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
         TabIndex        =   28
         Top             =   230
         UseMnemonic     =   0   'False
         Width           =   495
      End
   End
   Begin RichTextLib.RichTextBox Text6 
      Height          =   375
      Left            =   6120
      TabIndex        =   14
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
      Top             =   960
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
      TabIndex        =   4
      Top             =   120
      Width           =   3495
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   300
         Left            =   1920
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   480
         Width           =   1270
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         Format          =   662831105
         CurrentDate     =   38193
         MaxDate         =   40543
         MinDate         =   36526
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   360
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   480
         Width           =   1270
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         Format          =   662831105
         CurrentDate     =   38193
         MaxDate         =   40543
         MinDate         =   36526
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Данные отсутствуют..."
         ForeColor       =   &H80000002&
         Height          =   225
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "по"
         Height          =   255
         Left            =   1670
         TabIndex        =   6
         Top             =   530
         Width           =   265
      End
      Begin VB.Label Label1 
         Caption         =   "C"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   530
         Width           =   255
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Игра:"
      Height          =   855
      Left            =   3000
      TabIndex        =   3
      Top             =   0
      Width           =   1935
      Begin VB.OptionButton Option2 
         Caption         =   "Парная/Микст"
         Height          =   315
         Left            =   240
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Одиночная"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   7320
      MaxLength       =   37
      TabIndex        =   1
      Top             =   1920
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   240
      MaxLength       =   37
      TabIndex        =   0
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
         TabIndex        =   25
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Ноябрь"
         Enabled         =   0   'False
         Height          =   255
         Index           =   11
         Left            =   1320
         TabIndex        =   24
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Октябрь"
         Enabled         =   0   'False
         Height          =   255
         Index           =   10
         Left            =   1320
         TabIndex        =   23
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Сентябрь"
         Enabled         =   0   'False
         Height          =   255
         Index           =   9
         Left            =   1320
         TabIndex        =   22
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Август"
         Enabled         =   0   'False
         Height          =   255
         Index           =   8
         Left            =   1320
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Июль"
         Enabled         =   0   'False
         Height          =   255
         Index           =   7
         Left            =   1320
         TabIndex        =   10
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
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
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
      TabIndex        =   15
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
Dim var As Variant
Dim alltxt1 As String ' только при обработке данных с Инета
Dim alltxt2 As String
Dim alltxt3 As String
Dim maxday As Integer
Dim stopp As Integer
Dim PerSent As Double
Dim Pers As Integer ' для правильного показа кнопки Personal после загрузки
Private Declare Function ActivateKeyboardLayout Lib "user32" (ByVal HKL As Long, ByVal flags As Long) As Long
Const kb_lay_ru As Long = 68748313
Dim Iskomoe1 As String      ' якобы ускоряет поиск
Dim Iskomoe2 As String
Dim Tur1 As String
Dim Tur2 As String
Dim Tur3 As String
Dim Tur4 As String  ' для совместных игр
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_VSCROLL = &H115
Const SB_LINEUP = 4
'Const SB_LINELEFT = 0
'Const SB_LINEDOWN = 1
'Const SB_LINERIGHT = 1
'Const SB_PAGEUP = 2
'Const SB_PAGELEFT = 2
'Const SB_PAGEDOWN = 3
'Const SB_PAGERIGHT = 3
'Const SB_THUMBPOSITION = 4
'Const SB_THUMBTRACK = 5
'Const SB_TOP = 6
'Const SB_LEFT = 6
'Const SB_BOTTOM = 7
'Const SB_RIGHT = 7
'Const SB_ENDSCROLL = 8

'Private Sub ScrollUpB_Click()
'SendMessageLong Me.RText.hwnd, WM_VSCROLL, 200, 0
'End Sub

Private Sub Command1_Click()
cSearch_Click
End Sub

Private Sub Command2_Click()
cExit_Click
End Sub

Private Sub HDDdetectByWMI_Click()
' Объявляем переменные
'Dim strComputer        ' Имя компьютера
'Dim strNamespace       ' Имя пространства имен
'Dim strClass           ' Имя класса
Dim objClass           ' Объект SWbemObject (класс WMI)
Dim colOperatingSystems ' Коллекция экземпляров класса WMI
Dim objOperatingSystem ' Элемент коллекции
Dim strResult          ' Строка для вывода на экран

Set objClass = GetObject("WinMgmts:\\" & "." & "\" & "Root\CIMV2" & ":" & "Win32_PhysicalMedia")
Set colOperatingSystems = objClass.Instances_
For Each objOperatingSystem In colOperatingSystems
  strResult = strResult & objOperatingSystem.SerialNumber & vbCrLf
Next

Text2.Text = strResult
Text4.Text = Trim(Left$(Text2.Text, 20))

End Sub

Private Sub cOption_Click()
Form2.Show vbModal
End Sub

Sub daymonth()
x = DTPicker1.Month
Dim NeGod As Double
If x = 1 Or x = 3 Or x = 5 Or x = 7 Or x = 8 Or x = 10 Or x = 12 Then maxday = 31
If x = 4 Or x = 6 Or x = 9 Or x = 11 Then maxday = 30

NeGod = DTPicker1.Year / 4  ' проверка на весокосный
If NeGod = CInt(NeGod) Then
    If x = 2 Then maxday = 29
    Else
    If x = 2 Then maxday = 28
End If
End Sub

Private Sub Command6_Click()
Dim qw As Variant
qw = 1
'Text2.SetFocus
'Shish:
'On Error GoTo shish2
Do While InStr(qw, Text2.Text, "теннис", vbTextCompare)
qw = InStr(qw, Text2.Text, "теннис", vbTextCompare)
Text2.SelStart = qw - 1

Text2.SelLength = InStr(qw, Text2.Text, vbCrLf) - Text2.SelStart
Text2.SelColor = vbBlue 'vbGreen  'vbMagenta  '= vbRed
qw = InStr(qw, Text2.Text, "теннис", vbTextCompare) + Text2.SelLength
Loop
'GoTo Shish
'shish2:
Text1.SetFocus
qw = 1
SendMessageLong Me.Text2.hwnd, WM_VSCROLL, SB_LINEUP, 0
End Sub

'Private Sub ScrollUpB_Click()
'SendMessageLong Me.RText.hwnd, WM_VSCROLL, 200, 0
'End Sub

Private Sub Command8_Click()
SendMessageLong Me.Text2.hwnd, WM_VSCROLL, SB_LINEUP, 0

'SendMessageLong Me.RText.hwnd, WM_VSCROLL, SB_LINEDOWN, 0

'SendMessageLong Text2.hwnd, WM_VSCROLL, 200, 0
End Sub
Private Sub Command7_Click()
p = 1
Text2.SetFocus
Shish:
On Error GoTo shish2
p = InStr(p, Text2.Text, "***")
Text2.SelStart = p - 1

Text2.SelLength = InStr(p, Text2.Text, vbCrLf) - Text2.SelStart
Text2.SelColor = vbRed 'vbCyan
p = InStr(p, Text2.Text, "***") + Text2.SelLength
GoTo Shish
shish2:
Text1.SetFocus
p = 1
End Sub


Private Sub DTPicker1_Change()
    'DTPicker1.MaxDate = DTPicker2.Value
    DTPicker2.MinDate = DTPicker1.Value
    DTPicker1.MinDate = "01.01.2000"    ' изменить на продажу
    DTPicker1.MaxDate = DTPicker2.Value
End Sub
Private Sub DTPicker2_Change()
    'DTPicker2.MinDate = DTPicker1.Value
    DTPicker1.MaxDate = DTPicker2.Value
    DTPicker2.MaxDate = "01.01.2010"    ' изменить на продажу
    DTPicker2.MinDate = DTPicker1.Value
End Sub

Private Sub DTPicker1_DblClick()
    DTPicker1.Value = DTPicker2.Value
End Sub

Private Sub DTPicker2_DblClick()
    DTPicker2.Value = DTPicker1.Value
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
Dim f As Variant
Dim vtData() As Byte
Dim vtData2() As Byte

If State = 12 Then
        vtData = Inet1.GetChunk(1024) 'принимаем первую порцию данных

    For f = 0 To UBound(vtData)
    ReDim Preserve vtData2(f)
    vtData2(f) = vtData(f)
    Next f

        Do While LenB(CStr(vtData)) > 0
            
        vtData = Inet1.GetChunk(1024) 'следующая порция данных
    For f = 0 To UBound(vtData)
    ReDim Preserve vtData2(UBound(vtData2) + 1)
    vtData2(UBound(vtData2)) = vtData(f)
    Next f
        Loop
        
alltxt = vtData2
End If


' ' а раньше было (работало), но глючило, когда одновременно по модему через браузер открывались интернет странички
'Dim res As String, a As String
'If State = 12 Then 'если документ полностью получен, то..
'    Do
'        a = Inet1.GetChunk(1024, icByteArray)
'        res = res & a
'    Loop While Len(a) > 0
'    alltxt = res
'End If
End Sub

Private Sub cPersonal_Click()
Dim Response As String
Response = MsgBox(alltxt3, , "Личные встречи")
Text1.SetFocus
End Sub

Private Sub cStop_Click()
stopp = 1
Inet1.Cancel
Label4.Caption = "Загрузка прервана!.."
cStop.Enabled = False
End Sub
Private Sub Form_Load()
p = 1
OptionsLoad
    
    First = 0
LoadGOD
ForSearchGOD
If AllGod = 0 Then
If Len(Dir(CStr(Form1.Option5.Caption) + "\")) Then Form1.Option5.Enabled = True Else Form1.Option5.Enabled = False And Form1.Option4.Value = True
If Len(Dir(CStr(Form1.Option4.Caption) + "\")) Then Form1.Option4.Enabled = True Else Form1.Option4.Enabled = False And Form1.Option3.Value = True
If Len(Dir(CStr(Form1.Option3.Caption) + "\")) Then Form1.Option3.Enabled = True Else Form1.Option3.Enabled = False
vsecheki
Else
Option3.Enabled = False
Option4.Enabled = False
Option5.Enabled = False
For File = 1 To 12
Check1(File).Enabled = False
Next File
End If
    First = 1


DTPicker1.Year = Year(Date)
DTPicker1.Month = Month(Date)
DTPicker1.Day = Day(Date)
DTPicker2.Year = Year(Date)
DTPicker2.Month = Month(Date)
DTPicker2.Day = Day(Date)
DTPicker1.MaxDate = DTPicker2.Value
DTPicker2.MinDate = DTPicker1.Value

Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)

If Raskladka = 1 Then ActivateKeyboardLayout kb_lay_ru, 0 ' смена раскладки клавиатуры на RUS
If CheckSingl = 1 Then Option1.Value = True Else Option2.Value = True

LastData
End Sub

Private Sub Form_Unload(Cancel As Integer)
Module1.OptionsSave
End Sub

Private Sub cExit_Click()
Module1.OptionsSave
Form1.WindowState = 1   ' прежде чем выключиться, окно свернется
Inet1.Cancel
End
End Sub

Private Sub cSearch_Click()
p = 1
alltxt = ""
alltxt2 = ""
alltxt3 = ""                ' проба с персональными встречами
Form1.MousePointer = 13
If Correct = 0 Then
Iskomoe1 = LTrim(Text1.Text) 'Text1.Text '
Iskomoe2 = LTrim(Text3.Text) 'Text3.Text '
Text1.Text = LTrim(Text1.Text)
Text3.Text = LTrim(Text3.Text)
Else
Iskomoe1 = Text1.Text
Iskomoe2 = Text3.Text
End If

                    If AllGod = 0 Then
For File = 1 To 12
If Check1(File).Value = 1 Then
    Open App.Path + "\" + CStr(GOD) + "\" + CStr(File) + ".txt" For Input As #1
    If CheckMonth = 1 Then
    If File = 1 Then
    alltxt = alltxt + "           *** Январь - " + CStr(GOD) + " ***            " + vbCrLf
    alltxt2 = alltxt2 + "           *** Январь - " + CStr(GOD) + " ***            " + vbCrLf
    End If
    If File = 2 Then
    alltxt = alltxt + "           *** Февраль ***            " + vbCrLf
    alltxt2 = alltxt2 + "           *** Февраль ***            " + vbCrLf
    End If
    If File = 3 Then
    alltxt = alltxt + "           *** Март ***            " + vbCrLf
    alltxt2 = alltxt2 + "           *** Март ***            " + vbCrLf
    End If
    If File = 4 Then
    alltxt = alltxt + "           *** Апрель ***            " + vbCrLf
    alltxt2 = alltxt2 + "           *** Апрель ***            " + vbCrLf
    End If
    If File = 5 Then
    alltxt = alltxt + "           *** Май ***            " + vbCrLf
    alltxt2 = alltxt2 + "           *** Май ***            " + vbCrLf
    End If
    If File = 6 Then
    alltxt = alltxt + "           *** Июнь ***            " + vbCrLf
    alltxt2 = alltxt2 + "           *** Июль ***            " + vbCrLf
    End If
    If File = 7 Then
    alltxt = alltxt + "           *** Июль ***            " + vbCrLf
    alltxt2 = alltxt2 + "           *** Июль ***            " + vbCrLf
    End If
    If File = 8 Then
    alltxt = alltxt + "           *** Август ***            " + vbCrLf
    alltxt2 = alltxt2 + "           *** Август ***            " + vbCrLf
    End If
    If File = 9 Then
    alltxt = alltxt + "           *** Сентябрь ***            " + vbCrLf
    alltxt2 = alltxt2 + "           *** Сентябрь ***            " + vbCrLf
    End If
    If File = 10 Then
    alltxt = alltxt + "           *** Октябрь ***            " + vbCrLf
    alltxt2 = alltxt2 + "           *** Октябрь ***            " + vbCrLf
    End If
    If File = 11 Then
    alltxt = alltxt + "           *** Ноябрь ***            " + vbCrLf
    alltxt2 = alltxt2 + "           *** Ноябрь ***            " + vbCrLf
    End If
    If File = 12 Then
    alltxt = alltxt + "           *** Декабрь ***            " + vbCrLf
    alltxt2 = alltxt2 + "           *** Декабрь ***            " + vbCrLf
    End If
    End If
    If CheckTUR = 1 Then
    If Correct = 0 Then Search2 Else Search2correct
    Else: Search1
    End If
    Close #1
End If
Next File
                    Else
For GOD = 2000 To 2010
If Len(Dir(CStr(GOD) + "\")) Then
For File = 1 To 12
    
    If Len(Dir(CStr(GOD) + "\" + CStr(File) + ".txt")) Then
    Open App.Path + "\" + CStr(GOD) + "\" + CStr(File) + ".txt" For Input As #1
    If CheckMonth = 1 Then
    If File = 1 Then
    alltxt = alltxt + "           *** Январь - " + CStr(GOD) + " ***            " + vbCrLf
    alltxt2 = alltxt2 + "           *** Январь - " + CStr(GOD) + " ***            " + vbCrLf
    End If
    If File = 2 Then
    alltxt = alltxt + "           *** Февраль ***            " + vbCrLf
    alltxt2 = alltxt2 + "           *** Февраль ***            " + vbCrLf
    End If
    If File = 3 Then
    alltxt = alltxt + "           *** Март ***            " + vbCrLf
    alltxt2 = alltxt2 + "           *** Март ***            " + vbCrLf
    End If
    If File = 4 Then
    alltxt = alltxt + "           *** Апрель ***            " + vbCrLf
    alltxt2 = alltxt2 + "           *** Апрель ***            " + vbCrLf
    End If
    If File = 5 Then
    alltxt = alltxt + "           *** Май ***            " + vbCrLf
    alltxt2 = alltxt2 + "           *** Май ***            " + vbCrLf
    End If
    If File = 6 Then
    alltxt = alltxt + "           *** Июнь ***            " + vbCrLf
    alltxt2 = alltxt2 + "           *** Июль ***            " + vbCrLf
    End If
    If File = 7 Then
    alltxt = alltxt + "           *** Июль ***            " + vbCrLf
    alltxt2 = alltxt2 + "           *** Июль ***            " + vbCrLf
    End If
    If File = 8 Then
    alltxt = alltxt + "           *** Август ***            " + vbCrLf
    alltxt2 = alltxt2 + "           *** Август ***            " + vbCrLf
    End If
    If File = 9 Then
    alltxt = alltxt + "           *** Сентябрь ***            " + vbCrLf
    alltxt2 = alltxt2 + "           *** Сентябрь ***            " + vbCrLf
    End If
    If File = 10 Then
    alltxt = alltxt + "           *** Октябрь ***            " + vbCrLf
    alltxt2 = alltxt2 + "           *** Октябрь ***            " + vbCrLf
    End If
    If File = 11 Then
    alltxt = alltxt + "           *** Ноябрь ***            " + vbCrLf
    alltxt2 = alltxt2 + "           *** Ноябрь ***            " + vbCrLf
    End If
    If File = 12 Then
    alltxt = alltxt + "           *** Декабрь ***            " + vbCrLf
    alltxt2 = alltxt2 + "           *** Декабрь ***            " + vbCrLf
    End If
    End If
    If CheckTUR = 1 Then
    If Correct = 0 Then Search2 Else Search2correct
    Else: Search1
    End If
    'If Check3.Value = 1 Then Search2 Else Search
    Close #1
End If
Next File
End If
Next GOD
                    End If
If Option3.Value = True Then GOD = Option3.Caption
If Option4.Value = True Then GOD = Option4.Caption
If Option5.Value = True Then GOD = Option5.Caption
'If Len(Dir(CStr(Form1.Option5.Caption) + "\")) Then Form1.Option5.Enabled = True Else Form1.Option5.Enabled = False And Form1.Option4.Value = True
'If Len(Dir(CStr(Form1.Option4.Caption) + "\")) Then Form1.Option4.Enabled = True Else Form1.Option4.Enabled = False And Form1.Option3.Value = True
'If Len(Dir(CStr(Form1.Option3.Caption) + "\")) Then Form1.Option3.Enabled = True Else Form1.Option3.Enabled = False

'Label4.ForeColor = &H80000002

Text2.Text = alltxt
Text4.Text = alltxt2
'Text7.Text = alltxt3        'проба

Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)
If alltxt3 <> "" Then
'If Text2.Text <> "" And Text4.Text <> "" Then
cPersonal.Enabled = True
Else
cPersonal.Enabled = False
End If
Form1.MousePointer = 0
Text1.SetFocus
End Sub
Sub Search1()
Dim neprav As Integer
neprav = 1
Do Until EOF(1)
Line Input #1, txt

If InStr(p, txt, ".20") <> 0 Then   ' защита года (дату в поиске не покажет)
    If InStr(p, txt, GOD) <> 0 Then neprav = 0 Else neprav = 1
Else
    If neprav = 0 Then  ' защита года

'If Len(Trim(Text1.Text)) > 1 Then Text2.Text = "cool" Else Text2.Text = ""
If InStr(p, txt, "теннис", vbTextCompare) = 0 Then

                        If Correct = 0 Then
If Len(Trim(Text1.Text)) > 1 Then
If InStr(p, txt, Iskomoe1, vbTextCompare) <> 0 Then
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
If Len(Trim(Text3.Text)) > 1 Then
If InStr(p, txt, Iskomoe2, vbTextCompare) <> 0 Then
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
If Text1.Text <> "" And Text3.Text <> "" Then
If InStr(p, txt, Iskomoe1, vbTextCompare) <> 0 Then   ' для личных встреч
If InStr(p, txt, Iskomoe2, vbTextCompare) <> 0 Then
    
    If Option1.Value = True Then        ' одиночная игра
        If InStr(p, txt, "/") <> 0 Then
        Else
        alltxt3 = alltxt3 + txt + vbCrLf
        End If
    End If
    If Option2.Value = True Then        ' парная/микст
        If InStr(p, txt, "/") <> 0 Then
        alltxt3 = alltxt3 + txt + vbCrLf
        End If
    End If

End If
End If
End If
'-------------------------------------------------------------
                        Else
If Len(Trim(Text1.Text)) > 1 Then
If InStr(p, txt, Iskomoe1) <> 0 Then
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
If Len(Trim(Text3.Text)) > 1 Then
If InStr(p, txt, Iskomoe2) <> 0 Then
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
If Text1.Text <> "" And Text3.Text <> "" Then
If InStr(p, txt, Iskomoe1) <> 0 Then   ' для личных встреч
If InStr(p, txt, Iskomoe2) <> 0 Then
    
    If Option1.Value = True Then        ' одиночная игра
        If InStr(p, txt, "/") <> 0 Then
        Else
        alltxt3 = alltxt3 + txt + vbCrLf
        End If
    End If
    If Option2.Value = True Then        ' парная/микст
        If InStr(p, txt, "/") <> 0 Then
        alltxt3 = alltxt3 + txt + vbCrLf
        End If
    End If

End If
End If
End If
'-------------------------------------------------------------
End If

End If
End If ' защита года
End If ' защита года
Loop
End Sub

Sub Search2()
Dim neprav As Integer
Dim ok As Integer
Tur1 = ""
Tur2 = ""
Tur3 = ""
Tur4 = ""
ok = 1
neprav = 1

Do Until EOF(1)
Line Input #1, txt

If InStr(p, txt, ".20") <> 0 Then   ' защита года (дату в поиске не покажет)
    If InStr(p, txt, GOD) <> 0 Then neprav = 0 Else neprav = 1
Else
    If neprav = 0 Then  ' защита года



If InStr(p, txt, "теннис", vbTextCompare) <> 0 Then
    
    If noncover = 1 Then
        
    If InStr(p, txt, " Ковер") <> 0 And corpet = 0 Then
    ok = 0
    GoTo esliOK
    Else: ok = 1
    End If
    If InStr(p, txt, " Хард") <> 0 And hard = 0 Then
    ok = 0
    GoTo esliOK
    Else: ok = 1
    End If
    If InStr(p, txt, " Грунт") <> 0 And soil = 0 Then
    ok = 0
    GoTo esliOK
    Else: ok = 1
    End If
    If InStr(p, txt, " Трава") <> 0 And grass = 0 Then
    ok = 0
    GoTo esliOK
    Else: ok = 1
    End If
    
    Else
    
    If InStr(p, txt, " Ковер") <> 0 And corpet = 1 Then
    ok = 1
    Tur2 = txt
    GoTo esliOK
    Else: ok = 0
    End If
    If InStr(p, txt, " Хард") <> 0 And hard = 1 Then
    ok = 1
    Tur2 = txt
    GoTo esliOK
    Else: ok = 0
    End If
    If InStr(p, txt, " Грунт") <> 0 And soil = 1 Then
    ok = 1
    Tur2 = txt
    GoTo esliOK
    Else: ok = 0
    End If
    If InStr(p, txt, " Трава") <> 0 And grass = 1 Then
    ok = 1
    Tur2 = txt
    GoTo esliOK
    Else: ok = 0
    End If
    If noncover = 1 Then
    ok = 1
    Tur2 = txt
    GoTo esliOK
    Else
    ok = 0
    GoTo esliOK
    End If
    End If
    
    Tur2 = txt
Else
esliOK:

If ok = 1 Then
If Len(Trim(Text1.Text)) > 1 Then
If InStr(p, txt, Iskomoe1, vbTextCompare) <> 0 Then
    If Option1.Value = True Then        ' одиночная игра
        If InStr(p, txt, "/") <> 0 Then
        Else
            If Tur1 = Tur2 Then
                alltxt = alltxt + txt + vbCrLf
            Else
                alltxt = alltxt + Tur2 + vbCrLf + txt + vbCrLf
                Tur1 = Tur2
            End If
        'alltxt = alltxt + txt + vbCrLf
        End If
    End If
    If Option2.Value = True Then        ' парная/микст
        If InStr(p, txt, "/") <> 0 Then
            If Tur1 = Tur2 Then
                alltxt = alltxt + txt + vbCrLf
            Else
                alltxt = alltxt + Tur2 + vbCrLf + txt + vbCrLf
                Tur1 = Tur2
            End If
        'alltxt = alltxt + txt + vbCrLf
        End If
    End If
End If
End If


If Len(Trim(Text3.Text)) > 1 Then
If InStr(p, txt, Iskomoe2, vbTextCompare) <> 0 Then
    If Option1.Value = True Then        ' одиночная игра
        If InStr(p, txt, "/") <> 0 Then
        Else
            If Tur3 = Tur2 Then
                alltxt2 = alltxt2 + txt + vbCrLf
            Else
                alltxt2 = alltxt2 + Tur2 + vbCrLf + txt + vbCrLf
                Tur3 = Tur2
            End If
        'alltxt2 = alltxt2 + txt + vbCrLf
        End If
    End If
    If Option2.Value = True Then        ' парная/микст
        If InStr(p, txt, "/") <> 0 Then
            If Tur3 = Tur2 Then
                alltxt2 = alltxt2 + txt + vbCrLf
            Else
                alltxt2 = alltxt2 + Tur2 + vbCrLf + txt + vbCrLf
                Tur3 = Tur2
            End If
        'alltxt2 = alltxt2 + txt + vbCrLf
        End If
    End If
End If
End If


'--------------------------------------------------------------
If Text1.Text <> "" And Text3.Text <> "" Then
If InStr(p, txt, Iskomoe1, vbTextCompare) <> 0 Then   ' для личных встреч
If InStr(p, txt, Iskomoe2, vbTextCompare) <> 0 Then
    
    If Option1.Value = True Then        ' одиночная игра
        If InStr(p, txt, "/") <> 0 Then
        Else
            If Tur4 = Tur2 Then
                alltxt3 = alltxt3 + txt + vbCrLf
            Else
                alltxt3 = alltxt3 + Tur2 + vbCrLf + txt + vbCrLf
                Tur4 = Tur2
            End If
        End If
    End If
    If Option2.Value = True Then        ' парная/микст
        If InStr(p, txt, "/") <> 0 Then
            If Tur4 = Tur2 Then
                alltxt3 = alltxt3 + txt + vbCrLf
            Else
                alltxt3 = alltxt3 + Tur2 + vbCrLf + txt + vbCrLf
                Tur4 = Tur2
            End If
        End If
    End If

End If
End If
End If

End If ' OK и прочее связанное с покрытием
End If  ' если "теннис" нашел/нет
'-------------------------------------------------------------
End If ' защита года
End If ' защита года
Loop
End Sub
Sub Search2correct()
Dim neprav As Integer
Dim ok As Integer
Tur1 = ""
Tur2 = ""
Tur3 = ""
Tur4 = ""
ok = 1
neprav = 1

Do Until EOF(1)
Line Input #1, txt

If InStr(p, txt, ".20") <> 0 Then   ' защита года (дату в поиске не покажет)
    If InStr(p, txt, GOD) <> 0 Then neprav = 0 Else neprav = 1
Else
    If neprav = 0 Then  ' защита года



If InStr(p, txt, "теннис", vbTextCompare) <> 0 Then
    
    If noncover = 1 Then
        
    If InStr(p, txt, " Ковер") <> 0 And corpet = 0 Then
    ok = 0
    GoTo esliOK
    Else: ok = 1
    End If
    If InStr(p, txt, " Хард") <> 0 And hard = 0 Then
    ok = 0
    GoTo esliOK
    Else: ok = 1
    End If
    If InStr(p, txt, " Грунт") <> 0 And soil = 0 Then
    ok = 0
    GoTo esliOK
    Else: ok = 1
    End If
    If InStr(p, txt, " Трава") <> 0 And grass = 0 Then
    ok = 0
    GoTo esliOK
    Else: ok = 1
    End If
    
    Else
    
    If InStr(p, txt, " Ковер") <> 0 And corpet = 1 Then
    ok = 1
    Tur2 = txt
    GoTo esliOK
    Else: ok = 0
    End If
    If InStr(p, txt, " Хард") <> 0 And hard = 1 Then
    ok = 1
    Tur2 = txt
    GoTo esliOK
    Else: ok = 0
    End If
    If InStr(p, txt, " Грунт") <> 0 And soil = 1 Then
    ok = 1
    Tur2 = txt
    GoTo esliOK
    Else: ok = 0
    End If
    If InStr(p, txt, " Трава") <> 0 And grass = 1 Then
    ok = 1
    Tur2 = txt
    GoTo esliOK
    Else: ok = 0
    End If
    If noncover = 1 Then
    ok = 1
    Tur2 = txt
    GoTo esliOK
    Else
    ok = 0
    GoTo esliOK
    End If
    End If
    
    Tur2 = txt
Else
esliOK:

If ok = 1 Then
If Len(Trim(Text1.Text)) > 1 Then
If InStr(p, txt, Iskomoe1) <> 0 Then
    If Option1.Value = True Then        ' одиночная игра
        If InStr(p, txt, "/") <> 0 Then
        Else
            If Tur1 = Tur2 Then
                alltxt = alltxt + txt + vbCrLf
            Else
                alltxt = alltxt + Tur2 + vbCrLf + txt + vbCrLf
                Tur1 = Tur2
            End If
        End If
    End If
    If Option2.Value = True Then        ' парная/микст
        If InStr(p, txt, "/") <> 0 Then
            If Tur1 = Tur2 Then
                alltxt = alltxt + txt + vbCrLf
            Else
                alltxt = alltxt + Tur2 + vbCrLf + txt + vbCrLf
                Tur1 = Tur2
            End If
        End If
    End If
End If
End If


If Len(Trim(Text3.Text)) > 1 Then
If InStr(p, txt, Iskomoe2) <> 0 Then
    If Option1.Value = True Then        ' одиночная игра
        If InStr(p, txt, "/") <> 0 Then
        Else
            If Tur3 = Tur2 Then
                alltxt2 = alltxt2 + txt + vbCrLf
            Else
                alltxt2 = alltxt2 + Tur2 + vbCrLf + txt + vbCrLf
                Tur3 = Tur2
            End If
        End If
    End If
    If Option2.Value = True Then        ' парная/микст
        If InStr(p, txt, "/") <> 0 Then
            If Tur3 = Tur2 Then
                alltxt2 = alltxt2 + txt + vbCrLf
            Else
                alltxt2 = alltxt2 + Tur2 + vbCrLf + txt + vbCrLf
                Tur3 = Tur2
            End If
        End If
    End If
End If
End If

'--------------------------------------------------------------
If Text1.Text <> "" And Text3.Text <> "" Then
If InStr(p, txt, Iskomoe1) <> 0 Then   ' для личных встреч
If InStr(p, txt, Iskomoe2) <> 0 Then
    
    If Option1.Value = True Then        ' одиночная игра
        If InStr(p, txt, "/") <> 0 Then
        Else
            If Tur4 = Tur2 Then
                alltxt3 = alltxt3 + txt + vbCrLf
            Else
                alltxt3 = alltxt3 + Tur2 + vbCrLf + txt + vbCrLf
                Tur4 = Tur2
            End If
        End If
    End If
    If Option2.Value = True Then        ' парная/микст
        If InStr(p, txt, "/") <> 0 Then
            If Tur4 = Tur2 Then
                alltxt3 = alltxt3 + txt + vbCrLf
            Else
                alltxt3 = alltxt3 + Tur2 + vbCrLf + txt + vbCrLf
                Tur4 = Tur2
            End If
        End If
    End If

End If
End If
End If

End If ' OK и прочее связанное с покрытием
End If  ' если "теннис" нашел/нет
'-------------------------------------------------------------
End If ' защита года
End If ' защита года
Loop
End Sub
'****************************************************************************************
Sub MarkLoad()
    PerSent = 100 / ProgressBar1.Max + PerSent
    Form1.Caption = "TenScore 2005" & " (" & CInt(PerSent) & "% - загружено)"
    Label4.Caption = "Загрузка данных - " & CInt(PerSent) & "%"
End Sub

Private Sub cLoad_Click()
Dim Raz As Integer
Text5.Text = ""
alltxt1 = ""
'If DTPicker1.Year = GOD And DTPicker2.Year = GOD Then   ' эта и 3 снизу!!!
stopp = 0
PerSent = 0
Label4.ForeColor = &H80000002
Label4.Caption = "Идет загрузка данных..."

On Error GoTo Connless

ProgressBar1.Max = DTPicker2.Value - DTPicker1.Value + 1 'PGBar
cStop.Enabled = True
cLoad.Enabled = False
DTPicker1.Enabled = False
DTPicker2.Enabled = False

Command1.Enabled = False
cSearch.Enabled = False
If cPersonal.Enabled = True Then Pers = 1 Else Pers = 0
cPersonal.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
Option5.Enabled = False
cOption.Enabled = False

Form1.MousePointer = 13

If DTPicker1.Year = DTPicker2.Year Then
If DTPicker1.Month < DTPicker2.Month Then

    For m = DTPicker1.Month To DTPicker2.Month - 1
    daymonth
    LoadOldfile      ' даст данные в alltxt2 и alltxt3
    For d = DTPicker1.Day To maxday
    LoadInform       ' загрузка из InterNet
    TextCorrector
    If alltxt <> "" Then
    alltxt1 = alltxt1 & DTPicker1.Day & "." & DTPicker1.Month & "." & DTPicker1.Year _
    & vbCrLf & alltxt & vbCrLf
    End If
        If stopp = 1 Then GoTo Kones
    ProgressBar1.Value = ProgressBar1.Value + 1
    MarkLoad
    If DTPicker1.Day = maxday Then GoTo mainerror ' путь в мертвую зону: m = 50
    DTPicker1.Day = DTPicker1.Day + 1  ' счетчик по дням

If m = 50 Then
mainerror:
Zapis
alltxt1 = ""
DTPicker1.Day = 1
DTPicker1.Month = DTPicker1.Month + 1

End If
    Next d
    Next m
End If
End If

'******************************************************************************************************
'***************Text2.Text = (god2 - god1 - 1) * 12 + 12 - mes1 + mes2 + 1*****************************
'******************************************************************************************************
If DTPicker1.Year <> DTPicker2.Year Then
    For Raz = 1 To (DTPicker2.Year - DTPicker1.Year - 1) * 12 + 12 - DTPicker1.Month + DTPicker2.Month
    'For m = DTPicker1.Month To DTPicker2.Month - 1
    daymonth
    LoadOldfile      ' даст данные в alltxt2 и alltxt3
    For d = DTPicker1.Day To maxday
    LoadInform       ' загрузка из InterNet
    TextCorrector
    If alltxt <> "" Then
    alltxt1 = alltxt1 & DTPicker1.Day & "." & DTPicker1.Month & "." & DTPicker1.Year _
    & vbCrLf & alltxt & vbCrLf
    End If
        If stopp = 1 Then GoTo Kones
    ProgressBar1.Value = ProgressBar1.Value + 1
    MarkLoad
    If DTPicker1.Day = maxday Then GoTo mainerror2 ' путь в мертвую зону: m = 50
    DTPicker1.Day = DTPicker1.Day + 1

    If m = 50 Then
mainerror2:
Zapis
    alltxt1 = ""
    DTPicker1.Day = 1
    If DTPicker1.Month = 12 Then
        DTPicker1.Month = 1
        DTPicker1.Year = DTPicker1.Year + 1
    Else
    DTPicker1.Month = DTPicker1.Month + 1
    End If

End If
    Next d
    Next Raz
End If
'******************************************************************************************************
'***************Text2.Text = (god2 - god1 - 1) * 12 + 12 - mes1 + mes2 + 1*****************************
'******************************************************************************************************

'------------------------------------------------------ деление несколько месяцев/один месяц
If DTPicker1.Month = DTPicker2.Month Then
'PGBar
    'daymonth
    LoadOldfile      ' даст данные в alltxt2 и alltxt3
If DTPicker1.Day < DTPicker2.Day Then
    For d = DTPicker1.Day To DTPicker2.Day - 1
    LoadInform       ' загрузка из InterNet
    TextCorrector
    If alltxt <> "" Then
    alltxt1 = alltxt1 & DTPicker1.Day & "." & DTPicker1.Month & "." & DTPicker1.Year _
    & vbCrLf & alltxt & vbCrLf
    End If
        If stopp = 1 Then GoTo Kones
    DTPicker1.Day = DTPicker1.Day + 1  ' счетчик по дням
    ProgressBar1.Value = ProgressBar1.Value + 1
    MarkLoad
    
    Next d
End If


If DTPicker1.Day = DTPicker2.Day Then
    LoadInform       ' загрузка из InterNet
    TextCorrector
    If alltxt <> "" Then
    alltxt1 = alltxt1 & DTPicker1.Day & "." & DTPicker1.Month & "." & DTPicker1.Year _
    & vbCrLf & alltxt & vbCrLf
    End If
        If stopp = 1 Then GoTo Kones
    ProgressBar1.Value = ProgressBar1.Value + 1
    MarkLoad

Kones:
Zapis
alltxt1 = ""
LastData 'If stopp = 0 Then Label4.Caption = "Загрузка данных завершена успешно..."
End If
alltxt = ""
End If
'******************************************************************************************************
'***************Text2.Text = (god2 - god1 - 1) * 12 + 12 - mes1 + mes2 + 1*****************************
'******************************************************************************************************

ProgressBar1.Value = 0
If m = 50 Then

Connless:
Inet1.Cancel
Label4.ForeColor = &HFF&
Label4.Caption = "Связь отсутствует!!!"
End If

DTPicker1.Enabled = True
DTPicker2.Enabled = True
Command1.Enabled = True
cSearch.Enabled = True
If Pers = 1 Then cPersonal.Enabled = True Else cPersonal.Enabled = False
cOption.Enabled = True
If Len(Dir(CStr(Form1.Option5.Caption) + "\")) Then Form1.Option5.Enabled = True Else Form1.Option5.Enabled = False
If Len(Dir(CStr(Form1.Option4.Caption) + "\")) Then Form1.Option4.Enabled = True Else Form1.Option4.Enabled = False
If Len(Dir(CStr(Form1.Option3.Caption) + "\")) Then Form1.Option3.Enabled = True Else Form1.Option3.Enabled = False
vsecheki
cStop.Enabled = False
cLoad.Enabled = True
Form1.MousePointer = 0
Form1.Caption = "TenScore 2005"
'Else
'Response = MsgBox("Программа TenScore " & GOD & " расчитана только на " & GOD & " год. " + vbCrLf + vbCrLf + "Если Вам понравилась программа TenScore и Вы хотите иметь Полную Версию на следующий года," + vbCrLf + "то закажите ее на сайте http://www.tenscore.narod.ru/", , "TenScore")
'End If
End Sub

Sub LoadInform()
'http://www.marathonbet.com/results.shtml?tr3=1&detal=1&day=29&month=10&year=2005&dayTo=29&monthTo=10&yearTo=2005
Inet1.Execute "http://www.marathonbet.com/results.shtml?tr3=1&detal=1&day=" & DTPicker1.Day & "&month=" & DTPicker1.Month & "&year=" & DTPicker1.Year & "&dayTo=" & DTPicker1.Day & "&monthTo=" & DTPicker1.Month & "&yearTo=" & DTPicker1.Year, "GET"
'Inet1.Execute "http://www.marathonbet.com/results.shtml?day=" & DTPicker1.Day & "&month=" & DTPicker1.Month & "&year=" & DTPicker1.Year & "&tr=-1", "GET"
Do While Inet1.StillExecuting: DoEvents: Loop ' кольцо для нескольких дней запроса (на выделенке глюк)
End Sub

Sub TextCorrector()
alltxt = ReplaceStr(alltxt, "Теннис", vbNullString, vbTextCompare)
If alltxt <> "" Then
alltxt = ReplaceStr2(alltxt, "Триатлон", vbNullString, vbTextCompare)
alltxt = ReplaceStr2(alltxt, "Тяжелая атлетика", vbNullString, vbTextCompare)
alltxt = ReplaceStr2(alltxt, "Фехтование", vbNullString, vbTextCompare)
alltxt = ReplaceStr2(alltxt, "Формула", vbNullString, vbTextCompare)
alltxt = ReplaceStr2(alltxt, "Футбол", vbNullString, vbTextCompare)
alltxt = ReplaceStr2(alltxt, "Хоккей", vbNullString, vbTextCompare)
alltxt = ReplaceStr2(alltxt, "теннису отсутствуют.</b></font></pre></div>", vbNullString, vbTextCompare)
alltxt = ReplaceStr2(alltxt, "</pre></div>", vbNullString, vbTextCompare)

alltxt = ReplaceStr3(alltxt, "</span><pre>", vbNullString, vbBinaryCompare)
alltxt = ReplaceStr3(alltxt, "</pre><span class=cap>", vbNullString, vbBinaryCompare)
alltxt = ReplaceStr3(alltxt, "&nbsp;", vbNullString, vbBinaryCompare)

alltxt = ReplaceStr3(alltxt, Chr$(13), vbCrLf, vbBinaryCompare)

alltxt = ReplaceStr4(alltxt, "Матч отменен", vbNullString, vbTextCompare)
alltxt = ReplaceStr5(alltxt, "(По", vbNullString, vbTextCompare) ' (По", vbTextCompare "(в" ( По
alltxt = ReplaceStr5(alltxt, "(в", vbNullString, vbTextCompare)
alltxt = ReplaceStr5(alltxt, "( По", vbNullString, vbTextCompare)
alltxt = ReplaceStr5(alltxt, "(и", vbNullString, vbTextCompare)
alltxt = ReplaceStr5(alltxt, "(н", vbNullString, vbTextCompare)  '19.8.2002 - Энквист - Выплата с коэфф. 1 (не участвовал в турнире)
End If
End Sub

Public Function ReplaceStr(ByVal strString As String, ByVal strReplace As String, _
    Optional ByVal strReplaceWith As String = vbNullString, _
    Optional CompareMethod As VbCompareMethod) As String
    On Error Resume Next
    Dim iLenOut As Integer, iLenIn As Integer
    Dim i As Long
    iLenOut = Len(strReplace)
    iLenIn = Len(strReplaceWith)
    If Len(strString) > 0 Then
        If iLenOut > 0 Then
        If InStr(InStr(1, strString, strReplace, CompareMethod) + iLenOut, strString, strReplace, CompareMethod) Then
        i = InStr(InStr(1, strString, strReplace, CompareMethod) + iLenOut, strString, strReplace, CompareMethod)
        strString = Mid$(strString, i)
        Else:
        strString = ""
        alltxt = ""
        Exit Function
        End If
        End If
    End If
    ReplaceStr = strString
End Function
Public Function ReplaceStr2(ByVal strString As String, ByVal strReplace As String, _
    Optional ByVal strReplaceWith As String = vbNullString, _
    Optional CompareMethod As VbCompareMethod) As String
    On Error Resume Next
    Dim iLenOut As Integer, iLenIn As Integer
    Dim i As Long
    i = 0
    iLenOut = Len(strReplace)
    iLenIn = Len(strReplaceWith)
    If Len(strString) > 0 Then
        If iLenOut > 0 Then
            i = InStr(1, strString, strReplace, CompareMethod)
            strString = Left$(strString, i - 1)
        End If
    End If
    ReplaceStr2 = strString
End Function
Public Function ReplaceStr3(ByVal strString As String, ByVal strReplace As String, _
    Optional ByVal strReplaceWith As String = vbNullString, _
    Optional CompareMethod As VbCompareMethod) As String
        On Error Resume Next
    Dim iLenOut As Integer, iLenIn As Integer
    Dim i As Long
    iLenOut = Len(strReplace)
    iLenIn = Len(strReplaceWith)
    If Len(strString) > 0 Then
        If iLenOut > 0 Then
            i = InStr(1, strString, strReplace, CompareMethod)
            Do Until i = 0
                If iLenIn > 0 Then
                    strString = Left$(strString, i - 1) & strReplaceWith & Mid$(strString, i + iLenOut)
                Else
                    strString = Left$(strString, i - 1) & Mid$(strString, i + iLenOut)
                End If
                i = InStr(i + iLenIn, strString, strReplace, CompareMethod)
            Loop
        End If
    End If
    ReplaceStr3 = strString
End Function
Public Function ReplaceStr32(ByVal strString As String, ByVal strReplace As String, _
    Optional ByVal strReplaceWith As String = vbNullString, _
    Optional CompareMethod As VbCompareMethod) As String
        On Error Resume Next
    Dim iLenOut As Integer, iLenIn As Integer
    Dim i As Long
    iLenOut = Len(strReplace)
    iLenIn = Len(strReplaceWith)
    If Len(strString) > 0 Then
        If iLenOut > 0 Then
            i = InStr(1, strString, strReplace, CompareMethod)
            strString = Left$(strString, i - 1) & strReplaceWith & Mid$(strString, i + iLenOut)
            Do Until i = 0
                If iLenIn > 0 Then
                    strString = Left$(strString, i - 1) & strReplaceWith & Mid$(strString, i + iLenOut)
                Else
                    strString = Left$(strString, i - 1) & Mid$(strString, i + iLenOut)
                End If
                i = InStr(i + iLenIn, strString, strReplace, CompareMethod)
            'Text2.Text = ReplaceStr(Text2.Text, Chr$(127), vbCrLf, vbBinaryCompare)
            'Mid(MyString, 3, 7) = "пятницу" ' MyString = "В пятницу утром".
            Loop
        End If
    End If
    ReplaceStr32 = strString
End Function
Public Function ReplaceStr4(ByVal strString As String, ByVal strReplace As String, _
    Optional ByVal strReplaceWith As String = vbNullString, _
    Optional CompareMethod As VbCompareMethod) As String
        On Error Resume Next
    Dim iLenOut As Integer, iLenIn As Integer
    Dim i As Long
    Dim n As Long
    iLenOut = Len(strReplace)
    iLenIn = Len(strReplaceWith)
    If Len(strString) > 0 Then
        If iLenOut > 0 Then
            i = InStr(1, strString, strReplace, CompareMethod)
            n = InStr(i, strString, vbCrLf, CompareMethod)
            Do Until i = 0
                If iLenIn > 0 Then
                    strString = Left$(strString, i + 11) & strReplaceWith & Mid$(strString, n)
                Else
                    strString = Left$(strString, i + 11) & Mid$(strString, n)
                End If
                i = InStr(i + iLenIn + 2, strString, strReplace, CompareMethod)
                n = InStr(i, strString, vbCrLf, CompareMethod)
            Loop
        End If
    End If
    ReplaceStr4 = strString
End Function
    
Public Function ReplaceStr5(ByVal strString As String, ByVal strReplace As String, _
    Optional ByVal strReplaceWith As String = vbNullString, _
    Optional CompareMethod As VbCompareMethod) As String
        On Error Resume Next
    Dim iLenOut As Integer, iLenIn As Integer
    Dim i As Long
    Dim n As Long
    iLenOut = Len(strReplace)
    iLenIn = Len(strReplaceWith)
    If Len(strString) > 0 Then
        If iLenOut > 0 Then
            i = InStr(1, strString, strReplace, CompareMethod)
            n = InStr(i, strString, vbCrLf, CompareMethod)
            Do Until i = 0
                If iLenIn > 0 Then
                    strString = Left$(strString, i - 1) & strReplaceWith & Mid$(strString, n)
                Else
                    strString = Left$(strString, i - 1) & Mid$(strString, n)
                End If
                i = InStr(i + iLenIn, strString, strReplace, CompareMethod)
                n = InStr(i, strString, vbCrLf, CompareMethod)
            Loop
        End If
    End If
    ReplaceStr5 = strString
End Function

Sub LoadOldfile()
On Error Resume Next
    MkDir App.Path + "\" + CStr(DTPicker1.Year)
On Error GoTo error1
    If p = 50 Then
error1:
    Open App.Path + "\" + CStr(DTPicker1.Year) + "\" + CStr(DTPicker1.Month) + ".txt" For Output As #1
    Close #1
    End If
DoEvents:
    Open App.Path + "\" + CStr(DTPicker1.Year) + "\" + CStr(DTPicker1.Month) + ".txt" For Input As #1
    DevisionOldFileOnTwoParts
    Close #1
End Sub

Sub DevisionOldFileOnTwoParts()       ' разборка по датам старый файл
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
If q = 1 Then ' вторая часть alltxt3
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
End Sub

Sub Zapis()
For File = 1 To 12
If DTPicker1.Month = File Then
    Open App.Path + "\" + CStr(DTPicker1.Year) + "\" + CStr(File) + ".txt" For Output As #1
    Print #1, alltxt2
    Print #1, alltxt1
    Print #1, alltxt3
    Close #1
End If
Next File
End Sub

Private Sub Check1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then _
  ChangeTo = 1 - Check1(Index).Value: _
  Check1(Index).Value = ChangeTo: _
  ReleaseCapture
End Sub

Private Sub Check1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then Check1(Index).Value = ChangeTo
End Sub

Private Sub Image1_Click()
ShellExecute Me.hwnd, vbNullString, Linklabel, vbNullString, vbNullString, SW_SHOWNORMAL
End Sub

Private Sub Label5_Click()
ShellExecute Me.hwnd, vbNullString, "http://www.marathonbet.com/tab_2t.shtml?razd=-2", vbNullString, vbNullString, SW_SHOWNORMAL
Label5.ForeColor = &HFF0000
Label5.FontUnderline = False
End Sub

Private Sub Label7_Click()
ShellExecute Me.hwnd, vbNullString, "http://www.marathonbet.com/tab_2t.shtml?razd=-1", vbNullString, vbNullString, SW_SHOWNORMAL
Label7.ForeColor = &HFF0000
Label7.FontUnderline = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label7.ForeColor = &HFF0000
Label7.FontUnderline = False
Label5.ForeColor = &HFF0000
Label5.FontUnderline = False
End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label7.ForeColor = &HFF0000
Label7.FontUnderline = False
Label5.ForeColor = &HFF0000
Label5.FontUnderline = False
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label7.ForeColor = &HFF&
Label7.FontUnderline = True
Label5.ForeColor = &HFF0000
Label5.FontUnderline = False
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label5.ForeColor = &HFF&
Label5.FontUnderline = True
Label7.ForeColor = &HFF0000
Label7.FontUnderline = False
End Sub

Private Sub Option3_Click()
    If First = 1 Then SaveGOD
GOD = Option3.Caption
LoadGOD
    If First = 1 Then vsecheki '2003
    If First = 1 Then LastData
If Len(Dir(CStr(Form1.Option3.Caption) + "\")) Then Form1.Option3.Enabled = True Else Form1.Option3.Enabled = False
If Len(Dir(CStr(Form1.Option4.Caption) + "\")) Then Form1.Option4.Enabled = True Else Form1.Option4.Enabled = False
If Len(Dir(CStr(Form1.Option5.Caption) + "\")) Then Form1.Option5.Enabled = True Else Form1.Option5.Enabled = False
End Sub

Private Sub Option4_Click()
    If First = 1 Then SaveGOD
GOD = Option4.Caption
LoadGOD
    If First = 1 Then vsecheki '2004
    If First = 1 Then LastData
If Len(Dir(CStr(Form1.Option3.Caption) + "\")) Then Form1.Option3.Enabled = True Else Form1.Option3.Enabled = False
If Len(Dir(CStr(Form1.Option4.Caption) + "\")) Then Form1.Option4.Enabled = True Else Form1.Option4.Enabled = False
If Len(Dir(CStr(Form1.Option5.Caption) + "\")) Then Form1.Option5.Enabled = True Else Form1.Option5.Enabled = False
End Sub

Private Sub Option5_Click()
    If First = 1 Then SaveGOD
GOD = Option5.Caption
LoadGOD
    If First = 1 Then vsecheki '2005
    If First = 1 Then LastData
If Len(Dir(CStr(Form1.Option3.Caption) + "\")) Then Form1.Option3.Enabled = True Else Form1.Option3.Enabled = False
If Len(Dir(CStr(Form1.Option4.Caption) + "\")) Then Form1.Option4.Enabled = True Else Form1.Option4.Enabled = False
If Len(Dir(CStr(Form1.Option5.Caption) + "\")) Then Form1.Option5.Enabled = True Else Form1.Option5.Enabled = False
End Sub



