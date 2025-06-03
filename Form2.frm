VERSION 5.00
Object = "{A51095D7-8D17-11D6-9913-E1D1DF4BFD40}#1.0#0"; "XPButton.ocx"
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " TenScore - Опции"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6930
   ControlBox      =   0   'False
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Caption         =   "Дополнительные опции:"
      Height          =   855
      Left            =   120
      TabIndex        =   23
      Top             =   4080
      Width           =   6615
      Begin VB.CheckBox Check6 
         Caption         =   "Точное написание теннисистов (вкл. пробелы и заглавные буквы)"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   5415
      End
      Begin VB.CheckBox Check3 
         Caption         =   "При загрузке переключить на русский язык"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   480
         Width           =   3735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Год поиска:"
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   6615
      Begin VB.CheckBox Check5 
         Caption         =   "Все года"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   210
         Width           =   980
      End
      Begin VB.OptionButton Option6 
         Caption         =   "2006"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5760
         TabIndex        =   21
         Top             =   210
         Width           =   735
      End
      Begin VB.OptionButton Option5 
         Caption         =   "2005"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5040
         TabIndex        =   20
         Top             =   210
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton Option4 
         Caption         =   "2004"
         Enabled         =   0   'False
         Height          =   255
         Left            =   4320
         TabIndex        =   19
         Top             =   210
         Width           =   735
      End
      Begin VB.OptionButton Option3 
         Caption         =   "2003"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3600
         TabIndex        =   18
         Top             =   210
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "2002"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2880
         TabIndex        =   17
         Top             =   210
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "2001"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2160
         TabIndex        =   16
         Top             =   210
         Width           =   735
      End
      Begin VB.OptionButton Option0 
         Caption         =   "2000"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1440
         TabIndex        =   15
         Top             =   210
         Width           =   735
      End
   End
   Begin XPButton.UserControl1 cExit 
      Height          =   615
      Left            =   4920
      TabIndex        =   13
      Top             =   3360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      Caption         =   "OK"
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
   Begin VB.Frame Frame2 
      Caption         =   "Изменение ссылки при нажатии на основной Лейбл ""TenScore"":"
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   6615
      Begin XPButton.UserControl1 cChange 
         Height          =   255
         Left            =   5160
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         Caption         =   "Изменить"
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
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Text            =   "http://www.marathonbet.com/odds.shtml"
         Top             =   240
         Width           =   6135
      End
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Хард"
      Enabled         =   0   'False
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   7
      Top             =   2640
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Турниры, где покарытие не указано"
      Enabled         =   0   'False
      Height          =   255
      Index           =   4
      Left            =   960
      TabIndex        =   6
      Top             =   3600
      Value           =   1  'Checked
      Width           =   3135
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Трава"
      Enabled         =   0   'False
      Height          =   255
      Index           =   3
      Left            =   960
      TabIndex        =   5
      Top             =   3360
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Ковер"
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   4
      Top             =   3120
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Грунт"
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   3
      Top             =   2880
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Опции поиска:"
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   4095
      Begin VB.CheckBox Check2 
         Caption         =   "С указанием турниров / выбрать турниры"
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   600
         Width           =   3500
      End
      Begin VB.CheckBox Check1 
         Caption         =   "С указанием месяцев"
         Height          =   255
         Left            =   480
         TabIndex        =   0
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Обратная связь"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5160
      TabIndex        =   11
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Посмотреть ""Help"" (on-line)"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   4545
      Picture         =   "Form2.frx":030A
      Top             =   2520
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   4440
      Picture         =   "Form2.frx":5AEC
      Top             =   2040
      Width           =   720
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cExit_Click()

If Check4(0).Value = 1 Then hard = 1 Else hard = 0
If Check4(1).Value = 1 Then soil = 1 Else soil = 0
If Check4(2).Value = 1 Then corpet = 1 Else corpet = 0
If Check4(3).Value = 1 Then grass = 1 Else grass = 0
If Check4(4).Value = 1 Then noncover = 1 Else noncover = 0

If Text1.Text <> "" Then
Linklabel = Text1.Text
Else: Linklabel = "http://www.marathonbet.com/odds.shtml"
End If
Module1.OptionsSave

SaveGOD
If Option0.Value = True Then GOD = 2000
If Option1.Value = True Then GOD = 2001
If Option2.Value = True Then GOD = 2002
If Option3.Value = True Then GOD = 2003
If Option4.Value = True Then GOD = 2004
If Option5.Value = True Then GOD = 2005
If Option6.Value = True Then GOD = 2006

LoadGOD
ForSearchGOD
LastData

If AllGod = 0 Then
If Len(Dir(CStr(Form1.Option3.Caption) + "\")) Then Form1.Option3.Enabled = True Else Form1.Option3.Enabled = False
If Len(Dir(CStr(Form1.Option4.Caption) + "\")) Then Form1.Option4.Enabled = True Else Form1.Option4.Enabled = False
If Len(Dir(CStr(Form1.Option5.Caption) + "\")) Then Form1.Option5.Enabled = True Else Form1.Option5.Enabled = False
vsecheki
Else
Form1.Option3.Enabled = False
Form1.Option4.Enabled = False
Form1.Option5.Enabled = False
For File = 1 To 12
Form1.Check1(File).Enabled = False
Next File
End If

Unload Form2
Form1.Text1.SetFocus
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then CheckMonth = 1 Else CheckMonth = 0
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then CheckTUR = 1 Else CheckTUR = 0
If Check2.Value = 1 Then
    Check4(0).Enabled = True
    Check4(1).Enabled = True
    Check4(2).Enabled = True
    Check4(3).Enabled = True
    Check4(4).Enabled = True
    Else
    Check4(0).Enabled = False
    Check4(1).Enabled = False
    Check4(2).Enabled = False
    Check4(3).Enabled = False
    Check4(4).Enabled = False
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then Raskladka = 1 Else Raskladka = 0
End Sub

Private Sub cChange_Click()
If Text1.Enabled = False Then Text1.Enabled = True Else Text1.Enabled = False
If cChange.Caption <> "OK" Then cChange.Caption = "OK" Else cChange.Caption = "Изменить"
End Sub
Private Sub Check4_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then _
  ChangeTo = 1 - Check4(Index).Value: _
  Check4(Index).Value = ChangeTo: _
  ReleaseCapture
End Sub

Private Sub Check4_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then Check4(Index).Value = ChangeTo
End Sub

Private Sub Check5_Click()
If Check5.Value = 0 Then
AllGod = 0
If Len(Dir("2000\")) Then Option0.Enabled = True Else Option0.Enabled = False
If Len(Dir("2001\")) Then Option1.Enabled = True Else Option1.Enabled = False
If Len(Dir("2002\")) Then Option2.Enabled = True Else Option2.Enabled = False
If Len(Dir("2003\")) Then Option3.Enabled = True Else Option3.Enabled = False
If Len(Dir("2004\")) Then Option4.Enabled = True Else Option4.Enabled = False
If Len(Dir("2005\")) Then Option5.Enabled = True Else Option5.Enabled = False
If Len(Dir("2006\")) Then Option6.Enabled = True Else Option6.Enabled = False
Else
AllGod = 1
Option0.Enabled = False
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
Option5.Enabled = False
Option6.Enabled = False
End If
End Sub


Private Sub Label1_Click()
ShellExecute Me.hwnd, vbNullString, "http://www.marathonbet.com/tab_2t.shtml?razd=-2", vbNullString, vbNullString, SW_SHOWNORMAL
Label1.ForeColor = &HFF0000
Label1.FontUnderline = False
End Sub

Private Sub Label2_Click()
ShellExecute hwnd, "open", "mailto:tenscore@narod.ru?subject=TenScore&body=%0A%0A%0A%0A%0A%0AЕсли вами обнаружены какие-то недостатки программы, она выдает ошибки, то для того чтобы было проще понять в чем причина и быстрее ее устранить, укажите на N ошибки (если N был показан) и описанием событий (ваших действий и ответных действий программы приведших к ошибке). Также укажите какую Операционную Систему вы используете (пример: Windows XP SP1) и краткие характеристики компьютера (процессор, ОЗУ, HDD), на который была установлена программа TenScore.", 0, 0, 5
Label2.ForeColor = &HFF0000
Label2.FontUnderline = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label2.ForeColor = &HFF0000
Label2.FontUnderline = False
Label1.ForeColor = &HFF0000
Label1.FontUnderline = False
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label2.ForeColor = &HFF&
Label2.FontUnderline = True
Label1.ForeColor = &HFF0000
Label1.FontUnderline = False
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label1.ForeColor = &HFF&
Label1.FontUnderline = True
Label2.ForeColor = &HFF0000
Label2.FontUnderline = False
End Sub
Private Sub Form_Load()
Text1.Text = Linklabel
Check1.Value = CheckMonth
Check2.Value = CheckTUR
    If CheckTUR = 1 Then
    Check4(0).Enabled = True
    Check4(1).Enabled = True
    Check4(2).Enabled = True
    Check4(3).Enabled = True
    Check4(4).Enabled = True
    Else
    Check4(0).Enabled = False
    Check4(1).Enabled = False
    Check4(2).Enabled = False
    Check4(3).Enabled = False
    Check4(4).Enabled = False
    End If
Check4(0).Value = hard
Check4(1).Value = soil
Check4(2).Value = corpet
Check4(3).Value = grass
Check4(4).Value = noncover
Check3.Value = Raskladka

If Len(Dir("2000\")) Then Option0.Enabled = True Else Option0.Enabled = False
If Len(Dir("2001\")) Then Option1.Enabled = True Else Option1.Enabled = False
If Len(Dir("2002\")) Then Option2.Enabled = True Else Option2.Enabled = False
If Len(Dir("2003\")) Then Option3.Enabled = True Else Option3.Enabled = False
If Len(Dir("2004\")) Then Option4.Enabled = True Else Option4.Enabled = False
If Len(Dir("2005\")) Then Option5.Enabled = True Else Option5.Enabled = False
If Len(Dir("2006\")) Then Option6.Enabled = True Else Option6.Enabled = False

If GOD = 2000 Then Option0.Value = True
If GOD = 2001 Then Option1.Value = True
If GOD = 2002 Then Option2.Value = True
If GOD = 2003 Then Option3.Value = True
If GOD = 2004 Then Option4.Value = True
If GOD = 2005 Then Option5.Value = True
If GOD = 2006 Then Option6.Value = True

If Correct = 1 Then Check6.Value = 1 Else Check6.Value = 0
If AllGod = 1 Then Check5.Value = 1 Else Check5.Value = 0
End Sub


















