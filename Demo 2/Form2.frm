VERSION 5.00
Object = "{A51095D7-8D17-11D6-9913-E1D1DF4BFD40}#1.0#0"; "XPButton.ocx"
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " TenScore - Опции Демо версии"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8340
   ControlBox      =   0   'False
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   6720
      Width           =   6135
   End
   Begin XPButton.UserControl1 cExit 
      Height          =   615
      Left            =   3480
      TabIndex        =   2
      Top             =   8040
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
   Begin VB.Line Line6 
      BorderColor     =   &H8000000C&
      X1              =   480
      X2              =   5640
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      X1              =   480
      X2              =   5640
      Y1              =   7455
      Y2              =   7455
   End
   Begin VB.Label Label11 
      Caption         =   "Если у Вас остались вопросы, их можно задать по адресу:"
      Height          =   255
      Left            =   1080
      TabIndex        =   16
      Top             =   7680
      Width           =   4575
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000C&
      X1              =   480
      X2              =   5640
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      X1              =   480
      X2              =   5640
      Y1              =   6015
      Y2              =   6015
   End
   Begin VB.Image Image4 
      Height          =   720
      Left            =   3240
      Picture         =   "Form2.frx":030A
      Top             =   5160
      Width           =   720
   End
   Begin VB.Label Label15 
      Caption         =   "Полная версия - платная (200 рублей), и изготавливается под  Ваш Заказ."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   4680
      Width           =   6855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   480
      X2              =   5640
      Y1              =   3855
      Y2              =   3855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      X1              =   480
      X2              =   5640
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label Label13 
      Caption         =   "2 - предусмотрена возможность Поиска с учетом Покрытия (хард, грунт, ковер, трава)"
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   2040
      Width           =   6735
   End
   Begin VB.Image Image3 
      Height          =   690
      Left            =   240
      Picture         =   "Form2.frx":5AEC
      Top             =   240
      Width           =   2670
   End
   Begin VB.Label Label14 
      Caption         =   "№ вашей Демо Версии:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   6480
      Width           =   2175
   End
   Begin VB.Label Label12 
      Caption         =   "Посмотреть условия Заказа Полной версии TenScore"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3960
      TabIndex        =   11
      Top             =   5400
      Width           =   4215
   End
   Begin VB.Label Label10 
      Caption         =   "Для Заказа будет необходим № вашей Демо Версии (он указан ниже)."
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   4920
      Width           =   5655
   End
   Begin VB.Label Label9 
      Caption         =   "Демо версия является бесплатной и свободно распространяемой."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   4200
      Width           =   6135
   End
   Begin VB.Label Label8 
      Caption         =   "1 - активна кнопка ""Личные встречи"""
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Label Label7 
      Caption         =   "5 - можно изменять настройки программы, которые будут сохранены при выходе"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   2760
      Width           =   7095
   End
   Begin VB.Label Label6 
      Caption         =   "4 - возможность скачивать данные периодами (с указанием начальной и конечной даты)"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   2520
      Width           =   6855
   End
   Begin VB.Label Label5 
      Caption         =   "3 - данные теннисных матчей могут быть скачены за любой день"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   2280
      Width           =   5055
   End
   Begin VB.Label Label4 
      Caption         =   "В полной версии TenScore:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Вы используете демонстрационную версию программы ""TenScore"", в ней намеренно ограниченны функции полной версии."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   3120
      TabIndex        =   3
      Top             =   240
      Width           =   5175
   End
   Begin VB.Label Label2 
      Caption         =   "tenscore@k66.ru"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6000
      TabIndex        =   1
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Более подробное описание  Полной версии TenScore"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3840
      TabIndex        =   0
      Top             =   3240
      Width           =   4215
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   5400
      Picture         =   "Form2.frx":6870
      Top             =   7440
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   3240
      Picture         =   "Form2.frx":C052
      Top             =   3000
      Width           =   720
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lplplplpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1 ' обе для ссылок
Dim l As String
Private Declare Function GetVolumeSerialNumber Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Public Function VolumeSerialNumber(ByVal RootPath As String) As String

Dim VolLabel As String
Dim VolSize As Long
Dim Serial As Long
Dim MaxLen As Long
Dim Flags As Long
Dim Name As String
Dim NameSize As Long
Dim s As String
Dim ret As Boolean

ret = GetVolumeSerialNumber(RootPath, VolLabel, VolSize, Serial, MaxLen, Flags, Name, NameSize)

If ret Then
    s = Format(Hex(Serial), "00000000")
    l = Format(Serial)
    VolumeSerialNumber = s 'Left(s, 4) + "-" + Right(s, 4)
Else
    VolumeSerialNumber = "00000000"
End If
End Function

Private Sub Form_Load()
On Error GoTo ZapasVar
' Объявляем переменные
Dim strComputer        ' Имя компьютера
Dim strNamespace       ' Имя пространства имен
Dim strClass           ' Имя класса
Dim objClass           ' Объект SWbemObject (класс WMI)
Dim colOperatingSystems ' Коллекция экземпляров класса WMI
Dim objOperatingSystem ' Элемент коллекции
Dim strResult          ' Строка для вывода на экран
Dim MyString As String
Dim MyString2 As String
Dim FinalSting As String

strComputer = "."
strNamespace = "Root\CIMV2"
strClass = "Win32_PhysicalMedia"
Set objClass = GetObject("WinMgmts:\\" & strComputer & "\" & strNamespace & ":" & strClass)
Set colOperatingSystems = objClass.Instances_
For Each objOperatingSystem In colOperatingSystems
  strResult = strResult & objOperatingSystem.SerialNumber & vbCrLf
Next

Text1.Text = Mid(strResult, 1, 40) ' выведет только 1-й HDD

ZapasVar:

MyString = VolumeSerialNumber(Left(App.Path, 3))
MyString2 = l
FinalSting = Mid(MyString2, 5, 1) & Mid(MyString2, 4, 1) & _
    Mid(MyString, 5, 1) & Mid(MyString, 6, 1) & _
    Mid(MyString2, 3, 1) & Mid(MyString2, 1, 1) & _
    Mid(MyString2, 2, 1) & Mid(MyString2, 7, 1) & _
    Mid(MyString2, 1, 1) & Mid(MyString2, 9, 1) & _
    Mid(MyString, 3, 1) & Mid(MyString, 4, 1) & _
    Mid(MyString2, 6, 1) & Mid(MyString2, 2, 1) & _
    Mid(MyString, 1, 1) & Mid(MyString, 2, 1) & _
    Mid(MyString2, 7, 1) & Mid(MyString2, 2, 1) & _
    Mid(MyString, 3, 1) & Mid(MyString, 4, 1) & _
    Mid(MyString2, 1, 1) & Mid(MyString2, 5, 1) & _
    Mid(MyString2, 2, 1) & Mid(MyString2, 9, 1) & _
    Mid(MyString2, 4, 1) & Mid(MyString2, 3, 1) & _
    Mid(MyString, 3, 1) & Mid(MyString, 1, 1) & _
    Mid(MyString2, 5, 1) & Mid(MyString2, 7, 1) & _
    Mid(MyString, 5, 1) & Mid(MyString, 3, 1) & _
    Mid(MyString2, 2, 1) & Mid(MyString2, 3, 1) & _
    Mid(MyString, 6, 1) & Mid(MyString, 1, 1) & _
    Mid(MyString2, 5, 1) & Mid(MyString2, 7, 1) & _
    Mid(MyString, 3, 1) & Mid(MyString, 4, 1) '& _

Text1.Text = FinalSting
End Sub

Private Sub Command3_Click()
End
End Sub

'*****************************************************************

Private Sub cExit_Click()
Unload Form2
Form1.Text1.SetFocus
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

