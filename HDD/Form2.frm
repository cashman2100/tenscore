VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " TenScore - About"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7545
   ControlBox      =   0   'False
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cExit 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   615
      Left            =   2880
      TabIndex        =   0
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label Label12 
      Caption         =   "http://www.tenscore.narod.ru/"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3840
      TabIndex        =   12
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Label11 
      Caption         =   "* услови€ заказа смотрите на WEB сайте"
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   3480
      Width           =   3255
   End
   Begin VB.Label Label10 
      Caption         =   $"Form2.frx":030A
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   6975
   End
   Begin VB.Label Label9 
      Caption         =   " ƒемо верви€ €вл€етс€ бесплатной и свободно распростран€емой."
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   2280
      Width           =   5295
   End
   Begin VB.Label Label8 
      Caption         =   "- поиск не только ќдиночных игр, но и ѕарных/ћикстов"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label Label7 
      Caption         =   "- кнопка ""Ћичные встречи"" активна"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label Label6 
      Caption         =   "- возможность скачивать данные периодами (с указанием дат), но не более одного года"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   6855
   End
   Begin VB.Label Label5 
      Caption         =   "- данные теннисных матчей могут быть скачены за любой день, в пределах GOD года"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   6855
   End
   Begin VB.Label Label4 
      Caption         =   "¬ полной версии TenScore:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "¬ы используете демонстрационную версию программы ""TenScore"", в ней намеренно ограниченны функции полной версии."
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6495
   End
   Begin VB.Label Label2 
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
      Left            =   4080
      TabIndex        =   2
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "—ерийный номер вашего жесткого диска (HDD):"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   3735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetVolumeSerialNumber Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lplplplpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1
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
    'Create an 8 character string
    s = Format(Hex(Serial), "00000000")
    'Adds the '-' between the first 4 characters and the last 4 characters
    VolumeSerialNumber = Left(s, 4) + "-" + Right(s, 4)
Else
    'If the call to API function fails the function returns a zero serial number
    VolumeSerialNumber = "0000-0000"
End If
End Function
Private Sub cExit_Click()
End
End Sub

Private Sub Form_Load()
Label2.Caption = VolumeSerialNumber("C:\") 'Shows the serial number of your Hard Disk
End Sub

Private Sub Label12_Click()
ShellExecute Me.hWnd, vbNullString, "http://www.tenscore.narod.ru/", vbNullString, vbNullString, SW_SHOWNORMAL
Label12.ForeColor = &HFF0000
Label12.FontUnderline = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label12.ForeColor = &HFF0000
Label12.FontUnderline = False
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label12.ForeColor = &HFF&
Label12.FontUnderline = True
End Sub



