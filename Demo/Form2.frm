VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " TenScore - Демо Версия"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6750
   ControlBox      =   0   'False
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cExit 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   615
      Left            =   2520
      TabIndex        =   0
      Top             =   5520
      Width           =   1455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cExit_Click()
Form1.Enabled = True
Unload Form2
Form1.Text1.SetFocus
End Sub
