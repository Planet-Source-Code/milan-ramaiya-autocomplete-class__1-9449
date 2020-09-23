VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "AutoComplete DEMO"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "FMain.frx":0000
      Top             =   1080
      Width           =   4335
   End
   Begin VB.CheckBox chkAutoAdd 
      Caption         =   "Enable AutoAdd"
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.CheckBox chkAutoComplete 
      Caption         =   "Enable AutoComplete"
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.ComboBox cboTest 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Text            =   "bleh"
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim acTest As New CAutoComplete

Private Sub chkAutoAdd_Click()
acTest.AutoAdd = chkAutoAdd.Value
End Sub

Private Sub chkAutoComplete_Click()
If chkAutoComplete.Value = 1 Then
    Set acTest.LinkedComboBox = cboTest
Else
    Set acTest.LinkedComboBox = Nothing
End If
End Sub


