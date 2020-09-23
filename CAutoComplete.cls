VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAutoComplete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mAutoAdd As Boolean
Private WithEvents mComboBox As ComboBox
Attribute mComboBox.VB_VarHelpID = -1

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const CB_ERR = (-1)
Private Const CB_FINDSTRING = &H14C
Private Const CB_FINDSTRINGEXACT = &H158
Private Const CB_SHOWDROPDOWN = &H14F
Public Property Let AutoAdd(var As Boolean)
mAutoAdd = var
End Property

Public Property Get AutoAdd() As Boolean
AutoAdd = mAutoAdd
End Property
Private Function FindItem(sFind As String, Optional IndexStart As Long = -1, Optional bExact As Boolean = False) As Long
Dim wMsg As Long
wMsg = IIf(bExact, CB_FINDSTRINGEXACT, CB_FINDSTRING)
sFind = sFind & Chr(0)
FindItem = SendMessage(mComboBox.hwnd, wMsg, IndexStart, ByVal sFind)
End Function

Public Property Set LinkedComboBox(cbo As ComboBox)
Set mComboBox = cbo
End Property

Public Property Get LinkedComboBox() As ComboBox
Set LinkedComboBox = mComboBox
End Property



Private Sub ShowDropDown(bShow As Boolean)
SendMessage mComboBox.hwnd, CB_SHOWDROPDOWN, bShow, 0
End Sub



Private Sub mComboBox_KeyPress(KeyAscii As Integer)
If mAutoAdd Then
    With mComboBox
        Dim iLoc As Long
        If KeyAscii = 13 Then
            iLoc = FindItem(.Text, , True)
            If iLoc = -1 Then .AddItem .Text
        End If
    End With
End If
End Sub

Private Sub mComboBox_KeyUp(KeyCode As Integer, Shift As Integer)
With mComboBox
    Dim iLoc As Integer, sText As String
    If KeyCode <> vbKeyBack And KeyCode > 48 Then
        sText = .Text
        iLoc = FindItem(sText, , False)
        If iLoc <> -1 Then
            ShowDropDown True
            .Text = .List(iLoc)
            .SelStart = Len(sText) - 1
            .SelLength = Len(.Text) - .SelStart
        Else
            Dim iSelStart As Integer, iSelLength As Integer
            iSelStart = .SelStart
            iSelLength = .SelLength
            ShowDropDown False
            .Text = sText
            .SelStart = iSelStart
            .SelLength = iSelLength
        End If
    End If
End With
End Sub
