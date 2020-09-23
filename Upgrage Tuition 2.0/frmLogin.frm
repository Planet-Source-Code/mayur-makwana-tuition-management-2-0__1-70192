VERSION 5.00
Object = "{C30897B9-75AC-11D2-94E3-000000000000}#1.4#0"; "arbutton.ocx"
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H00636363&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1440
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   2805
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   850.799
   ScaleMode       =   0  'User
   ScaleWidth      =   2633.743
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ARButtonCtrl.ARButton cmdOk 
      Height          =   375
      Left            =   795
      TabIndex        =   4
      Top             =   945
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "&OK"
      ForeColor       =   6513507
      ForeColorOnMouse=   15461355
      BackColorOnMouse=   6513507
      BackColor       =   15461355
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1215
      TabIndex        =   0
      Top             =   105
      Width           =   1440
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1215
      PasswordChar    =   "!"
      TabIndex        =   1
      Top             =   525
      Width           =   1440
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00656261&
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0F0F0&
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   2
      Top             =   105
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00656261&
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0F0F0&
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   3
      Top             =   555
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdcancel_Click()
   End
End Sub

Private Sub cmdOK_Click()
Dim i As Integer
Dim RsCheck As New ADODb.Recordset
Set RsCheck = Nothing
RsCheck.Open "Select * from users", cn, adOpenKeyset, adLockReadOnly
If RsCheck.RecordCount > 0 Then
    RsCheck.MoveFirst
    For i = 1 To RsCheck.RecordCount
        If txtUserName.Text = RsCheck.Fields(0) And txtpassword.Text = RsCheck.Fields(2) Then
            frmMain.Show
'            If rsCheck.Fields(3) = "M" Then
'                frmMain.Toolbar1.Buttons(5).Enabled = True
'            Else
'                frmMain.Toolbar1.Buttons(5).Enabled = False
'            End If
            Unload Me
            Exit Sub
        Else
            RsCheck.MoveNext
        End If
    Next
    MsgBox "Not a Valid User Try Again", vbCritical, "Sorry..."
    Unload Me
    Me.Show
Else
    frmUsers.Show
    Unload Me
End If
End Sub

Private Sub Form_Load()
Con

End Sub

Private Sub txtPassword_GotFocus()
Changecolor True, Me.ActiveControl
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub txtPassword_LostFocus()
Changecolor False, txtbox
End Sub

Private Sub txtUserName_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab

End Sub

Private Sub txtUserName_LostFocus()
Changecolor False, txtbox

End Sub
