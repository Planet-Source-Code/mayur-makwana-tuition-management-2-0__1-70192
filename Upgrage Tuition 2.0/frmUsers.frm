VERSION 5.00
Object = "{C30897B9-75AC-11D2-94E3-000000000000}#1.4#0"; "ARButton.ocx"
Begin VB.Form frmUsers 
   Appearance      =   0  'Flat
   BackColor       =   &H00636363&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User's Entry"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4020
   Icon            =   "frmUsers.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   4020
   StartUpPosition =   2  'CenterScreen
   Begin ARButtonCtrl.ARButton cmdadd 
      Height          =   360
      Left            =   105
      TabIndex        =   7
      Top             =   2700
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   635
      Caption         =   "&New"
      ForeColor       =   6513507
      ForeColorOnMouse=   16053492
      ForeColorOnFocus=   16053492
      BackColorOnMouse=   6513507
      BackColor       =   16053492
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
   Begin VB.TextBox txtfullname 
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
      Height          =   315
      Left            =   210
      TabIndex        =   1
      Top             =   1215
      Width           =   2775
   End
   Begin VB.TextBox txtpassword 
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
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   210
      MaxLength       =   10
      PasswordChar    =   "#"
      TabIndex        =   2
      Top             =   1905
      Width           =   1335
   End
   Begin VB.TextBox txtUser 
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
      Height          =   315
      Left            =   195
      MaxLength       =   10
      TabIndex        =   0
      Top             =   540
      Width           =   1455
   End
   Begin ARButtonCtrl.ARButton cmdedit 
      Height          =   360
      Left            =   840
      TabIndex        =   8
      Top             =   2700
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   635
      Caption         =   "&Edit"
      ForeColor       =   6513507
      ForeColorOnMouse=   16053492
      ForeColorOnFocus=   16053492
      BackColorOnMouse=   6513507
      BackColor       =   16053492
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
   Begin ARButtonCtrl.ARButton cmdsave 
      Height          =   360
      Left            =   1560
      TabIndex        =   9
      Top             =   2700
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   635
      Caption         =   "&Save"
      ForeColor       =   6513507
      ForeColorOnMouse=   16053492
      ForeColorOnFocus=   16053492
      BackColorOnMouse=   6513507
      BackColor       =   16053492
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
   Begin ARButtonCtrl.ARButton cmdcancel 
      Height          =   360
      Left            =   2295
      TabIndex        =   10
      Top             =   2700
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   635
      Caption         =   "&Cancel"
      ForeColor       =   6513507
      ForeColorOnMouse=   16053492
      ForeColorOnFocus=   16053492
      BackColorOnMouse=   6513507
      BackColor       =   16053492
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
   Begin ARButtonCtrl.ARButton cmdexit 
      Height          =   360
      Left            =   3030
      TabIndex        =   11
      Top             =   2700
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   635
      Caption         =   "&Exit"
      ForeColor       =   6513507
      ForeColorOnMouse=   16053492
      ForeColorOnFocus=   16053492
      BackColorOnMouse=   6513507
      BackColor       =   16053492
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
   Begin ARButtonCtrl.ARButton cmdfirst 
      Height          =   360
      Left            =   1020
      TabIndex        =   12
      Top             =   2340
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   635
      Caption         =   "<<"
      ForeColor       =   6513507
      ForeColorOnMouse=   16053492
      ForeColorOnFocus=   16053492
      BackColorOnMouse=   6513507
      BackColor       =   16053492
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
   Begin ARButtonCtrl.ARButton cmdprevious 
      Height          =   360
      Left            =   1455
      TabIndex        =   13
      Top             =   2340
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   635
      Caption         =   "<"
      ForeColor       =   6513507
      ForeColorOnMouse=   16053492
      ForeColorOnFocus=   16053492
      BackColorOnMouse=   6513507
      BackColor       =   16053492
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
   Begin ARButtonCtrl.ARButton cmdnext 
      Height          =   360
      Left            =   1905
      TabIndex        =   14
      Top             =   2340
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   635
      Caption         =   ">"
      ForeColor       =   6513507
      ForeColorOnMouse=   16053492
      ForeColorOnFocus=   16053492
      BackColorOnMouse=   6513507
      BackColor       =   16053492
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
   Begin ARButtonCtrl.ARButton cmdlast 
      Height          =   360
      Left            =   2370
      TabIndex        =   6
      Top             =   2340
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   635
      Caption         =   ">>"
      ForeColor       =   6513507
      ForeColorOnMouse=   16053492
      ForeColorOnFocus=   16053492
      BackColorOnMouse=   6513507
      BackColor       =   16053492
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
   Begin VB.Label Label5 
      BackColor       =   &H00F0F0F0&
      Height          =   285
      Left            =   0
      TabIndex        =   16
      Top             =   3165
      Width           =   4110
   End
   Begin VB.Label Label4 
      BackColor       =   &H00F0F0F0&
      Height          =   270
      Left            =   -75
      TabIndex        =   15
      Top             =   -30
      Width           =   4110
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "User Full Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F4F4F4&
      Height          =   330
      Left            =   195
      TabIndex        =   5
      Top             =   990
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F4F4F4&
      Height          =   330
      Left            =   195
      TabIndex        =   4
      Top             =   1665
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F4F4F4&
      Height          =   330
      Left            =   195
      TabIndex        =   3
      Top             =   315
      Width           =   1575
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub FormLoad_State()
    cmdadd.Enabled = True
    cmdedit.Enabled = True
    cmdsave.Enabled = False
    cmdcancel.Enabled = True
    cmdfirst.Enabled = True
    cmdnext.Enabled = True
    cmdprevious.Enabled = True
    cmdlast.Enabled = True
End Sub

Public Sub addedit_state()
    cmdadd.Enabled = False
    cmdedit.Enabled = False
    cmdsave.Enabled = True
    cmdcancel.Enabled = True
    cmdfirst.Enabled = False
    cmdnext.Enabled = False
    cmdprevious.Enabled = False
    cmdlast.Enabled = False
End Sub

Public Sub enable_fields()
Dim ctrl As Variant
On Error Resume Next
    For Each ctrl In Me
        If TypeOf ctrl Is TextBox Then
            ctrl.Enabled = True
        ElseIf TypeOf ctrl Is ListBox Then
            ctrl.Enabled = True
        ElseIf TypeOf ctrl Is ComboBox Then
            ctrl.Enabled = True
        ElseIf TypeOf ctrl Is CheckBox Then
            ctrl.Enabled = True
        ElseIf TypeOf ctrl Is DTPicker Then
            ctrl.Enabled = True
       ' ElseIf TypeOf ctrl Is DataCombo Then
        '    ctrl.Text = vbNullString
        'ElseIf TypeOf ctrl Is MSFlexGrid Then
        '    ctrl.Clear
        End If
    Next
End Sub

Public Sub Disable_Fields()
Dim ctrl As Variant
On Error Resume Next
    For Each ctrl In Me
        If TypeOf ctrl Is TextBox Then
            ctrl.Enabled = False
        ElseIf TypeOf ctrl Is ListBox Then
            ctrl.Enabled = False
        ElseIf TypeOf ctrl Is ComboBox Then
            ctrl.Enabled = False
        ElseIf TypeOf ctrl Is CheckBox Then
            ctrl.Enabled = False
        ElseIf TypeOf ctrl Is DTPicker Then
            ctrl.Enabled = False
       ' ElseIf TypeOf ctrl Is DataCombo Then
        '    ctrl.Text = vbNullString
        'ElseIf TypeOf ctrl Is MSFlexGrid Then
        '    ctrl.Clear
        End If
    Next
End Sub

Public Sub set_fields()
On Error Resume Next
    txtUser.Text = rs.Fields(0)
    txtfullname.Text = rs.Fields(1)
    txtpassword.Text = rs.Fields(2)
   End Sub

Private Sub cmbtype_GotFocus()
Changecolor True, Me.ActiveControl
End Sub

Private Sub cmbtype_LostFocus()
Changecolor False, txtbox
End Sub

Private Sub cmdadd_Click()
cmdcancel_Click
addedit_state
enable_fields
Status = False
End Sub

Private Sub cmdcancel_Click()
Unload Me
Me.Show
End Sub

Private Sub cmdedit_Click()
addedit_state
Status = True
enable_fields
txtUser.Enabled = False
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdfirst_Click()
rs.MoveFirst
set_fields
cmdedit.Enabled = True
End Sub

Private Sub cmdlast_Click()
rs.MoveLast
set_fields
cmdedit.Enabled = True
End Sub

Private Sub cmdnext_Click()
On Error Resume Next
If rs.EOF = False Then
    rs.MoveNext
    set_fields
Else
    MsgBox "This is Last Record", vbCritical, "Check It"
End If
cmdedit.Enabled = True
End Sub

Private Sub cmdprevious_Click()
On Error Resume Next
If rs.BOF = False Then
    rs.MovePrevious
    set_fields
Else
    MsgBox "This is First Record", vbCritical, "Check It"
End If
cmdedit.Enabled = True
End Sub

Private Sub cmdsave_Click()
Dim RsCheck As New ADODb.Recordset
If txtUser.Text = "" Then
    MsgBox "You Must Enter User ID First", vbCritical, "Check It"
    txtUser.SetFocus
    Exit Sub
End If
If txtpassword.Text = "" Then
    MsgBox "You Must Enter Password", vbCritical, "Check It"
    txtpassword.SetFocus
    Exit Sub
End If
Set rs = Nothing
If Status = False Then
    Set RsCheck = Nothing
    RsCheck.Open "select * from users where userid = '" & txtUser.Text & "'", cn, adOpenKeyset, adLockReadOnly
    If RsCheck.RecordCount > 0 Then
        MsgBox "This Record is Already Entered", vbCritical, "Check It"
        txtUser.SetFocus
        Exit Sub
    End If
    rs.Open "select * from users", cn, adOpenKeyset, adLockOptimistic
    rs.AddNew
    rs!userid = txtUser.Text
    rs!fullname = txtfullname.Text
    rs!Password = txtpassword.Text
   
Else
    rs.Open "select * from users where userid = '" & txtUser.Text & "'", cn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then
        rs!userid = txtUser.Text
        rs!fullname = txtfullname.Text
        rs!Password = txtpassword.Text
       
    End If
End If
rs.Update
MsgBox "Your Record Has Been Saved Successfuly", vbInformation, "Congrtulation"
Unload Me
Me.Show
End Sub

Private Sub Form_Load()
Con
Set rs = Nothing
rs.Open "select * from users", cn, adOpenKeyset, adLockOptimistic
txtfullname.Enabled = False
txtUser.Enabled = False
txtpassword.Enabled = False
cmdedit.Enabled = False
cmdsave.Enabled = False
End Sub

Private Sub txtfullname_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtfullname_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtPassword_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtPassword_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtUser_GotFocus()
Changecolor True, Me.ActiveControl
End Sub

Private Sub txtUser_LostFocus()
Changecolor False, txtbox
End Sub
