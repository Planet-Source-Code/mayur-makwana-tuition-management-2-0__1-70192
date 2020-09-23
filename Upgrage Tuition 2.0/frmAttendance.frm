VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{C30897B9-75AC-11D2-94E3-000000000000}#1.4#0"; "ARButton.ocx"
Begin VB.Form frmAttendance 
   Appearance      =   0  'Flat
   BackColor       =   &H00636363&
   Caption         =   "Attendance"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8385
   Icon            =   "frmAttendance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   8385
   StartUpPosition =   1  'CenterOwner
   Begin MSComCtl2.DTPicker DTPattendance 
      Height          =   300
      Left            =   165
      TabIndex        =   18
      Top             =   225
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   529
      _Version        =   393216
      Format          =   54657025
      CurrentDate     =   39019
   End
   Begin VB.ComboBox cmbstatus 
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
      ItemData        =   "frmAttendance.frx":08CA
      Left            =   7140
      List            =   "frmAttendance.frx":08D4
      TabIndex        =   4
      Top             =   810
      Width           =   1005
   End
   Begin VB.TextBox txtname 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4185
      TabIndex        =   17
      Top             =   810
      Width           =   2955
   End
   Begin VB.ComboBox cmbroll 
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
      ItemData        =   "frmAttendance.frx":08DE
      Left            =   3180
      List            =   "frmAttendance.frx":08E0
      TabIndex        =   3
      Top             =   810
      Width           =   1005
   End
   Begin VB.ComboBox cmbbatch 
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
      ItemData        =   "frmAttendance.frx":08E2
      Left            =   2175
      List            =   "frmAttendance.frx":08E4
      TabIndex        =   2
      Top             =   810
      Width           =   1005
   End
   Begin VB.ComboBox cmbstd 
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
      ItemData        =   "frmAttendance.frx":08E6
      Left            =   1170
      List            =   "frmAttendance.frx":08E8
      TabIndex        =   1
      Top             =   810
      Width           =   1005
   End
   Begin MSComctlLib.ListView lstv2 
      Height          =   2445
      Left            =   135
      TabIndex        =   9
      Top             =   1140
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   4313
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Branch"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Standard"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Batch"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Roll No"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Name"
         Object.Width           =   5362
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Status"
         Object.Width           =   1827
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   840
      Top             =   3105
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.ComboBox cmbbranch 
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
      ItemData        =   "frmAttendance.frx":08EA
      Left            =   165
      List            =   "frmAttendance.frx":08EC
      TabIndex        =   0
      Top             =   810
      Width           =   1005
   End
   Begin ARButtonCtrl.ARButton cmdcancel 
      Height          =   360
      Left            =   3735
      TabIndex        =   6
      Top             =   3690
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   635
      Caption         =   "&Cancel"
      ForeColor       =   -2147483630
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
      Left            =   4830
      TabIndex        =   7
      Top             =   3690
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   635
      Caption         =   "&Exit"
      ForeColor       =   -2147483630
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
   Begin ARButtonCtrl.ARButton Command1 
      Height          =   300
      Left            =   1755
      TabIndex        =   10
      Top             =   225
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   529
      Caption         =   "&Search"
      ForeColor       =   -2147483630
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
      Left            =   2640
      TabIndex        =   5
      Top             =   3690
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   635
      Caption         =   "&Save"
      ForeColor       =   -2147483630
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
   Begin VB.Label Label7 
      BackColor       =   &H00636363&
      Caption         =   "Attendance Date:"
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
      Height          =   240
      Left            =   180
      TabIndex        =   16
      Top             =   15
      Width           =   1590
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
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
      Height          =   255
      Left            =   7155
      TabIndex        =   15
      Top             =   615
      Width           =   825
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
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
      Height          =   255
      Left            =   4170
      TabIndex        =   14
      Top             =   615
      Width           =   825
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Roll No:"
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
      Height          =   255
      Left            =   3180
      TabIndex        =   13
      Top             =   615
      Width           =   825
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Batch:"
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
      Height          =   255
      Left            =   2175
      TabIndex        =   12
      Top             =   615
      Width           =   825
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Branch:"
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
      Height          =   255
      Left            =   180
      TabIndex        =   11
      Top             =   615
      Width           =   825
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Standard:"
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
      Height          =   255
      Left            =   1170
      TabIndex        =   8
      Top             =   615
      Width           =   825
   End
End
Attribute VB_Name = "frmAttendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmpRs As New ADODb.Recordset
Dim J As Integer

Private Sub cmbbatch_GotFocus()
Dim i As Integer
Changecolor True, Me.ActiveControl
cmbbatch.Clear
Set tmpRs = Nothing
tmpRs.Open "select distinct batch from student where branch = '" & cmbbranch.Text & "' and std = '" & cmbstd.Text & "'", cn, adOpenKeyset, adLockReadOnly
If tmpRs.RecordCount > 0 Then
    For i = 1 To tmpRs.RecordCount
        cmbbatch.AddItem tmpRs.Fields(0)
        tmpRs.MoveNext
    Next
End If

End Sub

Private Sub cmbbranch_gotfocus()
Dim i As Integer
Changecolor True, Me.ActiveControl
cmbbranch.Clear
Set tmpRs = Nothing
tmpRs.Open "select distinct branch from student", cn, adOpenKeyset, adLockReadOnly
If tmpRs.RecordCount > 0 Then
    For i = 1 To tmpRs.RecordCount
        cmbbranch.AddItem IIf(IsNull(tmpRs.Fields(0)), " ", tmpRs.Fields(0))
        tmpRs.MoveNext
    Next
End If
End Sub

Private Sub cmbroll_Click()
'Dim i As Integer
Changecolor True, Me.ActiveControl
'cmbroll.Clear
Set tmpRs = Nothing
tmpRs.Open "select stdname from student where batch = '" & cmbbatch.Text & "' and branch = '" & cmbbranch.Text & "' and std = '" & cmbstd.Text & "' and rollno = '" & cmbroll.Text & "'", cn, adOpenKeyset, adLockReadOnly
If tmpRs.RecordCount > 0 Then
    txtname.Text = tmpRs.Fields(0)
End If
End Sub

Private Sub cmbroll_GotFocus()
Dim i As Integer
Changecolor True, Me.ActiveControl
cmbroll.Clear
Set tmpRs = Nothing
tmpRs.Open "select distinct rollno from student where batch = '" & cmbbatch.Text & "' and branch = '" & cmbbranch.Text & "' and std = '" & cmbstd.Text & "'", cn, adOpenKeyset, adLockReadOnly
If tmpRs.RecordCount > 0 Then
    For i = 1 To tmpRs.RecordCount
        cmbroll.AddItem tmpRs.Fields(0)
        tmpRs.MoveNext
    Next
End If
End Sub

Private Sub cmbstatus_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    J = lstv2.ListItems.Count + 1
    lstv2.ListItems.Add = cmbbranch.Text
    lstv2.ListItems(J).SubItems(1) = cmbstd.Text
    lstv2.ListItems(J).SubItems(2) = cmbbatch.Text
    lstv2.ListItems(J).SubItems(3) = cmbroll.Text
    lstv2.ListItems(J).SubItems(4) = txtname.Text
    lstv2.ListItems(J).SubItems(5) = cmbstatus.Text
    'cmbbranch.ListIndex = 0
    'cmbstd.ListIndex = 0
    'cmbbatch.ListIndex = 0
    cmbroll.ListIndex = 0
    txtname.Text = ""
    cmbstatus.ListIndex = 0
    cmbroll.SetFocus
End If
End Sub

Private Sub cmbstd_GotFocus()
Dim i As Integer
Changecolor True, Me.ActiveControl
cmbstd.Clear
Set tmpRs = Nothing
tmpRs.Open "select distinct std from student where branch = '" & cmbbranch.Text & "'", cn, adOpenKeyset, adLockReadOnly
If tmpRs.RecordCount > 0 Then
    For i = 1 To tmpRs.RecordCount
        cmbstd.AddItem tmpRs.Fields(0)
        tmpRs.MoveNext
    Next
End If
End Sub

Private Sub cmdcancel_Click()
Unload Me
Me.Show
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdsave_Click()
Dim rssave1 As New ADODb.Recordset
Set tmpRs = Nothing
tmpRs.Open "delete from Attendance where adate = #" & DTPattendance.Value & "#", cn, adOpenKeyset, adLockOptimistic
For J = 1 To lstv2.ListItems.Count
    Set rssave1 = Nothing
    rssave1.Open "select * from Attendance", cn, adOpenDynamic, adLockOptimistic
    rssave1.AddNew
    rssave1.Fields(0) = DTPattendance.Value
    rssave1.Fields(1) = Trim(lstv2.ListItems(J).Text)
    rssave1.Fields(2) = Trim(lstv2.ListItems(J).SubItems(1))
    rssave1.Fields(3) = Trim(lstv2.ListItems(J).SubItems(2))
    rssave1.Fields(4) = Trim(lstv2.ListItems(J).SubItems(3))
    rssave1.Fields(5) = Trim(lstv2.ListItems(J).SubItems(4))
    rssave1.Fields(6) = Trim(lstv2.ListItems(J).SubItems(5))
    rssave1.Update
Next
        MsgBox "Your Record Have Been Save Sucessfully", vbInformation, "Congractulation"
End Sub

Private Sub Command1_Click()
On Error Resume Next
Static i As Integer
Dim rsgrid As New ADODb.Recordset
lstv2.ListItems.Clear
    Set rsgrid = Nothing
    rsgrid.Open "select * from Attendance where adate=#" & DTPattendance.Value & "#", cn, adOpenDynamic, adLockOptimistic
    If rsgrid.RecordCount > 0 Then
        For J = 1 To rsgrid.RecordCount
            lstv2.ListItems.Add = rsgrid.Fields(1)
            lstv2.ListItems(J).SubItems(1) = rsgrid.Fields(2)
            lstv2.ListItems(J).SubItems(2) = rsgrid.Fields(3)
            lstv2.ListItems(J).SubItems(3) = rsgrid.Fields(4)
            lstv2.ListItems(J).SubItems(4) = rsgrid.Fields(5)
            lstv2.ListItems(J).SubItems(5) = rsgrid.Fields(6)
            rsgrid.MoveNext
        Next J
    End If
    Exit Sub

End Sub

Private Sub Command_Click()

End Sub

Private Sub Form_Load()
Con
Set rs = Nothing
rs.Open "select * from Attendance", cn, adOpenKeyset, adLockOptimistic

Set tmpRs = Nothing
tmpRs.Open "Select roll from student", cn, adOpenKeyset, adLockReadOnly
If tmpRs.RecordCount > 0 Then
    cmbroll.Clear
    Do While tmpRs.EOF = False
        cmbroll.AddItem tmpRs.Fields(0)
        tmpRs.MoveNext
    Loop
End If
txtname.Enabled = False
End Sub

Private Sub lstv2_DblClick()
On Error GoTo errhand:
    cmbbranch.Text = lstv2.SelectedItem.Text
    cmbstd.Text = lstv2.SelectedItem.SubItems(1)
    cmbbatch.Text = lstv2.SelectedItem.SubItems(2)
    cmbroll.Text = lstv2.SelectedItem.SubItems(3)
    txtname.Text = lstv2.SelectedItem.SubItems(4)
    cmbstatus.Text = lstv2.SelectedItem.SubItems(5)
    lstv2.ListItems.Remove (lstv2.SelectedItem.Index)
    Exit Sub
errhand:
    MsgBox "Some Error !! " & Err.Description, vbInformation
    Err.Clear
End Sub

