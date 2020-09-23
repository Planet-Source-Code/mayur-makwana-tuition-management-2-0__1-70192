VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{2D5F9802-7034-4F68-BAC9-E056D364E154}#1.0#0"; "CompControls.ocx"
Object = "{C30897B9-75AC-11D2-94E3-000000000000}#1.4#0"; "ARButton.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0FF&
   Caption         =   "Tuition"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9450
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmmain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   9450
   WindowState     =   2  'Maximized
   Begin ARButtonCtrl.ARButton ARButton10 
      Height          =   285
      Left            =   1245
      TabIndex        =   3
      Top             =   6975
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   503
      Caption         =   "&Find"
      ForeColor       =   16777215
      ForeColorOnMouse=   -2147483630
      BackColorOnMouse=   -2147483643
      BackColor       =   6513507
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CompControler.CompControl CompControl1 
      Left            =   840
      Top             =   2880
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   1185
      TabIndex        =   2
      Top             =   6720
      Width           =   9630
      Begin VB.PictureBox ARButton4 
         Height          =   285
         Left            =   135
         ScaleHeight     =   225
         ScaleWidth      =   630
         TabIndex        =   4
         Top             =   255
         Width           =   690
      End
      Begin ARButtonCtrl.ARButton ARButton7 
         Height          =   285
         Left            =   900
         TabIndex        =   5
         Top             =   255
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         Caption         =   "&Open CD ROM"
         ForeColor       =   16777215
         ForeColorOnMouse=   -2147483630
         BackColorOnMouse=   -2147483643
         BackColor       =   6513507
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ARButtonCtrl.ARButton ARButton8 
         Height          =   285
         Left            =   2130
         TabIndex        =   6
         Top             =   255
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         Caption         =   "&Explore"
         ForeColor       =   16777215
         ForeColorOnMouse=   -2147483630
         BackColorOnMouse=   -2147483643
         BackColor       =   6513507
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ARButtonCtrl.ARButton ARButton6 
         Height          =   285
         Left            =   3255
         TabIndex        =   7
         Top             =   255
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   503
         Caption         =   "&Minimize"
         ForeColor       =   16777215
         ForeColorOnMouse=   -2147483630
         BackColorOnMouse=   -2147483643
         BackColor       =   6513507
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ARButtonCtrl.ARButton ARButton3 
         Height          =   285
         Left            =   4395
         TabIndex        =   8
         Top             =   255
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   503
         Caption         =   "&Display Setting"
         ForeColor       =   16777215
         ForeColorOnMouse=   -2147483630
         BackColorOnMouse=   -2147483643
         BackColor       =   6513507
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ARButtonCtrl.ARButton ARButton1 
         Height          =   285
         Left            =   5880
         TabIndex        =   9
         Top             =   255
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   503
         Caption         =   "&System"
         ForeColor       =   16777215
         ForeColorOnMouse=   -2147483630
         BackColorOnMouse=   -2147483643
         BackColor       =   6513507
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ARButtonCtrl.ARButton ARButton2 
         Height          =   285
         Left            =   7380
         TabIndex        =   10
         Top             =   255
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   503
         Caption         =   "&Add-Remove Programes"
         ForeColor       =   16777215
         ForeColorOnMouse=   -2147483630
         BackColorOnMouse=   -2147483643
         BackColor       =   6513507
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   4560
      TabIndex        =   1
      Top             =   7335
      Width           =   3255
      Begin ARButtonCtrl.ARButton ARButton5 
         Height          =   285
         Left            =   105
         TabIndex        =   11
         Top             =   195
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   503
         Caption         =   "&Log Off"
         ForeColor       =   16777215
         ForeColorOnMouse=   -2147483630
         BackColorOnMouse=   -2147483643
         BackColor       =   6513507
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ARButtonCtrl.ARButton ARButton9 
         Height          =   285
         Left            =   960
         TabIndex        =   12
         Top             =   195
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   503
         Caption         =   "&Restart"
         ForeColor       =   16777215
         ForeColorOnMouse=   -2147483630
         BackColorOnMouse=   -2147483643
         BackColor       =   6513507
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ARButtonCtrl.ARButton ARButton11 
         Height          =   285
         Left            =   2115
         TabIndex        =   13
         Top             =   195
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   503
         Caption         =   "&Shut Down"
         ForeColor       =   16777215
         ForeColorOnMouse=   -2147483630
         BackColorOnMouse=   -2147483643
         BackColor       =   6513507
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   795
      Top             =   1395
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":44FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":7C55
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":B7B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":EED7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":12525
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":15CAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1965C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1CFB0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   2790
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   4921
      ButtonWidth     =   2064
      ButtonHeight    =   1588
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Student"
            Key             =   "Cust"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Attendance"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Test && Result"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Fees"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Users"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reports"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "About Us"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Calculator"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ARButton1_Click()
CompControl1.Add_HardWare
End Sub

Private Sub ARButton10_Click()
CompControl1.FindFiles
End Sub

Private Sub ARButton11_Click()
CompControl1.ShutDown
End Sub

Private Sub ARButton2_Click()
CompControl1.Add_Remove
End Sub

Private Sub ARButton3_Click()
CompControl1.Display_Settings
End Sub

Private Sub ARButton5_Click()
CompControl1.LogOff
End Sub

Private Sub ARButton6_Click()
CompControl1.MinimizeAll
End Sub

Private Sub ARButton7_Click()
CompControl1.OpenCDROM
End Sub

Private Sub ARButton8_Click()
CompControl1.OpenExplore
End Sub

Private Sub ARButton9_Click()
CompControl1.Restart
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyF2
    frmstudent.Show
Case vbKeyF3
    frmAttendance.Show
Case vbKeyF4
    frmtestresult.Show
Case vbKeyF6
    frmfees.Show
Case vbKeyF7
    frmreport.Show
Case vbKeyF8
    
'Case vbKeyF9
'        about.Show
'        about.Timer1.Enabled = False
End Select
End Sub

Private Sub Form_Load()
frmMain.Caption = "Welcome To, Upgrage Tuition " & "         Today : " & Date & " (" & Format(Date, "DDDD") & ")"
End Sub

Private Sub Form_Resize()
'Me.WindowState = 2
End Sub

Private Sub Timer1_Timer()
Label14.Caption = Time
If Label14.Caption > "4:00:00 AM" Or Label14.Caption < "10:00:00 AM" Then
Label15.Caption = "GOOD MORNING"
End If

If Label14.Caption > "10:00:00 AM" And Label14.Caption < "12:00:00 PM" Then
    Label15.Caption = "GOOD NOON"
End If

If Label14.Caption > "12:00:00 PM" And Label14.Caption < "5:00:00 PM" Then
    Label15.Caption = "GOOD AFTERNOON"
End If

If Label14.Caption > "5:00:00 PM" And Label14.Caption < "8:00:00 PM" Then
    Label15.Caption = "GOOD EVENING"
End If

If Label14.Caption > "8:00:00 PM" Or Label14.Caption < "4:00:00 AM" Then
    Label15.Caption = "GOOD NIGHT"
End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        frmstudent.Show
    Case 3
        frmAttendance.Show
    Case 5
        frmtestresult.Show
    Case 7
        frmfees.Show
    Case 9
        frmUsers.Show
    Case 11
        frmreport.Show
    Case 13
        frmabout.Show
        'about.Timer1.Enabled = False
    Case 15
        Shell "C:\WINDOWS\system32\calc.exe"
    
End Select
End Sub

