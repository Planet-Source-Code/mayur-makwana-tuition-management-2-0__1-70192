VERSION 5.00
Object = "{C30897B9-75AC-11D2-94E3-000000000000}#1.4#0"; "ARButton.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmreport 
   Appearance      =   0  'Flat
   BackColor       =   &H00636363&
   Caption         =   "Report"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7095
   Icon            =   "frmreport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   7095
   StartUpPosition =   1  'CenterOwner
   Begin Crystal.CrystalReport CR1 
      Left            =   960
      Top             =   5190
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00636363&
      Caption         =   "Attendance Wise"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F4F4F4&
      Height          =   855
      Left            =   210
      TabIndex        =   3
      Top             =   1695
      Width           =   6675
      Begin ARButtonCtrl.ARButton Command16 
         Height          =   360
         Left            =   120
         TabIndex        =   10
         Top             =   300
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   635
         Caption         =   "Status Wise"
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
      Begin ARButtonCtrl.ARButton Command19 
         Height          =   360
         Left            =   2280
         TabIndex        =   11
         Top             =   300
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   635
         Caption         =   "Date Wise"
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
      Begin ARButtonCtrl.ARButton Command18 
         Height          =   360
         Left            =   4440
         TabIndex        =   12
         Top             =   300
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   635
         Caption         =   "Batch Wise"
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
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00636363&
      Caption         =   "Fees Detail"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F4F4F4&
      Height          =   855
      Left            =   210
      TabIndex        =   2
      Top             =   3990
      Width           =   6675
      Begin ARButtonCtrl.ARButton Command17 
         Height          =   360
         Left            =   135
         TabIndex        =   19
         Top             =   300
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   635
         Caption         =   "Debit Fees Wise"
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
      Begin ARButtonCtrl.ARButton Command12 
         Height          =   360
         Left            =   2295
         TabIndex        =   20
         Top             =   300
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   635
         Caption         =   "Standard Wise"
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
      Begin ARButtonCtrl.ARButton Command13 
         Height          =   360
         Left            =   4470
         TabIndex        =   21
         Top             =   300
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   635
         Caption         =   "Branch Wise"
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
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00636363&
      Caption         =   "Student Detail"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F4F4F4&
      Height          =   1380
      Left            =   210
      TabIndex        =   1
      Top             =   255
      Width           =   6660
      Begin ARButtonCtrl.ARButton Command5 
         Height          =   360
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   635
         Caption         =   "Standard Wise"
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
      Begin ARButtonCtrl.ARButton Command3 
         Height          =   360
         Left            =   2280
         TabIndex        =   5
         Top             =   360
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   635
         Caption         =   "Sector Wise"
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
      Begin ARButtonCtrl.ARButton Command4 
         Height          =   360
         Left            =   4455
         TabIndex        =   6
         Top             =   360
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   635
         Caption         =   "Surname Wise"
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
      Begin ARButtonCtrl.ARButton Command6 
         Height          =   360
         Left            =   120
         TabIndex        =   7
         Top             =   780
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   635
         Caption         =   "Branch Wise"
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
      Begin ARButtonCtrl.ARButton Command8 
         Height          =   360
         Left            =   2280
         TabIndex        =   8
         Top             =   780
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   635
         Caption         =   "School Wise"
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
      Begin ARButtonCtrl.ARButton Command7 
         Height          =   360
         Left            =   4455
         TabIndex        =   9
         Top             =   780
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   635
         Caption         =   "Addmision Date Wise"
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
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00636363&
      Caption         =   "Test && Result"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F4F4F4&
      Height          =   1320
      Left            =   210
      TabIndex        =   0
      Top             =   2610
      Width           =   6675
      Begin ARButtonCtrl.ARButton cmdschool 
         Height          =   360
         Left            =   135
         TabIndex        =   13
         Top             =   330
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   635
         Caption         =   "Chapter Wise"
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
      Begin ARButtonCtrl.ARButton Command1 
         Height          =   360
         Left            =   2295
         TabIndex        =   14
         Top             =   330
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   635
         Caption         =   "Test Date Wise"
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
      Begin ARButtonCtrl.ARButton Command2 
         Height          =   360
         Left            =   4470
         TabIndex        =   15
         Top             =   330
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   635
         Caption         =   "Subject Wise"
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
      Begin ARButtonCtrl.ARButton Command11 
         Height          =   360
         Left            =   135
         TabIndex        =   16
         Top             =   765
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   635
         Caption         =   "Standard Wise"
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
      Begin ARButtonCtrl.ARButton Command10 
         Height          =   360
         Left            =   2295
         TabIndex        =   17
         Top             =   765
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   635
         Caption         =   "Roll No. Wise"
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
      Begin ARButtonCtrl.ARButton Command9 
         Height          =   360
         Left            =   4470
         TabIndex        =   18
         Top             =   765
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   635
         Caption         =   "Result Graph"
         ForeColor       =   6513507
         ForeColorOnMouse=   16053492
         ForeColorOnFocus=   16053492
         BackColorOnMouse=   6513507
         BackColor       =   12632256
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
   End
   Begin ARButtonCtrl.ARButton Command14 
      Height          =   360
      Left            =   2025
      TabIndex        =   22
      Top             =   5070
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   635
      Caption         =   "&Refresh"
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
   Begin ARButtonCtrl.ARButton Command15 
      Height          =   360
      Left            =   3570
      TabIndex        =   23
      Top             =   5070
      Width           =   1470
      _ExtentX        =   2593
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
End
Attribute VB_Name = "frmreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdschool_Click()
CR1.Connect = "TUITION"
CR1.ReportFileName = App.Path + "\REPORTS\Test & Result\Chapter Wise.RPT"
CR1.Action = 1
End Sub

Private Sub Command1_Click()
CR1.Connect = "TUITION"
CR1.ReportFileName = App.Path + "\REPORTS\Test & Result\Test Date Wise.RPT"
CR1.Action = 1
End Sub

Private Sub Command10_Click()
CR1.Connect = "TUITION"
CR1.ReportFileName = App.Path + "\REPORTS\Test & Result\Roll No. Wise.RPT"
CR1.Action = 1
End Sub

Private Sub Command11_Click()
CR1.Connect = "TUITION"
CR1.ReportFileName = App.Path + "\REPORTS\Test & Result\Standard Wise.RPT"
CR1.Action = 1
End Sub

Private Sub Command12_Click()
CR1.Connect = "TUITION"
CR1.ReportFileName = App.Path + "\REPORTS\Fees Detail\Standard Wise.RPT"
CR1.Action = 1
End Sub

Private Sub Command13_Click()
CR1.Connect = "TUITION"
CR1.ReportFileName = App.Path + "\REPORTS\Fees Detail\Branch Wise.RPT"
CR1.Action = 1
End Sub

Private Sub Command14_Click()
Unload Me
Me.Show
End Sub

Private Sub Command15_Click()
Unload Me
End Sub

Private Sub Command16_Click()
CR1.Connect = "TUITION"
CR1.ReportFileName = App.Path + "\REPORTS\Attendance\Status Wise.RPT"
CR1.Action = 1
End Sub

Private Sub Command17_Click()
CR1.Connect = "TUITION"
CR1.ReportFileName = App.Path + "\REPORTS\Fees Detail\Debit Fees Wise.RPT"
CR1.Action = 1
End Sub

Private Sub Command18_Click()
CR1.Connect = "TUITION"
CR1.ReportFileName = App.Path + "\REPORTS\Attendance\Batch Wise.RPT"
CR1.Action = 1
End Sub

Private Sub Command19_Click()
CR1.Connect = "TUITION"
CR1.ReportFileName = App.Path + "\REPORTS\Attendance\Att.Date Wise.RPT"
CR1.Action = 1
End Sub

Private Sub Command2_Click()
CR1.Connect = "TUITION"
CR1.ReportFileName = App.Path + "\REPORTS\Test & Result\Subject Wise.RPT"
CR1.Action = 1
End Sub

Private Sub Command3_Click()
CR1.Connect = "TUITION"
CR1.ReportFileName = App.Path + "\REPORTS\Student\Sector Wise.RPT"
CR1.Action = 1
End Sub

Private Sub Command4_Click()
CR1.Connect = "TUITION"
CR1.ReportFileName = App.Path + "\REPORTS\Student\Surname Wise.RPT"
CR1.Action = 1
End Sub

Private Sub Command5_Click()
CR1.Connect = "TUITION"
CR1.ReportFileName = App.Path + "\REPORTS\Student\Standard Wise.RPT"
CR1.DiscardSavedData = True
CR1.Action = 1
End Sub

Private Sub Command6_Click()
CR1.Connect = "TUITION"
CR1.ReportFileName = App.Path + "\REPORTS\Student\Branch Wise.RPT"
CR1.Action = 1
End Sub

Private Sub Command7_Click()
CR1.Connect = "TUITION"
CR1.ReportFileName = App.Path + "\REPORTS\Student\Addmission Date Wise.RPT"
CR1.DiscardSavedData = True
CR1.Action = 1
End Sub

Private Sub Command8_Click()
CR1.Connect = "TUITION"
CR1.ReportFileName = App.Path + "\REPORTS\Student\School Wise.RPT"
CR1.Action = 1
End Sub

Private Sub Command9_Click()
CR1.Connect = "TUITION"
CR1.ReportFileName = App.Path + "\REPORTS\Test & Result\Graph.RPT"
CR1.Action = 1
End Sub

