VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{C30897B9-75AC-11D2-94E3-000000000000}#1.4#0"; "ARButton.ocx"
Begin VB.Form frmtestresult 
   Appearance      =   0  'Flat
   BackColor       =   &H00636363&
   Caption         =   "Test & Result"
   ClientHeight    =   5145
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9525
   Icon            =   "FRMTEST.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   9525
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Appearance      =   0  'Flat
      BackColor       =   &H00636363&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4665
      Left            =   0
      TabIndex        =   40
      Top             =   240
      Visible         =   0   'False
      Width           =   9315
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H00636363&
         ForeColor       =   &H80000008&
         Height          =   630
         Left            =   255
         TabIndex        =   41
         Top             =   180
         Width           =   8835
         Begin VB.ComboBox cmbFind 
            Height          =   315
            ItemData        =   "FRMTEST.frx":08CA
            Left            =   930
            List            =   "FRMTEST.frx":08E0
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   180
            Width           =   1365
         End
         Begin VB.TextBox txtFind 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2370
            TabIndex        =   42
            Top             =   180
            Width           =   2115
         End
         Begin MSComCtl2.DTPicker dtpTo 
            Height          =   315
            Left            =   7320
            TabIndex        =   44
            Top             =   180
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            Format          =   20709377
            CurrentDate     =   38704
         End
         Begin MSComCtl2.DTPicker dtpFrom 
            Height          =   315
            Left            =   5550
            TabIndex        =   45
            Top             =   180
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            Format          =   20709377
            CurrentDate     =   38704
         End
         Begin VB.Label Label22 
            BackColor       =   &H00636363&
            Caption         =   "Find By:"
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
            Height          =   285
            Left            =   165
            TabIndex        =   48
            Top             =   210
            Width           =   795
         End
         Begin VB.Label Label23 
            BackColor       =   &H00636363&
            Caption         =   "To"
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
            Height          =   285
            Left            =   6945
            TabIndex        =   47
            Top             =   225
            Width           =   285
         End
         Begin VB.Label Label24 
            BackColor       =   &H00636363&
            Caption         =   "From Date:"
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
            Height          =   285
            Left            =   4560
            TabIndex        =   46
            Top             =   225
            Width           =   1335
         End
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   4230
         Top             =   3060
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
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
      Begin MSDataGridLib.DataGrid Dg1 
         Height          =   3360
         Left            =   255
         TabIndex        =   49
         Top             =   1080
         Width           =   8820
         _ExtentX        =   15558
         _ExtentY        =   5927
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin ARButtonCtrl.ARButton Command3 
         Height          =   285
         Left            =   270
         TabIndex        =   51
         Top             =   810
         Width           =   8805
         _ExtentX        =   15531
         _ExtentY        =   503
         Caption         =   "&Go"
         ForeColor       =   16053492
         ForeColorOnMouse=   6513507
         BackColorOnMouse=   16053492
         BackColor       =   6513507
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
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H00F4F4F4&
      Height          =   300
      Left            =   -15
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   -45
      Width           =   9555
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      Height          =   330
      Left            =   -1485
      TabIndex        =   38
      Top             =   7050
      Width           =   9180
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      BackColor       =   &H00F4F4F4&
      Height          =   300
      Left            =   0
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   4890
      Width           =   9555
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00636363&
      Caption         =   "Commands"
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
      Height          =   1230
      Left            =   1260
      TabIndex        =   35
      Top             =   3345
      Width           =   3855
      Begin ARButtonCtrl.ARButton cmdnew 
         Height          =   330
         Left            =   240
         TabIndex        =   0
         Top             =   300
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   582
         Caption         =   "&New"
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
      Begin ARButtonCtrl.ARButton cmdedit 
         Height          =   330
         Left            =   1455
         TabIndex        =   14
         Top             =   300
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   582
         Caption         =   "&Edit"
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
         Height          =   330
         Left            =   2715
         TabIndex        =   13
         Top             =   300
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   582
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
      Begin ARButtonCtrl.ARButton cmdcancel 
         Height          =   330
         Left            =   240
         TabIndex        =   15
         Top             =   735
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   582
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
      Begin ARButtonCtrl.ARButton cmddelete 
         Height          =   330
         Left            =   1455
         TabIndex        =   16
         Top             =   735
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   582
         Caption         =   "&Delete"
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
         Height          =   330
         Left            =   2715
         TabIndex        =   17
         Top             =   735
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   582
         Caption         =   "e&Xit"
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
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00636363&
      Caption         =   "Goto"
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
      Height          =   1230
      Left            =   5085
      TabIndex        =   34
      Top             =   3345
      Width           =   2670
      Begin ARButtonCtrl.ARButton cmdfirst 
         Height          =   330
         Left            =   255
         TabIndex        =   18
         Top             =   300
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   582
         Caption         =   "&First"
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
      Begin ARButtonCtrl.ARButton cmdlast 
         Height          =   330
         Left            =   1530
         TabIndex        =   19
         Top             =   300
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   582
         Caption         =   "&Last"
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
      Begin ARButtonCtrl.ARButton cmdnext 
         Height          =   330
         Left            =   255
         TabIndex        =   20
         Top             =   750
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   582
         Caption         =   "nex&T"
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
      Begin ARButtonCtrl.ARButton cmdprevous 
         Height          =   330
         Left            =   1530
         TabIndex        =   21
         Top             =   750
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   582
         Caption         =   "&Prevous"
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
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00636363&
      Caption         =   "Test Detail"
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
      Height          =   1755
      Left            =   180
      TabIndex        =   25
      Top             =   1260
      Width           =   9000
      Begin VB.TextBox txtremark 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1035
         TabIndex        =   12
         Top             =   1305
         Width           =   7695
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00636363&
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   8520
         Begin VB.TextBox txtch 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4710
            TabIndex        =   9
            Top             =   465
            Width           =   870
         End
         Begin VB.ComboBox cmbsub 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "FRMTEST.frx":091F
            Left            =   3210
            List            =   "FRMTEST.frx":0926
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   465
            Width           =   1440
         End
         Begin VB.TextBox txtpf 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   7530
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   465
            Width           =   885
         End
         Begin VB.TextBox txtper 
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
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   6825
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   465
            Width           =   675
         End
         Begin VB.TextBox txttotal 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6255
            TabIndex        =   11
            Top             =   465
            Width           =   540
         End
         Begin VB.TextBox txtobtain 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5610
            TabIndex        =   10
            Top             =   465
            Width           =   615
         End
         Begin MSComCtl2.DTPicker DTPtest 
            Height          =   315
            Left            =   1545
            TabIndex        =   7
            Top             =   465
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            Format          =   20709377
            CurrentDate     =   38726
         End
         Begin VB.ComboBox cmbtype 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "FRMTEST.frx":0933
            Left            =   60
            List            =   "FRMTEST.frx":0940
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   465
            Width           =   1455
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Chapter:"
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
            Left            =   4725
            TabIndex        =   58
            Top             =   255
            Width           =   750
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Pass/Fail"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   7515
            TabIndex        =   33
            Top             =   270
            Width           =   1020
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Per.(%):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   6825
            TabIndex        =   32
            Top             =   255
            Width           =   1020
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
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
            Left            =   6255
            TabIndex        =   31
            Top             =   255
            Width           =   810
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Obtain:"
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
            Left            =   5610
            TabIndex        =   30
            Top             =   255
            Width           =   645
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Subject:"
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
            Left            =   3210
            TabIndex        =   29
            Top             =   240
            Width           =   1020
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Test Date:"
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
            Left            =   1545
            TabIndex        =   28
            Top             =   255
            Width           =   1020
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Test Type:"
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
            Left            =   60
            TabIndex        =   27
            Top             =   240
            Width           =   1020
         End
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Remark:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   1350
         Width           =   1020
      End
   End
   Begin VB.TextBox txtroll 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      TabIndex        =   23
      Top             =   570
      Width           =   780
   End
   Begin ARButtonCtrl.ARButton cmdfind 
      Height          =   4695
      Left            =   9315
      TabIndex        =   22
      Top             =   240
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   8281
      Caption         =   "&F"
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
      ItemData        =   "FRMTEST.frx":095A
      Left            =   2130
      List            =   "FRMTEST.frx":095C
      TabIndex        =   2
      Top             =   555
      Width           =   1005
   End
   Begin VB.ComboBox cmbsname 
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
      ItemData        =   "FRMTEST.frx":095E
      Left            =   5460
      List            =   "FRMTEST.frx":0960
      TabIndex        =   5
      Top             =   555
      Width           =   3660
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
      ItemData        =   "FRMTEST.frx":0962
      Left            =   3240
      List            =   "FRMTEST.frx":0964
      TabIndex        =   3
      Top             =   555
      Width           =   1005
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
      ItemData        =   "FRMTEST.frx":0966
      Left            =   1020
      List            =   "FRMTEST.frx":0968
      TabIndex        =   1
      Top             =   555
      Width           =   1005
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
      ItemData        =   "FRMTEST.frx":096A
      Left            =   4365
      List            =   "FRMTEST.frx":096C
      TabIndex        =   4
      Top             =   555
      Width           =   960
   End
   Begin VB.Label Label9 
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
      Left            =   4350
      TabIndex        =   57
      Top             =   360
      Width           =   735
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
      Left            =   3240
      TabIndex        =   55
      Top             =   360
      Width           =   825
   End
   Begin VB.Label Label19 
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
      Left            =   2145
      TabIndex        =   54
      Top             =   375
      Width           =   825
   End
   Begin VB.Label Label18 
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
      Left            =   1020
      TabIndex        =   53
      Top             =   375
      Width           =   825
   End
   Begin VB.Label Label1 
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
      Left            =   5460
      TabIndex        =   52
      Top             =   375
      Width           =   825
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Sr. No:"
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
      Left            =   165
      TabIndex        =   24
      Top             =   375
      Width           =   750
   End
End
Attribute VB_Name = "frmtestresult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmpRs As New ADODb.Recordset
Dim tmpGe As New ADODb.Recordset
Dim tmpMa As New ADODb.Recordset
Dim tmpMz As New ADODb.Recordset

Private Sub cmbbatch_GotFocus()
Dim i As Integer
Changecolor True, Me.ActiveControl
cmbbatch.Clear
Set tmpRs = Nothing
tmpRs.Open "select distinct batch from student where branch = '" & cmbbranch.Text & "' and std = '" & cmbstd.Text & "'", cn, adOpenKeyset, adLockReadOnly
If tmpRs.RecordCount > 0 Then
    For i = 1 To tmpRs.RecordCount
        cmbbatch.AddItem IIf(IsNull(tmpRs.Fields(0)), " ", tmpRs.Fields(0))
        tmpRs.MoveNext
    Next
End If
End Sub

Private Sub cmbbatch_LostFocus()
Changecolor False, txtbox

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

Private Sub cmbbranch_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub cmbroll_GotFocus()
Dim i As Integer
Changecolor True, Me.ActiveControl
cmbroll.Clear
Set tmpRs = Nothing
tmpRs.Open "select distinct rollno from student where batch = '" & cmbbatch.Text & "' and branch = '" & cmbbranch.Text & "' and std = '" & cmbstd.Text & "'", cn, adOpenKeyset, adLockReadOnly
If tmpRs.RecordCount > 0 Then
    For i = 1 To tmpRs.RecordCount
        cmbroll.AddItem IIf(IsNull(tmpRs.Fields(0)), " ", tmpRs.Fields(0))
        tmpRs.MoveNext
    Next
End If
End Sub

Private Sub cmbroll_LostFocus()
Changecolor False, txtbox
End Sub

Private Sub cmbstd_GotFocus()
Dim i As Integer
Changecolor True, Me.ActiveControl
cmbstd.Clear
Set tmpRs = Nothing
tmpRs.Open "select distinct std from student where branch = '" & cmbbranch.Text & "'", cn, adOpenKeyset, adLockReadOnly
If tmpRs.RecordCount > 0 Then
    For i = 1 To tmpRs.RecordCount
        cmbstd.AddItem IIf(IsNull(tmpRs.Fields(0)), " ", tmpRs.Fields(0))
        tmpRs.MoveNext
    Next
End If

End Sub

Private Sub cmbstd_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub cmbsub_GotFocus()
Changecolor True, Me.ActiveControl
End Sub

Private Sub cmbsub_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub cmbtype_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub cmbtype_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub cmdcancel_Click()
Unload Me
Me.Show
End Sub

Private Sub cmddelete_Click()
If MsgBox("Are You Sure to Delete This Record?", vbYesNo, "Confirmation") = vbYes Then
    If txtroll.Text <> "" Then
        cn.Execute "delete from testresult where roll = " & txtroll.Text & ""
        Unload Me
        Me.Show
    Else
        MsgBox "Please Select Record First", vbCritical, "Check It"
    End If
End If
End Sub

Private Sub cmdedit_Click()
addedit_state
Status = True
enable_fields
'cmbsname.SetFocus
End Sub

Private Sub cmdfind_Click()
    If cmdfind.Caption = "&F" Then
    Frame7.Visible = True
    Frame7.ZOrder 0
    cmdfind.Caption = "&H"
    enable_fields
ElseIf cmdfind.Caption = "&H" Then
    Frame7.Visible = False
    cmdfind.Caption = "&F"
    Disable_Fields
End If
End Sub

Private Sub cmdfirst_Click()
    On Error Resume Next
If rs.BOF = False Then
rs.MoveFirst
set_fields
Else
MsgBox "Please Insert Record", vbInformation, "Information"
End If
End Sub

Private Sub cmdlast_Click()
    On Error Resume Next
If rs.EOF = False Then
rs.MoveLast
set_fields
Else
MsgBox "Please Insert Record", vbInformation, "Information"
End If
End Sub
Public Sub set_fields()
    On Error Resume Next
    txtroll.Text = rs.Fields(0)
    cmbroll.Text = rs.Fields(1)
    cmbbranch.Text = rs.Fields(2)
    cmbstd.Text = rs.Fields(3)
    cmbbatch.Text = rs.Fields(4)
    cmbsname.Text = rs.Fields(5)
    cmbtype.Text = rs.Fields(6)
    DTPtest.Value = rs.Fields(7)
    cmbsub.Text = rs.Fields(8)
    txtobtain.Text = rs.Fields(9)
    txttotal.Text = rs.Fields(10)
    txtper.Text = rs.Fields(11)
    txtpf.Text = rs.Fields(12)
    txtremark.Text = rs.Fields(13)
    txtch.Text = rs.Fields(14)
    End Sub

Private Sub cmdnew_Click()
Unload Me
Me.Show
addedit_state
enable_fields
txtroll.Enabled = True
Status = False
cmbbranch.SetFocus
End Sub
Public Sub addedit_state()
    cmdnew.Enabled = False
    cmdedit.Enabled = False
    cmdsave.Enabled = True
    cmdcancel.Enabled = True
    cmddelete.Enabled = False
    cmdexit.Enabled = False
    cmdfirst.Enabled = False
    cmdlast.Enabled = False
    cmdnext.Enabled = False
    cmdprevous.Enabled = False
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
        ElseIf TypeOf ctrl Is ListView Then
            ctrl.Enabled = True
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
        ElseIf TypeOf ctrl Is ListView Then
            ctrl.Enabled = False
        'ElseIf TypeOf ctrl Is MSFlexGrid Then
        '    ctrl.Clear
        End If
    Next
End Sub
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
                
Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdnext_Click()
    On Error Resume Next
If rs.EOF = False Then
rs.MoveNext
set_fields
Else
MsgBox "This Is Last Record", vbCritical, "Cheak it"
End If
End Sub

Private Sub cmdprevous_Click()
    On Error Resume Next
If rs.BOF = False Then
rs.MovePrevious
set_fields
Else
MsgBox "This Is First Record", vbCritical, "Cheak it"
End If
End Sub

Private Sub cmdsave_Click()
Dim RsCheck As New ADODb.Recordset
    Dim rssave1 As New ADODb.Recordset
 
Set rs = Nothing
If Status = False Then
    Set RsCheck = Nothing
    RsCheck.Open "select * from testresult where roll = " & txtroll.Text & "", cn, adOpenKeyset, adLockReadOnly
    If RsCheck.RecordCount > 0 Then
        MsgBox "This Record is Already Entered", vbCritical, "Check It"
        Exit Sub
    End If
    rs.Open "select * from testresult", cn, adOpenKeyset, adLockOptimistic
    rs.AddNew
    rs!roll = txtroll.Text
    rs!branch = cmbbranch.Text
    rs!std = cmbstd.Text
    rs!batch = cmbbatch.Text
    rs!Name = cmbsname.Text
    rs!testtype = cmbtype.Text
    rs!testdate = DTPtest.Value
    rs!subject = cmbsub.Text
    rs!ch = txtch.Text
    rs!obtain = Val(txtobtain.Text)
    rs!total = Val(txttotal.Text)
    rs!per = Val(txtper.Text)
    rs!passfail = txtpf.Text
    rs!remark = txtremark.Text
    rs!rollno = cmbroll.Text
Else
    rs.Open "select * from testresult where roll = " & txtroll.Text & "", cn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount > 0 Then
    rs!roll = txtroll.Text
    rs!branch = cmbbranch.Text
    rs!std = cmbstd.Text
    rs!batch = cmbbatch.Text
    rs!Name = cmbsname.Text
    rs!testtype = cmbtype.Text
    rs!testdate = DTPtest.Value
    rs!subject = cmbsub.Text
    rs!ch = txtch.Text
    rs!obtain = Val(txtobtain.Text)
    rs!total = Val(txttotal.Text)
    rs!per = Val(txtper.Text)
    rs!passfail = txtpf.Text
    rs!remark = txtremark.Text
    rs!rollno = cmbroll.Text
End If
End If
    rs.Update
MsgBox "Your Record Has Been Saved Successfuly", vbInformation, "Congrtulation"
Unload Me
Me.Show
End Sub



Private Sub Command3_Click()
Select Case cmbFind.ListIndex
Case 0
    If txtFind.Text = "" Then
        Adodc1.RecordSource = "select * from testresult where testdate >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and testdate <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "#"
    Else
        Adodc1.RecordSource = "select * from testresult where testdate >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and testdate <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "# and name like '" & txtFind.Text & "%'"
    End If
Case 1
    If txtFind.Text = "" Then
        Adodc1.RecordSource = "select * from testresult where testdate >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and testdate <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "#"
    Else
        Adodc1.RecordSource = "select * from testresult where testdate >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and testdate <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "# and rollno like '" & txtFind.Text + "%" & "'"
    End If
Case 2
    If txtFind.Text = "" Then
        Adodc1.RecordSource = "select * from testresult where testdate >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and testdate <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "#"
    Else
        Adodc1.RecordSource = "select * from testresult where testdate >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and testdate <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "# and std like '" & txtFind.Text + "%" & "'"
    End If
Case 3
    If txtFind.Text = "" Then
        Adodc1.RecordSource = "select * from testresult where testdate >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and testdate <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "#"
    Else
        Adodc1.RecordSource = "select * from testresult where testdate >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and testdate <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "# and passfail like '" & txtFind.Text + "%" & "'"
    End If
Case 4
    If txtFind.Text = "" Then
        Adodc1.RecordSource = "select * from testresult where testdate >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and testdate <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "#"
    Else
        Adodc1.RecordSource = "select * from testresult where testdate >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and testdate <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "# and testtype like '" & txtFind.Text + "%" & "'"
    End If
Case 5
    If txtFind.Text = "" Then
        Adodc1.RecordSource = "select * from testresult where testdate >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and testdate <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "#"
    Else
        Adodc1.RecordSource = "select * from testresult where testdate >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and testdate <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "# and remark like '" & txtFind.Text + "%" & "'"
    End If
End Select
    Adodc1.Refresh
    Set DG1.DataSource = Adodc1
    DG1.ReBind
    DG1.Refresh

End Sub

Private Sub DTPtest_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub DTPtest_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub Form_Load()
Con
Set rs = Nothing
rs.Open "select * from testresult order by roll", cn, adOpenKeyset, adLockOptimistic
New_NO
Disable_Fields
Adodc1.ConnectionString = cn
Adodc1.RecordSource = "SELECT * FROM testresult ORDER BY testdate"
Set DG1.DataSource = Adodc1
dtpFrom.Value = Date
dtpTo.Value = Date
End Sub
Public Sub New_NO()
Dim tmpRs As New ADODb.Recordset
Set tmpRs = Nothing
If rs.RecordCount > 0 Then
    tmpRs.Open "select max(roll) from testresult", cn, adOpenKeyset, adLockReadOnly
    If tmpRs.RecordCount > 0 Then
        txtroll.Text = tmpRs.Fields(0) + 1
    End If
Else
txtroll.Text = 1
End If
End Sub

Private Sub txtcon1_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtcon1_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtcon2_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtcon2_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtch_GotFocus()
Changecolor True, Me.ActiveControl
End Sub

Private Sub txtch_LostFocus()
Changecolor False, txtbox
End Sub

Private Sub txtobtain_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtobtain_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtremark_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtremark_LostFocus()
Changecolor False, txtbox

End Sub


Private Sub txtroll_GotFocus()
Dim i As Integer
Changecolor True, Me.ActiveControl
cmbsname.Clear
Set tmpRs = Nothing
tmpRs.Open "select stdname from student where branch = '" & cmbbranch.Text & "' and std = '" & cmbstd.Text & "'", cn, adOpenKeyset, adLockReadOnly
If tmpRs.RecordCount > 0 Then
    For i = 1 To tmpRs.RecordCount
        cmbsname.AddItem tmpRs.Fields(0)
        tmpRs.MoveNext
    Next
End If

End Sub

Private Sub txtroll_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtschool_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtschool_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtscon1_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtscon1_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtscon2_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtscon2_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub cmbsname_GotFocus()
Dim i As Integer
Changecolor True, Me.ActiveControl
cmbsname.Clear
Set tmpRs = Nothing
tmpRs.Open "select stdname from student where branch = '" & cmbbranch.Text & "' and std = '" & cmbstd.Text & "' and batch = '" & cmbbatch.Text & "' and rollno = '" & cmbroll.Text & "'", cn, adOpenKeyset, adLockReadOnly
If tmpRs.RecordCount > 0 Then
    For i = 1 To tmpRs.RecordCount
        cmbsname.AddItem tmpRs.Fields(0)
        tmpRs.MoveNext
    Next
End If

End Sub

Private Sub cmbsname_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txttotal_Change()
On Error Resume Next
txtper.Text = (100 * Val(txtobtain.Text)) / Val(txttotal.Text)
End Sub

Private Sub txtper_Change()
txtper.Enabled = False
If Val(txtper.Text) < 35 Then
txtpf.Text = "Fail"
Else
txtpf.Text = "Pass"
End If
End Sub

Private Sub txtpf_Click()
txtpf.Enabled = False
End Sub

Private Sub txtobtain_Change()
'txtper.Text = (100 * Val(txtobtain.Text)) / Val(txttotal.Text)
End Sub

Private Sub txttotal_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txttotal_LostFocus()
Changecolor False, txtbox

End Sub
