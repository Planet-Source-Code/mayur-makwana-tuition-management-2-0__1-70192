VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{C30897B9-75AC-11D2-94E3-000000000000}#1.4#0"; "ARButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmstudent 
   Appearance      =   0  'Flat
   BackColor       =   &H00636363&
   Caption         =   "Student's Detail"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8880
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin ARButtonCtrl.ARButton cmdfind 
      Height          =   5535
      Left            =   8640
      TabIndex        =   25
      Top             =   240
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   9763
      Caption         =   "&F"
      ForeColor       =   6513507
      ForeColorOnMouse=   16053492
      ForeColorOnFocus=   16053492
      BackColorOnMouse=   6513507
      BackColor       =   16777215
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
      Height          =   5520
      Left            =   0
      TabIndex        =   47
      Top             =   240
      Visible         =   0   'False
      Width           =   8610
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00636363&
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   255
         TabIndex        =   49
         Top             =   150
         Width           =   8145
         Begin VB.ComboBox cmbFind 
            Height          =   315
            ItemData        =   "Form1.frx":08CA
            Left            =   210
            List            =   "Form1.frx":08E3
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   360
            Width           =   1365
         End
         Begin VB.TextBox txtFind 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1770
            TabIndex        =   29
            Top             =   360
            Width           =   2805
         End
         Begin MSComCtl2.DTPicker dtpTo 
            Height          =   315
            Left            =   6660
            TabIndex        =   31
            Top             =   360
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            Format          =   16187393
            CurrentDate     =   38704
         End
         Begin MSComCtl2.DTPicker dtpFrom 
            Height          =   315
            Left            =   4785
            TabIndex        =   30
            Top             =   360
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            Format          =   16187393
            CurrentDate     =   38704
         End
         Begin VB.Label Label20 
            BackColor       =   &H00636363&
            Caption         =   "Find By"
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
            Left            =   210
            TabIndex        =   52
            Top             =   135
            Width           =   795
         End
         Begin VB.Label Label18 
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
            Left            =   6240
            TabIndex        =   51
            Top             =   405
            Width           =   285
         End
         Begin VB.Label Label19 
            BackColor       =   &H00636363&
            Caption         =   "From Date :"
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
            Left            =   4785
            TabIndex        =   50
            Top             =   135
            Width           =   1335
         End
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   1500
         Top             =   3090
         Visible         =   0   'False
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
      Begin MSDataGridLib.DataGrid DG1 
         Height          =   4095
         Left            =   255
         TabIndex        =   48
         Top             =   1245
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   7223
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
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
         Height          =   270
         Left            =   270
         TabIndex        =   55
         Top             =   990
         Width           =   8130
         _ExtentX        =   14340
         _ExtentY        =   476
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
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   345
      Left            =   -30
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   -90
      Width           =   8985
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   -75
      Locked          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   5760
      Width           =   9180
   End
   Begin VB.Frame Frame2 
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
      Left            =   4890
      TabIndex        =   44
      Top             =   4290
      Width           =   2670
      Begin ARButtonCtrl.ARButton cmdfirst 
         Height          =   330
         Left            =   315
         TabIndex        =   21
         Top             =   315
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
         Left            =   1590
         TabIndex        =   22
         Top             =   315
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
         Left            =   315
         TabIndex        =   23
         Top             =   765
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
         Left            =   1590
         TabIndex        =   24
         Top             =   765
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
      Left            =   1080
      TabIndex        =   43
      Top             =   4290
      Width           =   3855
      Begin ARButtonCtrl.ARButton cmdnew 
         Height          =   330
         Left            =   270
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
         Left            =   1485
         TabIndex        =   17
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
         Left            =   2745
         TabIndex        =   16
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
         Left            =   270
         TabIndex        =   18
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
         Left            =   1485
         TabIndex        =   19
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
         Left            =   2745
         TabIndex        =   20
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
   Begin VB.Frame frmstudent 
      Appearance      =   0  'Flat
      BackColor       =   &H00636363&
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   225
      TabIndex        =   27
      Top             =   420
      Width           =   8325
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
         ItemData        =   "Form1.frx":0927
         Left            =   4695
         List            =   "Form1.frx":0929
         TabIndex        =   14
         Top             =   3300
         Width           =   795
      End
      Begin VB.TextBox txtrollno 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1020
         TabIndex        =   1
         Top             =   555
         Width           =   1005
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Select Image"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5970
         TabIndex        =   15
         Top             =   3300
         Width           =   1890
      End
      Begin VB.TextBox txtschoolname 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   210
         TabIndex        =   8
         Top             =   2520
         Width           =   2730
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
         ItemData        =   "Form1.frx":092B
         Left            =   2460
         List            =   "Form1.frx":0935
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3300
         Width           =   1170
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
         ItemData        =   "Form1.frx":094B
         Left            =   3735
         List            =   "Form1.frx":0955
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   3300
         Width           =   795
      End
      Begin VB.TextBox txtsname 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2145
         TabIndex        =   2
         Top             =   555
         Width           =   2145
      End
      Begin VB.TextBox txtfname 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4395
         TabIndex        =   3
         Top             =   555
         Width           =   1995
      End
      Begin VB.TextBox txtsurname 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6495
         TabIndex        =   4
         Top             =   555
         Width           =   1650
      End
      Begin VB.TextBox txthadd 
         Appearance      =   0  'Flat
         Height          =   810
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1290
         Width           =   2910
      End
      Begin VB.TextBox txtcon2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3405
         TabIndex        =   7
         Top             =   1830
         Width           =   1770
      End
      Begin VB.TextBox txtcon1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3405
         TabIndex        =   6
         Top             =   1305
         Width           =   1770
      End
      Begin VB.TextBox txtdesig 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3165
         TabIndex        =   9
         Top             =   2520
         Width           =   2010
      End
      Begin VB.TextBox txtsector 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   225
         TabIndex        =   10
         Top             =   3330
         Width           =   750
      End
      Begin VB.TextBox txtroll 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   225
         TabIndex        =   26
         Top             =   555
         Width           =   690
      End
      Begin MSComCtl2.DTPicker DTPadd 
         Height          =   315
         Left            =   1065
         TabIndex        =   11
         Top             =   3300
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16187393
         CurrentDate     =   38723
      End
      Begin VB.Label Label12 
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
         ForeColor       =   &H00F0F0F0&
         Height          =   255
         Left            =   4680
         TabIndex        =   58
         Top             =   3105
         Width           =   825
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
         ForeColor       =   &H00F0F0F0&
         Height          =   255
         Left            =   1020
         TabIndex        =   57
         Top             =   345
         Width           =   750
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "------Photograph------"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   6060
         TabIndex        =   56
         Top             =   900
         Width           =   1680
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   2070
         Left            =   5970
         Stretch         =   -1  'True
         Top             =   1155
         Width           =   1860
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "School's Name"
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
         Height          =   255
         Left            =   210
         TabIndex        =   54
         Top             =   2295
         Width           =   2250
      End
      Begin VB.Label Label17 
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
         ForeColor       =   &H00F0F0F0&
         Height          =   255
         Left            =   2460
         TabIndex        =   53
         Top             =   3105
         Width           =   825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Student's Name:"
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
         Height          =   255
         Left            =   2145
         TabIndex        =   42
         Top             =   345
         Width           =   1515
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Father's Name:"
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
         Height          =   255
         Left            =   4395
         TabIndex        =   41
         Top             =   345
         Width           =   1515
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Surname:"
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
         Height          =   255
         Left            =   6480
         TabIndex        =   40
         Top             =   360
         Width           =   1515
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Home Address:"
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
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile:"
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
         Height          =   255
         Left            =   3405
         TabIndex        =   38
         Top             =   1635
         Width           =   1515
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone:"
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
         Height          =   255
         Left            =   3405
         TabIndex        =   37
         Top             =   1110
         Width           =   1515
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Designation of Father:"
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
         Height          =   255
         Left            =   3165
         TabIndex        =   36
         Top             =   2310
         Width           =   1950
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Sector:"
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
         Height          =   255
         Left            =   225
         TabIndex        =   35
         Top             =   3120
         Width           =   780
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "A'mision Date:"
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
         Height          =   255
         Left            =   1080
         TabIndex        =   34
         Top             =   3075
         Width           =   1470
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
         ForeColor       =   &H00F0F0F0&
         Height          =   255
         Left            =   225
         TabIndex        =   33
         Top             =   345
         Width           =   750
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
         ForeColor       =   &H00F0F0F0&
         Height          =   255
         Left            =   3735
         TabIndex        =   32
         Top             =   3090
         Width           =   825
      End
   End
   Begin MSComDlg.CommonDialog Cd 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmstudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IPath As String
Dim tmpRs As New ADODb.Recordset

Private Sub cmbbatch_GotFocus()
Dim i As Integer
Changecolor True, Me.ActiveControl
cmbbatch.Clear
Set tmpRs = Nothing
tmpRs.Open "select distinct batch from student", cn, adOpenKeyset, adLockReadOnly
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
Changecolor True, Me.ActiveControl

End Sub

Private Sub cmbbranch_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub cmbstd_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub cmbstd_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub cmdcancel_Click()
Unload Me
Me.Show
End Sub

Private Sub cmddelete_Click()
If MsgBox("Are You Sure to Delete This Record?", vbYesNo, "Confirmation") = vbYes Then
    If txtroll.Text <> "" Then
        cn.Execute "delete from student where roll = " & txtroll.Text & ""
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
txtrollno.SetFocus
txtroll.Enabled = False
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
Public Sub addedit_state()
    cmdnew.Enabled = False
    cmdedit.Enabled = False
    cmdsave.Enabled = True
    cmddelete.Enabled = False
    cmdcancel.Enabled = True
    cmdexit.Enabled = False
    cmdfirst.Enabled = False
    cmdlast.Enabled = False
    cmdnext.Enabled = False
    cmdprevous.Enabled = False
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

Private Sub cmdexit_Click()
Unload Me
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
Public Sub set_fields()
On Error Resume Next
txtroll.Text = rs.Fields(0)
txtrollno.Text = rs.Fields(1)
txtsname.Text = rs.Fields(2)
txtfname.Text = rs.Fields(3)
txtsurname.Text = rs.Fields(4)
txthadd.Text = rs.Fields(5)
txtcon1.Text = rs.Fields(6)
txtcon2.Text = rs.Fields(7)
DTPadd.Value = rs.Fields(8)
txtdesig.Text = rs.Fields(9)
txtsector.Text = rs.Fields(10)
cmbbranch.Text = rs.Fields(11)
cmbstd.Text = rs.Fields(12)
txtschoolname.Text = rs.Fields(13)
Image1.Picture = LoadPicture(IIf(IsNull(rs.Fields(14)), "", rs.Fields(14)))
cmbbatch.Text = rs.Fields(15)
End Sub

Private Sub cmdlast_Click()
'On Error Resume Next
If rs.EOF = False Then
rs.MoveLast
set_fields
Else
MsgBox "Please Insert Record", vbInformation, "Information"
End If
End Sub

Private Sub cmdnew_Click()
cmdcancel_Click
addedit_state
enable_fields
txtroll.Enabled = False
Status = False
txtrollno.SetFocus
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
    Set tmpRs = Nothing
    tmpRs.Open "Select * from student where rollno = '" & txtrollno.Text & "'", cn, adOpenKeyset, adLockReadOnly
    If tmpRs.RecordCount > 0 Then
        MsgBox " This rollno is already assigned to  " & tmpRs.Fields("stdname"), vbCritical, "Note.."
        Exit Sub
    End If
    rs.Open "select * from student", cn, adOpenKeyset, adLockOptimistic
    rs.AddNew
    rs!roll = txtroll.Text
    rs!rollno = txtrollno.Text
    rs!stdname = txtsname.Text
    rs!fname = txtfname.Text
    rs!surname = txtsurname.Text
    rs!haddress = txthadd.Text
    rs!contact1 = txtcon1.Text
    rs!contact2 = txtcon2.Text
    rs!addmisiondate = DTPadd.Value
    rs!schoolname = txtschoolname.Text
    rs!fatherdesg = txtdesig.Text
    rs!sector = txtsector.Text
    rs!branch = cmbbranch.Text
    rs!std = cmbstd.Text
    rs!ImgPath = IPath
    rs!batch = cmbbatch.Text
Else
    rs.Open "select * from student where roll = " & txtroll.Text & "", cn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount > 0 Then
        rs!roll = txtroll.Text
        rs!rollno = txtrollno.Text
        rs!stdname = txtsname.Text
        rs!fname = txtfname.Text
        rs!surname = txtsurname.Text
        rs!haddress = txthadd.Text
        rs!contact1 = txtcon1.Text
        rs!contact2 = txtcon2.Text
        rs!addmisiondate = DTPadd.Value
        rs!schoolname = txtschoolname.Text
        rs!fatherdesg = txtdesig.Text
        rs!sector = txtsector.Text
        rs!branch = cmbbranch.Text
        rs!std = cmbstd.Text
        rs!ImgPath = IPath
        rs!batch = cmbbatch.Text
End If
End If
    rs.Update
MsgBox "Your Record Has Been Saved Successfuly", vbInformation, "Congrtulation"
IPath = ""
Unload Me
Me.Show
End Sub

Private Sub Command1_Click()
On Error Resume Next
Cd.CancelError = True
Err.Clear
If txtsname.Text <> "" Then
    Cd.DialogTitle = "Select Image For Item [" & txtsname.Text & "]"
    Cd.Filter = "All Picture Files (*.bmp, *.gif, *.jpg, *.jpeg)|*.bmp;*.gif;*.jpg;*.jprg"
    Cd.ShowOpen
    If Err.Number = 0 Then
        If Cd.FileName <> "" Then
            Image1.Picture = LoadPicture(Cd.FileName)
            IPath = Cd.FileName
        End If
    End If
Else
    MsgBox "You Must Write Student Name", vbCritical, "Check It"
End If
End Sub

Private Sub Command3_Click()
Select Case cmbFind.ListIndex
Case 0
    If txtFind.Text = "" Then
        Adodc1.RecordSource = "select * from student where addmisiondate >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and addmisiondate <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "#"
    Else
        Adodc1.RecordSource = "select * from student where addmisiondate >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and addmisiondate <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "# and stdname like '" & txtFind.Text & "%'"
    End If
Case 1
    If txtFind.Text = "" Then
        Adodc1.RecordSource = "select * from student where addmisiondate >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and addmisiondate <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "#"
    Else
        Adodc1.RecordSource = "select * from student where addmisiondate >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and addmisiondate <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "# and surname like '" & txtFind.Text + "%" & "'"
    End If
Case 2
    If txtFind.Text = "" Then
        Adodc1.RecordSource = "select * from student where addmisiondate >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and addmisiondate <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "#"
    Else
        Adodc1.RecordSource = "select * from student where addmisiondate >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and addmisiondate <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "# and sector like '" & txtFind.Text + "%" & "'"
    End If
Case 3
    If txtFind.Text = "" Then
        Adodc1.RecordSource = "select * from student where addmisiondate >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and addmisiondate <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "#"
    Else
        Adodc1.RecordSource = "select * from student where addmisiondate >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and addmisiondate <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "# and rollno like '" & txtFind.Text + "%" & "'"
    End If
Case 4
    If txtFind.Text = "" Then
        Adodc1.RecordSource = "select * from student where addmisiondate >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and addmisiondate <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "#"
    Else
        Adodc1.RecordSource = "select * from student where addmisiondate >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and addmisiondate <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "# and std like '" & txtFind.Text + "%" & "'"
    End If
Case 5
    If txtFind.Text = "" Then
        Adodc1.RecordSource = "select * from student where addmisiondate >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and addmisiondate <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "#"
    Else
        Adodc1.RecordSource = "select * from student where addmisiondate >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and addmisiondate <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "# and schoolname like '" & txtFind.Text + "%" & "'"
    End If
Case 6
    If txtFind.Text = "" Then
        Adodc1.RecordSource = "select * from student where addmisiondate >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and addmisiondate <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "#"
    Else
        Adodc1.RecordSource = "select * from student where addmisiondate >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and addmisiondate <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "# and branch like '" & txtFind.Text + "%" & "'"
    End If
End Select
    Adodc1.Refresh
    Set DG1.DataSource = Adodc1
    DG1.ReBind
    DG1.Refresh
End Sub

Private Sub Dg1_Click()
On Error Resume Next
Set rs = Nothing
rs.Open "select * from student where roll = " & DG1.Columns(0).Text & "", cn, adOpenKeyset, adLockReadOnly
If rs.RecordCount > 0 Then
    set_fields
    Disable_Fields
    FormLoad_State
End If

End Sub

Private Sub DTPadd_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub DTPadd_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub DTPbirth_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub DTPbirth_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub Form_Load()
Con
Set rs = Nothing
rs.Open "select * from student order by roll", cn, adOpenKeyset, adLockOptimistic
New_NO
Disable_Fields
Adodc1.ConnectionString = cn
Adodc1.RecordSource = "SELECT * FROM student ORDER BY addmisiondate"
Set DG1.DataSource = Adodc1
dtpFrom.Value = Date
dtpTo.Value = Date
Image1.Picture = LoadPicture("")
End Sub
Public Sub New_NO()
Dim tmpRs As New ADODb.Recordset
Set tmpRs = Nothing
If rs.RecordCount > 0 Then
    tmpRs.Open "select max(roll) from student", cn, adOpenKeyset, adLockReadOnly
    If tmpRs.RecordCount > 0 Then
        txtroll.Text = tmpRs.Fields(0) + 1
    End If
Else
txtroll.Text = 1
End If
End Sub

Private Sub txtannual_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtannual_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtb1std_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtb1std_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtb2std_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtb2std_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtb3std_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtb3std_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtbro_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtbro_LostFocus()
Changecolor False, txtbox

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

Private Sub txtdesig_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtdesig_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtfname_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtfname_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txthadd_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txthadd_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtlast_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtlast_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtoadd_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtoadd_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtroll_GotFocus()
Changecolor True, Me.ActiveControl
End Sub

Private Sub txtroll_LostFocus()
Changecolor False, txtbox
End Sub

Private Sub txts1std_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txts1std_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txts2std_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txts2std_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txts3std_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txts3std_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtrollno_GotFocus()
Changecolor True, Me.ActiveControl
End Sub

Private Sub txtrollno_LostFocus()
Changecolor False, txtbox
End Sub

Private Sub txtschoolname_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtschoolname_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtsector_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtsector_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtsis_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtsis_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtsname_GotFocus()
Changecolor True, Me.ActiveControl
End Sub

Private Sub txtsname_LostFocus()
Changecolor False, txtbox
End Sub

Private Sub txtsurname_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtsurname_LostFocus()
Changecolor False, txtbox

End Sub

