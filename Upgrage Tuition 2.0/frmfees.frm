VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{C30897B9-75AC-11D2-94E3-000000000000}#1.4#0"; "ARButton.ocx"
Begin VB.Form frmfees 
   Appearance      =   0  'Flat
   BackColor       =   &H00636363&
   Caption         =   "Fees Detail"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8625
   Icon            =   "frmfees.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   8625
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
      Height          =   5715
      Left            =   -15
      TabIndex        =   62
      Top             =   255
      Visible         =   0   'False
      Width           =   8400
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H00636363&
         ForeColor       =   &H80000008&
         Height          =   1065
         Left            =   210
         TabIndex        =   63
         Top             =   195
         Width           =   7980
         Begin VB.TextBox txtFind 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1815
            TabIndex        =   65
            Top             =   345
            Width           =   2625
         End
         Begin VB.ComboBox cmbFind 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "frmfees.frx":08CA
            Left            =   255
            List            =   "frmfees.frx":08E0
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   345
            Width           =   1365
         End
         Begin MSComCtl2.DTPicker dtpTo 
            Height          =   315
            Left            =   6540
            TabIndex        =   66
            Top             =   345
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            Format          =   20709377
            CurrentDate     =   38704
         End
         Begin MSComCtl2.DTPicker dtpFrom 
            Height          =   315
            Left            =   4650
            TabIndex        =   67
            Top             =   345
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            Format          =   20709377
            CurrentDate     =   38704
         End
         Begin VB.Label Label24 
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
            Height          =   255
            Left            =   4650
            TabIndex        =   70
            Top             =   135
            Width           =   1335
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
            Left            =   6120
            TabIndex        =   69
            Top             =   405
            Width           =   285
         End
         Begin VB.Label Label22 
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
            Left            =   270
            TabIndex        =   68
            Top             =   150
            Width           =   795
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
         Height          =   3990
         Left            =   210
         TabIndex        =   71
         Top             =   1530
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   7038
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
         Left            =   225
         TabIndex        =   90
         Top             =   1260
         Width           =   7950
         _ExtentX        =   14023
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
   Begin VB.TextBox Text16 
      Appearance      =   0  'Flat
      BackColor       =   &H00F4F4F4&
      Enabled         =   0   'False
      Height          =   300
      Left            =   -30
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   5955
      Width           =   9555
   End
   Begin VB.TextBox Text15 
      Appearance      =   0  'Flat
      BackColor       =   &H00F4F4F4&
      Enabled         =   0   'False
      Height          =   300
      Left            =   -15
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   -45
      Width           =   9555
   End
   Begin VB.Frame Frame5 
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
      Left            =   75
      TabIndex        =   59
      Top             =   4245
      Width           =   2460
      Begin ARButtonCtrl.ARButton cmdnew 
         Height          =   330
         Left            =   75
         TabIndex        =   0
         Top             =   270
         Width           =   750
         _ExtentX        =   1323
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
         Left            =   840
         TabIndex        =   26
         Top             =   270
         Width           =   780
         _ExtentX        =   1376
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
         Left            =   1635
         TabIndex        =   28
         Top             =   270
         Width           =   765
         _ExtentX        =   1349
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
         Left            =   75
         TabIndex        =   29
         Top             =   705
         Width           =   750
         _ExtentX        =   1323
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
         Left            =   840
         TabIndex        =   30
         Top             =   705
         Width           =   780
         _ExtentX        =   1376
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
         Left            =   1635
         TabIndex        =   31
         Top             =   705
         Width           =   765
         _ExtentX        =   1349
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
   Begin VB.Frame Frame4 
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
      Left            =   2520
      TabIndex        =   58
      Top             =   4245
      Width           =   1935
      Begin ARButtonCtrl.ARButton cmdfirst 
         Height          =   345
         Left            =   75
         TabIndex        =   32
         Top             =   285
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   609
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
         Left            =   990
         TabIndex        =   33
         Top             =   285
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
         Left            =   75
         TabIndex        =   34
         Top             =   735
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
         Left            =   990
         TabIndex        =   35
         Top             =   735
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
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00636363&
      Caption         =   "Fees Detail"
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
      Height          =   4995
      Left            =   4485
      TabIndex        =   51
      Top             =   480
      Width           =   3900
      Begin VB.Frame Frame8 
         Appearance      =   0  'Flat
         BackColor       =   &H00636363&
         ForeColor       =   &H80000008&
         Height          =   4575
         Left            =   1635
         TabIndex        =   73
         Top             =   210
         Width           =   2220
         Begin VB.TextBox txtrec4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   45
            TabIndex        =   16
            Top             =   1995
            Width           =   765
         End
         Begin VB.TextBox txtrec3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   45
            TabIndex        =   14
            Top             =   1455
            Width           =   765
         End
         Begin VB.TextBox txtrec2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   45
            TabIndex        =   12
            Top             =   915
            Width           =   765
         End
         Begin VB.TextBox txtrec1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   45
            TabIndex        =   10
            Top             =   390
            Width           =   765
         End
         Begin VB.TextBox txtrec8 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   45
            TabIndex        =   25
            Top             =   4125
            Width           =   765
         End
         Begin VB.TextBox txtrec7 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   45
            TabIndex        =   22
            Top             =   3585
            Width           =   765
         End
         Begin VB.TextBox txtrec6 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   45
            TabIndex        =   20
            Top             =   3045
            Width           =   765
         End
         Begin VB.TextBox txtrec5 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   45
            TabIndex        =   18
            Top             =   2505
            Width           =   765
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   300
            Left            =   885
            TabIndex        =   11
            Top             =   390
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Format          =   20709377
            CurrentDate     =   38726
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   300
            Left            =   885
            TabIndex        =   13
            Top             =   915
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Format          =   20709377
            CurrentDate     =   38726
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   300
            Left            =   885
            TabIndex        =   15
            Top             =   1455
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Format          =   20709377
            CurrentDate     =   38726
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   300
            Left            =   885
            TabIndex        =   17
            Top             =   1995
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Format          =   20709377
            CurrentDate     =   38726
         End
         Begin MSComCtl2.DTPicker DTPicker5 
            Height          =   300
            Left            =   885
            TabIndex        =   19
            Top             =   2520
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Format          =   20709377
            CurrentDate     =   38726
         End
         Begin MSComCtl2.DTPicker DTPicker6 
            Height          =   300
            Left            =   885
            TabIndex        =   21
            Top             =   3045
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Format          =   20709377
            CurrentDate     =   38726
         End
         Begin MSComCtl2.DTPicker DTPicker7 
            Height          =   300
            Left            =   885
            TabIndex        =   24
            Top             =   3585
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Format          =   20709377
            CurrentDate     =   38726
         End
         Begin MSComCtl2.DTPicker DTPicker8 
            Height          =   300
            Left            =   885
            TabIndex        =   27
            Top             =   4125
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Format          =   20709377
            CurrentDate     =   38726
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Date-4:"
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
            Left            =   885
            TabIndex        =   89
            Top             =   1785
            Width           =   945
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Rec-4:"
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
            TabIndex        =   88
            Top             =   1785
            Width           =   945
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Date-3:"
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
            Left            =   885
            TabIndex        =   87
            Top             =   1245
            Width           =   945
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Rec-3:"
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
            TabIndex        =   86
            Top             =   1245
            Width           =   945
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Date-2:"
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
            Left            =   870
            TabIndex        =   85
            Top             =   705
            Width           =   945
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Rec-2:"
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
            TabIndex        =   84
            Top             =   705
            Width           =   945
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Date-1:"
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
            Left            =   870
            TabIndex        =   83
            Top             =   180
            Width           =   945
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Rec-1:"
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
            TabIndex        =   82
            Top             =   180
            Width           =   945
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "Date-8:"
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
            Left            =   885
            TabIndex        =   81
            Top             =   3915
            Width           =   945
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "Rec-8:"
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
            TabIndex        =   80
            Top             =   3915
            Width           =   945
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "Date-7:"
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
            Left            =   885
            TabIndex        =   79
            Top             =   3375
            Width           =   945
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "Rec-7:"
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
            TabIndex        =   78
            Top             =   3375
            Width           =   945
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "Date-6:"
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
            Left            =   870
            TabIndex        =   77
            Top             =   2835
            Width           =   945
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "Rec-6:"
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
            TabIndex        =   76
            Top             =   2835
            Width           =   945
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "Date-5:"
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
            Left            =   870
            TabIndex        =   75
            Top             =   2310
            Width           =   945
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "Rec-5:"
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
            TabIndex        =   74
            Top             =   2310
            Width           =   945
         End
      End
      Begin VB.TextBox txtfinal 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   120
         TabIndex        =   42
         Top             =   4350
         Width           =   1305
      End
      Begin VB.TextBox txtrec 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   105
         TabIndex        =   41
         Top             =   3930
         Width           =   1305
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00636363&
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   105
         TabIndex        =   52
         Top             =   210
         Width           =   1500
         Begin VB.TextBox txtofee 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   90
            TabIndex        =   8
            Top             =   1455
            Width           =   1305
         End
         Begin VB.TextBox txttotal 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   90
            TabIndex        =   9
            Top             =   1995
            Width           =   1305
         End
         Begin VB.TextBox txtmfee 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   105
            TabIndex        =   7
            Top             =   915
            Width           =   1305
         End
         Begin VB.TextBox txttfee 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   120
            TabIndex        =   6
            Top             =   390
            Width           =   1305
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Fees:"
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
            Left            =   90
            TabIndex        =   56
            Top             =   1785
            Width           =   1515
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Other Fees:"
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
            Left            =   90
            TabIndex        =   55
            Top             =   1245
            Width           =   1515
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Material Fees:"
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
            Left            =   105
            TabIndex        =   54
            Top             =   705
            Width           =   1515
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Tuition Fees:"
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
            Left            =   105
            TabIndex        =   53
            Top             =   180
            Width           =   1515
         End
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Fee*:"
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
         Left            =   90
         TabIndex        =   57
         Top             =   3720
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00636363&
      Caption         =   "About Student"
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
      Height          =   3600
      Left            =   75
      TabIndex        =   43
      Top             =   450
      Width           =   4380
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
         ItemData        =   "frmfees.frx":091D
         Left            =   3435
         List            =   "frmfees.frx":091F
         TabIndex        =   4
         Top             =   525
         Width           =   810
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
         ItemData        =   "frmfees.frx":0921
         Left            =   2595
         List            =   "frmfees.frx":0923
         TabIndex        =   3
         Top             =   525
         Width           =   765
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
         ItemData        =   "frmfees.frx":0925
         Left            =   1620
         List            =   "frmfees.frx":0927
         TabIndex        =   23
         Top             =   1005
         Width           =   2475
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
         ItemData        =   "frmfees.frx":0929
         Left            =   750
         List            =   "frmfees.frx":092B
         TabIndex        =   1
         Top             =   540
         Width           =   1020
      End
      Begin VB.TextBox txtfname 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1605
         TabIndex        =   38
         Top             =   2430
         Width           =   2505
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
         ItemData        =   "frmfees.frx":092D
         Left            =   1845
         List            =   "frmfees.frx":092F
         TabIndex        =   2
         Top             =   525
         Width           =   675
      End
      Begin VB.TextBox txtroll 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   90
         TabIndex        =   37
         Top             =   540
         Width           =   615
      End
      Begin VB.TextBox txtcontact2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         TabIndex        =   40
         Top             =   3060
         Width           =   1590
      End
      Begin VB.TextBox txtcontact1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         TabIndex        =   39
         Top             =   3045
         Width           =   1560
      End
      Begin VB.TextBox txthaddress 
         Appearance      =   0  'Flat
         Height          =   825
         Left            =   1605
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1455
         Width           =   2490
      End
      Begin VB.Label Label5 
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
         Left            =   3435
         TabIndex        =   92
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label34 
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
         Left            =   2595
         TabIndex        =   91
         Top             =   315
         Width           =   825
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Standred:"
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
         Left            =   1740
         TabIndex        =   47
         Top             =   285
         Width           =   825
      End
      Begin VB.Label Label25 
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
         Left            =   765
         TabIndex        =   72
         Top             =   315
         Width           =   780
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
         ForeColor       =   &H00F4F4F4&
         Height          =   255
         Left            =   240
         TabIndex        =   50
         Top             =   2445
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
         ForeColor       =   &H00F4F4F4&
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   2835
         Width           =   825
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
         ForeColor       =   &H00F4F4F4&
         Height          =   255
         Left            =   2505
         TabIndex        =   48
         Top             =   2850
         Width           =   915
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
         Left            =   105
         TabIndex        =   46
         Top             =   330
         Width           =   750
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
         ForeColor       =   &H00F4F4F4&
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   1425
         Width           =   1515
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Stu. Name:"
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
         Left            =   240
         TabIndex        =   44
         Top             =   1020
         Width           =   1065
      End
   End
   Begin ARButtonCtrl.ARButton cmdfind 
      Height          =   5730
      Left            =   8385
      TabIndex        =   36
      Top             =   240
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   10107
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
   Begin VB.Line Line1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   3  'Dot
      BorderWidth     =   3
      X1              =   105
      X2              =   4395
      Y1              =   4170
      Y2              =   4170
   End
End
Attribute VB_Name = "frmfees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmpRs As New ADODb.Recordset
Dim tmpGe As New ADODb.Recordset
Dim tmpMa As New ADODb.Recordset
Dim tmpMz As New ADODb.Recordset
Dim tmpSs As New ADODb.Recordset
Dim tmpSe As New ADODb.Recordset
Dim tmpSa As New ADODb.Recordset
Dim tmpSz As New ADODb.Recordset

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

Private Sub cmbsname_GotFocus()
Dim i As Integer
Changecolor True, Me.ActiveControl
cmbsname.Clear
Set tmpRs = Nothing
tmpRs.Open "select stdname from student where branch = '" & cmbbranch.Text & "' and std = '" & cmbstd.Text & "' and batch = '" & cmbbatch.Text & "' and rollno = '" & cmbroll.Text & "'", cn, adOpenKeyset, adLockReadOnly
If tmpRs.RecordCount > 0 Then
    For i = 1 To tmpRs.RecordCount
        cmbsname.AddItem IIf(IsNull(tmpRs.Fields(0)), " ", tmpRs.Fields(0))
        tmpRs.MoveNext
    Next
End If
End Sub

Private Sub cmbsname_Validate(Cancel As Boolean)
Set tmpRs = Nothing
tmpRs.Open "select haddress, fname, contact1, contact2 from student where branch = '" & cmbbranch.Text & "' and std = '" & cmbstd.Text & "' and stdname = '" & cmbsname.Text & "'", cn, adOpenKeyset, adLockReadOnly
If tmpRs.RecordCount > 0 Then
    txthaddress.Text = tmpRs.Fields(0)
    txtfname.Text = tmpRs.Fields(1)
    'txtoadd.Text = tmpRs.Fields(2)
    txtcontact1.Text = tmpRs.Fields(2)
    txtcontact2.Text = tmpRs.Fields(3)
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
        cn.Execute "delete from fees where roll = '" & txtroll.Text & "'"
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
txttfee.SetFocus
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
'    cmdexit.Enabled = False
    cmdfirst.Enabled = False
    cmdlast.Enabled = False
    cmdnext.Enabled = False
    cmdprevous.Enabled = False
End Sub

Private Sub cmdexit_Click()
Unload Me
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
set_field
Else
MsgBox "Please Insert Record", vbInformation, "Information"
End If
End Sub

Private Sub cmdlast_Click()
On Error Resume Next
If rs.EOF = False Then
rs.MoveLast
set_field
Else
MsgBox "Please Insert Record", vbInformation, "Information"
End If
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

Private Sub cmdnext_Click()
On Error Resume Next
If rs.EOF = False Then
rs.MoveNext
set_field
Else
MsgBox "This Is Last Record ", vbCritical, "Warning"
End If
End Sub

Private Sub cmdprevous_Click()
On Error Resume Next
If rs.BOF = False Then
rs.MovePrevious
set_field
Else
MsgBox "This Is First Record", vbCritical, "Worning"
End If
End Sub

Private Sub cmdsave_Click()
Dim RsCheck As New ADODb.Recordset
    Dim rssave1 As New ADODb.Recordset
    Set rs = Nothing
If Status = False Then
    Set RsCheck = Nothing
     RsCheck.Open "select * from Fees where roll = " & txtroll.Text & "", cn, adOpenKeyset, adLockReadOnly
    If RsCheck.RecordCount > 0 Then
        MsgBox "This Record is Already Entered", vbCritical, "Check It"
        Exit Sub
    End If
        rs.Open "select * from fees", cn, adOpenKeyset, adLockOptimistic
    rs.AddNew
    rs!roll = txtroll.Text
    rs!rollno = cmbroll.Text
    rs!sname = cmbsname.Text
    rs!hadd = txthaddress.Text
    rs!std = cmbstd.Text
    rs!branch = cmbbranch.Text
    rs!batch = cmbbatch.Text
    rs!fname = txtfname.Text
    'rs!oadd = txtoadd.Text
    rs!mobile = Val(txtcontact1.Text)
    rs!phone = Val(txtcontact2.Text)
    rs!rec1 = Val(txtrec1.Text)
    rs!date1 = DTPicker1.Value
    rs!rec2 = Val(txtrec2.Text)
    rs!date2 = DTPicker2.Value
    rs!rec3 = Val(txtrec3.Text)
    rs!date3 = DTPicker3.Value
    rs!rec4 = Val(txtrec4.Text)
    rs!date4 = DTPicker4.Value
    rs!rec5 = Val(txtrec5.Text)
    rs!date5 = DTPicker5.Value
    rs!rec6 = Val(txtrec6.Text)
    rs!date6 = DTPicker6.Value
    rs!rec7 = Val(txtrec7.Text)
    rs!date7 = DTPicker7.Value
    rs!rec8 = Val(txtrec8.Text)
    rs!date8 = DTPicker8.Value
    rs!tfee = Val(txttfee.Text)
    rs!mfee = Val(txtmfee.Text)
    rs!ofee = Val(txtofee.Text)
    rs!total = Val(txttotal.Text)
    rs!recfee = Val(txtrec.Text)
    rs!final = Val(txtfinal.Text)
Else
    rs.Open "select * from fees where roll = " & txtroll.Text & "", cn, adOpenKeyset, adLockOptimistic
        If rs.RecordCount > 0 Then
    rs!roll = txtroll.Text
    rs!rollno = cmbroll.Text
    rs!sname = cmbsname.Text
    rs!hadd = txthaddress.Text
    rs!std = cmbstd.Text
    rs!branch = cmbbranch.Text
    rs!batch = cmbbatch.Text
    rs!fname = txtfname.Text
    'rs!oadd = txtoadd.Text
    rs!mobile = txtcontact1.Text
    rs!phone = txtcontact2.Text
    rs!rec1 = Val(txtrec1.Text)
    rs!date1 = DTPicker1.Value
    rs!rec2 = Val(txtrec2.Text)
    rs!date2 = DTPicker2.Value
    rs!rec3 = Val(txtrec3.Text)
    rs!date3 = DTPicker3.Value
    rs!rec4 = Val(txtrec4.Text)
    rs!date4 = DTPicker4.Value
    rs!rec5 = Val(txtrec5.Text)
    rs!date5 = DTPicker5.Value
    rs!rec6 = Val(txtrec6.Text)
    rs!date6 = DTPicker6.Value
    rs!rec7 = Val(txtrec7.Text)
    rs!date7 = DTPicker7.Value
    rs!rec8 = Val(txtrec8.Text)
    rs!date8 = DTPicker8.Value
    rs!tfee = Val(txttfee.Text)
    rs!mfee = Val(txtmfee.Text)
    rs!ofee = Val(txtofee.Text)
    rs!total = txttotal.Text
    rs!recfee = txtrec.Text
    rs!final = txtfinal.Text
End If
End If
rs.Update
MsgBox "Your Record Has Been Saved Successfuly", vbInformation, "congralution"
Unload Me
Me.Show
End Sub
Public Sub set_field()
On Error Resume Next
txtroll.Text = rs.Fields(0)
cmbroll.Text = rs.Fields(1)
cmbsname.Text = rs.Fields(2)
txthaddress.Text = rs.Fields(3)
cmbstd.Text = rs.Fields(4)
cmbbranch.Text = rs.Fields(5)
cmbbatch.Text = rs.Fields(6)
txtfname.Text = rs.Fields(7)
'txtoadd.Text = rs.Fields(7)
txtcontact1.Text = rs.Fields(8)
txtcontact2.Text = rs.Fields(9)
txtrec1.Text = rs.Fields(10)
DTPicker1.Value = rs.Fields(11)
txtrec2.Text = rs.Fields(12)
DTPicker2.Value = rs.Fields(13)
txtrec3.Text = rs.Fields(14)
DTPicker3.Value = rs.Fields(15)
txtrec4.Text = rs.Fields(16)
DTPicker4.Value = rs.Fields(17)
txtrec5.Text = rs.Fields(18)
DTPicker5.Value = rs.Fields(19)
txtrec6.Text = rs.Fields(20)
DTPicker6.Value = rs.Fields(21)
txtrec7.Text = rs.Fields(22)
DTPicker7.Value = rs.Fields(23)
txtrec8.Text = rs.Fields(24)
DTPicker8.Value = rs.Fields(25)
txttfee.Text = rs.Fields(26)
txtmfee.Text = rs.Fields(27)
txtofee.Text = rs.Fields(28)
txttotal.Text = rs.Fields(29)
txtrec.Text = rs.Fields(30)
txtfinal.Text = rs.Fields(31)
End Sub

Private Sub Command3_Click()
Select Case cmbFind.ListIndex
Case 0
    If txtFind.Text = "" Then
    Adodc1.RecordSource = "select * from fees where date1>=#" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and date1 <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "#"
        Else
        Adodc1.RecordSource = "select * from fees where date1 >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and date1 <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "# and rollno like '" & txtFind.Text & "%'"
    End If
Case 1
    If txtFind.Text = "" Then
        Adodc1.RecordSource = "select * from fees where date1 >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and date1 <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "#"
    Else
        Adodc1.RecordSource = "select * from fees where date1 >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and date1 <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "# and sname like '" & txtFind.Text + "%" & "'"
    End If
Case 2
    If txtFind.Text = "" Then
        Adodc1.RecordSource = "select * from fees where date1 >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and date1 <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "#"
    Else
        Adodc1.RecordSource = "select * from fees where date1 >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and date1 <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "# and std like '" & txtFind.Text + "%" & "'"
    End If
Case 3
    If txtFind.Text = "" Then
        Adodc1.RecordSource = "select * from fees where date1 >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and date1 <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "#"
    Else
        Adodc1.RecordSource = "select * from fees where date1 >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and date1 <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "# and branch like '" & txtFind.Text + "%" & "'"
    End If
Case 4
    If txtFind.Text = "" Then
        Adodc1.RecordSource = "select * from fees where date1 >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and date1 <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "#"
    Else
        Adodc1.RecordSource = "select * from fees where date1 >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and date1 <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "# and std tfee '" & txtFind.Text + "%" & "'"
    End If
Case 5
    If txtFind.Text = "" Then
        Adodc1.RecordSource = "select * from fees where date1 >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and date1 <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "#"
    Else
        Adodc1.RecordSource = "select * from fees where date1 >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and date1 <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "# and final like '" & txtFind.Text + "%" & "'"
    End If
'Case 6
'    If txtFind.Text = "" Then
'        Adodc1.RecordSource = "select * from fees where date1 >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and date1 <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "#"
'    Else
'        Adodc1.RecordSource = "select * from fees where date1 >= #" & Format(dtpFrom.Value, "mm/dd/yyyy") & "# and date1 <= #" & Format(dtpTo.Value, "mm/dd/yyyy") & "# and pmtby like '" & txtFind.Text + "%" & "'"
'    End If
End Select
    Adodc1.Refresh
    Set DG1.DataSource = Adodc1
    DG1.ReBind
    DG1.Refresh

End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub txtcontact1_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtcontact1_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtcontact2_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtcontact2_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtfinal_GotFocus()
'txtfinal.Text = Val(txttotal.Text) - Val(txtrec.Text)
txtfinal.Enabled = False
End Sub

Private Sub txtfname_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtfname_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txthaddress_GotFocus()

Changecolor True, Me.ActiveControl

End Sub

Private Sub txthaddress_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtmfee_Change()
txttotal.Text = Val(txttfee.Text) + Val(txtmfee.Text) + Val(txtofee.Text)
final

End Sub

Private Sub txtmfee_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtmfee_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtoadd_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtoadd_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtofee_Change()
txttotal.Text = Val(txttfee.Text) + Val(txtmfee.Text) + Val(txtofee.Text)
final

End Sub

Private Sub txtofee_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtofee_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtrec_GotFocus()
'    txtrec.Text = Val(txtrec1.Text) + Val(txtrec2.Text) + Val(txtrec3.Text) + Val(txtrec4.Text) + Val(txtrec5.Text) + Val(txtrec6.Text) + Val(txtrec7.Text) + Val(txtrec8.Text)
    txtrec.Enabled = False
End Sub

Private Sub txtrec1_Change()
txtrec.Text = Val(txtrec1.Text) + Val(txtrec2.Text) + Val(txtrec3.Text) + Val(txtrec4.Text) + Val(txtrec5.Text) + Val(txtrec6.Text) + Val(txtrec7.Text) + Val(txtrec8.Text)
final
End Sub
Public Sub final()
    txtfinal.Text = Val(txttotal.Text) - Val(txtrec.Text)
End Sub

Private Sub txtrec1_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtrec1_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtrec2_Change()
txtrec.Text = Val(txtrec1.Text) + Val(txtrec2.Text) + Val(txtrec3.Text) + Val(txtrec4.Text) + Val(txtrec5.Text) + Val(txtrec6.Text) + Val(txtrec7.Text) + Val(txtrec8.Text)
final
End Sub

Private Sub txtrec2_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtrec2_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtrec3_Change()
txtrec.Text = Val(txtrec1.Text) + Val(txtrec2.Text) + Val(txtrec3.Text) + Val(txtrec4.Text) + Val(txtrec5.Text) + Val(txtrec6.Text) + Val(txtrec7.Text) + Val(txtrec8.Text)
final

End Sub

Private Sub txtrec3_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtrec3_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtrec4_Change()
txtrec.Text = Val(txtrec1.Text) + Val(txtrec2.Text) + Val(txtrec3.Text) + Val(txtrec4.Text) + Val(txtrec5.Text) + Val(txtrec6.Text) + Val(txtrec7.Text) + Val(txtrec8.Text)
final

End Sub

Private Sub txtrec4_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtrec4_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtrec5_Change()
txtrec.Text = Val(txtrec1.Text) + Val(txtrec2.Text) + Val(txtrec3.Text) + Val(txtrec4.Text) + Val(txtrec5.Text) + Val(txtrec6.Text) + Val(txtrec7.Text) + Val(txtrec8.Text)
final

End Sub

Private Sub txtrec5_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtrec5_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtrec6_Change()
txtrec.Text = Val(txtrec1.Text) + Val(txtrec2.Text) + Val(txtrec3.Text) + Val(txtrec4.Text) + Val(txtrec5.Text) + Val(txtrec6.Text) + Val(txtrec7.Text) + Val(txtrec8.Text)
final

End Sub

Private Sub txtrec6_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtrec6_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtrec7_Change()
txtrec.Text = Val(txtrec1.Text) + Val(txtrec2.Text) + Val(txtrec3.Text) + Val(txtrec4.Text) + Val(txtrec5.Text) + Val(txtrec6.Text) + Val(txtrec7.Text) + Val(txtrec8.Text)
final

End Sub

Private Sub txtrec7_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtrec7_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtrec8_Change()
txtrec.Text = Val(txtrec1.Text) + Val(txtrec2.Text) + Val(txtrec3.Text) + Val(txtrec4.Text) + Val(txtrec5.Text) + Val(txtrec6.Text) + Val(txtrec7.Text) + Val(txtrec8.Text)
final

End Sub

Private Sub txtrec8_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtrec8_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtroll_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtroll_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txtsname_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txtsname_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txttfee_Change()
txttotal.Text = Val(txttfee.Text) + Val(txtmfee.Text) + Val(txtofee.Text)
final
End Sub

Private Sub txttfee_GotFocus()
Changecolor True, Me.ActiveControl

End Sub

Private Sub txttfee_LostFocus()
Changecolor False, txtbox

End Sub

Private Sub txttotal_GotFocus()
'   txttotal.Text = Val(txttfee.Text) + Val(txtmfee.Text) + Val(txtofee.Text)
   txttotal.Enabled = False
   End Sub
Private Sub Form_Load()
Con
Set rs = Nothing
rs.Open "select * from fees order by roll", cn, adOpenKeyset, adLockOptimistic
New_NO
Disable_Fields
Adodc1.ConnectionString = cn
Adodc1.RecordSource = "SELECT * FROM Fees ORDER BY date1"
Set DG1.DataSource = Adodc1
dtpFrom.Value = Date
dtpTo.Value = Date
End Sub
Public Sub New_NO()
Dim tmpRs As New ADODb.Recordset
Set tmpRs = Nothing
If rs.RecordCount > 0 Then
    tmpRs.Open "select max(roll) from fees", cn, adOpenKeyset, adLockReadOnly
    If tmpRs.RecordCount > 0 Then
        txtroll.Text = tmpRs.Fields(0) + 1
    End If
Else
txtroll.Text = 1
End If
End Sub

