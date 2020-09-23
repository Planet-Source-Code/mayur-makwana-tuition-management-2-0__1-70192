VERSION 5.00
Begin VB.Form frmabout 
   Appearance      =   0  'Flat
   BackColor       =   &H00636363&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3900
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7005
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   3900
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4800
      Top             =   1680
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Height          =   540
      Left            =   4275
      TabIndex        =   1
      Top             =   3075
      Width           =   2715
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "ulovesme@live.com"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   150
      TabIndex        =   0
      Top             =   2970
      Width           =   2970
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Unload Me
frmMain.Show
End Sub

