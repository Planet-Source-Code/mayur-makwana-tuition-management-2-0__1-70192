VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4050
   ClientLeft      =   4440
   ClientTop       =   3720
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   270
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   666
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2760
      Left            =   675
      Picture         =   "door.frx":0000
      ScaleHeight     =   317.242
      ScaleMode       =   0  'User
      ScaleWidth      =   302
      TabIndex        =   0
      Top             =   330
      Visible         =   0   'False
      Width           =   4530
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Height          =   450
         Left            =   2475
         TabIndex        =   2
         Top             =   1830
         Width           =   2040
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "ulovesme@live.com"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   90
         TabIndex        =   1
         Top             =   2025
         Width           =   2265
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim H, W, col As Long
    
    Me.Width = Picture1.Width * 15 'Size the Splashscreenform to
    Me.Height = Picture1.Height * 15 ' fit the Picture

    Call GetDesktop(Me) 'Make Screenshot of Screen behind frmSplash
                        'and copy it to the form
    Me.Show

    For W = 0 To Picture1.ScaleWidth 'This is the effect itself
        For H = 0 To Picture1.ScaleHeight
            If Gerade(W) = Gerade(H) Then
            ' ^ Makes sure only every second Pixel is shown
                col = GetPixel(Picture1.hdc, W, H)
                SetPixel Me.hdc, destX + W, destY + H, col
            Else
                col = GetPixel(Picture1.hdc, Picture1.ScaleWidth - W, H)
                SetPixel Me.hdc, destX + Picture1.ScaleWidth - W, destY + H, col
            End If
        Next H
    
        Me.Refresh
        DoEvents
        Pause 1 'Here you can change the speed of the effect
    Next W

    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmLogin.Show
End Sub

