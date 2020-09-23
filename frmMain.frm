VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "frmMain"
   ClientHeight    =   2835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   189
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   340
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraLCD 
      Caption         =   "kaLCD Control"
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin kaLCD.kaLCDDisplay LCD 
         Height          =   675
         Left            =   120
         Top             =   240
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   1191
         TotalCols       =   50
         TotalRows       =   5
      End
      Begin VB.CommandButton cmdLCDWriteText 
         Caption         =   "Write Text"
         Default         =   -1  'True
         Height          =   375
         Left            =   3480
         TabIndex        =   7
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtWriteText 
         Height          =   285
         Left            =   2160
         TabIndex        =   6
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtChangeSpeed 
         Height          =   285
         Left            =   2160
         MaxLength       =   3
         TabIndex        =   5
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdLCDChangeSpeed 
         Caption         =   "Change &Speed"
         Height          =   375
         Left            =   3480
         TabIndex        =   4
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdLCDClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   3480
         TabIndex        =   3
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton cmdLCDGraphicTest 
         Caption         =   "Graphic Test"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CommandButton cmdLCDAlphaNumericTest 
         Caption         =   "AlphaNumeric Test"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'1. HIve - Alternate for Collection and Dictionary (Intermediate)
'   by Omar Al Zabir
'2. Fast Graphics Filters (Intermediate)
'   by Manuel Augusto Santos

Private Sub cmdLCDAlphaNumericTest_Click()
    LCD.Mode = AlphaNumeric
    LCD.WriteSpeed = 100
    LCD.WriteText "First Line"
    LCD.WriteText "123456789012345678"
    LCD.WriteSpeed = 0
    LCD.WriteText "234567890123456789", , 4
    LCD.WriteSpeed = 100
    LCD.WriteText "345678901234567890", LCDAlignCenter
    LCD.WriteText "456789012345678901", LCDAlignRight
    LCD.WriteText "567890123456789012", , , 5
    LCD.WriteText "678901234567890123", LCDAlignRight, , 5
    LCD.WriteText "789012345678901234", , 3, 3
    LCD.WriteText "890123456789012345", , 5
    LCD.WriteText "901234567890123456"
    LCD.WriteText "Last Line"
End Sub

Private Sub cmdLCDChangeSpeed_Click()
    LCD.WriteSpeed = Val(txtChangeSpeed.Text)
End Sub

Private Sub cmdLCDClear_Click()
    LCD.Clear
End Sub

Private Sub cmdLCDGraphicTest_Click()
    LCD.Mode = Graphic
    Set LCD.Picture = LoadPicture(App.Path & "\Bart-n-God.jpg")
End Sub

Private Sub cmdLCDWriteText_Click()
    LCD.WriteText txtWriteText.Text
End Sub

Private Sub Form_Load()
    Width = (fraLCD.Width + 8) * Screen.TwipsPerPixelX
    Height = (fraLCD.Top + fraLCD.Height + 27) * Screen.TwipsPerPixelY
End Sub
