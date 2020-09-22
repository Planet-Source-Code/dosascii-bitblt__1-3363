VERSION 5.00
Begin VB.Form frmMAIN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BitBlt-Bitmap Block Transfer."
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.CommandButton cmdEXIT 
         Caption         =   "&Exit"
         Height          =   375
         Left            =   3120
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdRUN 
         Caption         =   "&BitBlt"
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Caption         =   "Destination:"
         Height          =   1575
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   4215
         Begin VB.PictureBox picDEST 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1200
            Left            =   0
            Picture         =   "frmMAIN.frx":0000
            ScaleHeight     =   1200
            ScaleWidth      =   4200
            TabIndex        =   8
            Top             =   240
            Width           =   4200
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Sources:"
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1335
         Begin VB.PictureBox picSRC1 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   720
            Picture         =   "frmMAIN.frx":106C2
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   6
            Top             =   240
            Width           =   480
         End
         Begin VB.PictureBox picSRC0 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   120
            Picture         =   "frmMAIN.frx":11304
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   5
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Label lblINFO 
         Caption         =   "Picture Boxs are being used for this program. Sorry for the crappy images."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1560
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Label lblAUT 
      Alignment       =   2  'Center
      Caption         =   "DosAscii : dosascii@hotmail.com"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   4455
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Simple BitBlt.
'DosAscii.
'dosascii@hotmail.com
'All picture boxz are set to:
'Appearance = Flat
'AutoRedraw = True
'AutoSize = True
'BorderStyle = None

Option Explicit 'This foreces all varibles to be declared now.

Public XPos 'X-axis
Public YPos 'Y-axis

Private Sub cmdEXIT_Click()
'Unload all the resources.
picSRC0.Picture = Nothing
picSRC1.Picture = Nothing
picDEST.Picture = Nothing
'Finaly, unload the app.
Unload Me
End Sub

Private Sub cmdRUN_Click()

'Colour , Colour, Mask
'Colour , Colour, Colour

' Paint the Mask onto the Destination using AND operator.
Call BitBlt(picDEST.hDC, XPos, YPos, picSRC0.ScaleWidth \ Screen.TwipsPerPixelX, picSRC0.ScaleHeight \ Screen.TwipsPerPixelY, picSRC1.hDC, 0, 0, SRCAND)

' Paint the Source onto the Destination using XOR operator.
Call BitBlt(picDEST.hDC, XPos, YPos, picSRC0.ScaleWidth \ Screen.TwipsPerPixelX, picSRC0.ScaleHeight \ Screen.TwipsPerPixelY, picSRC0.hDC, 0, 0, SRCINVERT)

' Update the screen with the updated image in memory.
picDEST.Refresh
picSRC0.Refresh
picSRC1.Refresh
End Sub

Private Sub Form_Load()
'Set the X and Y positions.
XPos = 125
YPos = 20
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Unload all the resources.
picSRC0.Picture = Nothing
picSRC1.Picture = Nothing
picDEST.Picture = Nothing
'Finaly, unload the app.
Unload Me
End Sub
