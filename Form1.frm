VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Volcanoe Simulator 2"
   ClientHeight    =   5715
   ClientLeft      =   1935
   ClientTop       =   1725
   ClientWidth     =   7950
   DrawWidth       =   3
   FillColor       =   &H000000C0&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   381
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   530
   Begin VB.CommandButton Command3 
      Caption         =   "About"
      Height          =   255
      Left            =   1680
      TabIndex        =   16
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Continue Flow"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   4920
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   5400
      Width           =   1455
   End
   Begin VB.PictureBox fountainm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4500
      Left            =   2640
      Picture         =   "Form1.frx":030A
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   1
      Top             =   5760
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.PictureBox fountains 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   375
      Picture         =   "Form1.frx":3748C
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   0
      Top             =   5865
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.Timer Timer1 
      Left            =   3960
      Top             =   4200
   End
   Begin VB.PictureBox board 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000040C0&
      BorderStyle     =   0  'None
      DrawWidth       =   10
      Height          =   4545
      Left            =   120
      ScaleHeight     =   300
      ScaleMode       =   0  'User
      ScaleWidth      =   250
      TabIndex        =   2
      Top             =   360
      Width           =   3750
   End
   Begin VB.PictureBox buffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000040C0&
      BorderStyle     =   0  'None
      DrawWidth       =   10
      Height          =   4575
      Left            =   4080
      ScaleHeight     =   308.081
      ScaleMode       =   0  'User
      ScaleWidth      =   250
      TabIndex        =   4
      Top             =   360
      Width           =   3750
   End
   Begin VB.Frame Frame2 
      Caption         =   "Speed"
      Height          =   735
      Left            =   4080
      TabIndex        =   10
      Top             =   4920
      Width           =   975
      Begin MSComctlLib.Slider sldSpeed 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         Min             =   1
         SelStart        =   9
         Value           =   9
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Outflow"
      Height          =   735
      Left            =   3000
      TabIndex        =   8
      Top             =   4920
      Width           =   975
      Begin MSComctlLib.Slider sldFlow 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   3
         Min             =   1
         SelStart        =   3
         Value           =   3
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Output Width"
      Height          =   735
      Left            =   5160
      TabIndex        =   12
      Top             =   4920
      Width           =   1095
      Begin MSComctlLib.Slider sldWidth 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393216
         Min             =   1
         SelStart        =   5
         Value           =   5
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Pressure"
      Height          =   735
      Left            =   6360
      TabIndex        =   14
      Top             =   4920
      Width           =   1455
      Begin MSComctlLib.Slider sldPressure 
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Min             =   1
         Max             =   30
         SelStart        =   15
         Value           =   15
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Mountain Terrain:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Your View:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim running As Long

Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
Select Case Command2.Caption
Case "Stop Flow"
running = False
Command2.Caption = "Continue Flow"
Case "Continue Flow"
running = True
Command2.Caption = "Stop Flow"
End Select
End Sub

Private Sub Command3_Click()
MsgBox ("Volcanoe Simulator 2 by Kevin Fleet" & vbCrLf & "A Program that simulates a volcanoe (sortof)" & vbCrLf & "Copyright(R) 2002 KevCom"), vbOKOnly, "About Volcanoe Simulator"
End Sub

Private Sub Form_Load()
Randomize

running = False
Randomize
board.picture = Nothing
Me.Show
Dim c As Long, i As Long
Do
If c > 10000 - (sldSpeed * 1000) And GetTickCount > 300 Then
c = 0

board.Cls
buffer.Cls
BitBlt buffer.hDC, buffer.ScaleWidth \ 2 - (250 \ 2), 0, 250, 300, fountainm.hDC, 0, 0, SRCAND
BitBlt board.hDC, board.ScaleWidth \ 2 - (250 \ 2), 0, 250, 300, fountainm.hDC, 0, 0, SRCAND
BitBlt board.hDC, board.ScaleWidth \ 2 - (250 \ 2), board.ScaleHeight - 100, 250, 100, fountains.hDC, 0, 0, SRCINVERT

For i = 0 To 200
If P(i).Act = True Then
board.PSet (P(i).x, P(i).y), P(i).Color
If P(i).Hit >= 15 Then buffer.PSet (P(i).x, P(i).y), vbBlack
End If
Next i

If running = True Then
For i = 0 To (sldFlow.Value - 1)
MakeDroplet board
Next i
End If

MoveDrops board

Else
c = c + 1
End If
DoEvents
Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
