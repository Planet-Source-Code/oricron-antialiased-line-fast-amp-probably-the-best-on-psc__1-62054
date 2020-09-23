VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   8475
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   315
      Left            =   5775
      ScaleHeight     =   255
      ScaleWidth      =   855
      TabIndex        =   31
      Top             =   5475
      Width           =   915
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7425
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   5100
      Width           =   840
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5775
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   5100
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear Canvas"
      Height          =   315
      Left            =   6750
      TabIndex        =   25
      Top             =   5475
      Width           =   1590
   End
   Begin VB.PictureBox picClr 
      BackColor       =   &H00400000&
      Height          =   315
      Index           =   23
      Left            =   4500
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   24
      Top             =   5475
      Width           =   315
   End
   Begin VB.PictureBox picClr 
      BackColor       =   &H00004000&
      Height          =   315
      Index           =   22
      Left            =   4500
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   23
      Top             =   5100
      Width           =   315
   End
   Begin VB.PictureBox picClr 
      BackColor       =   &H00800000&
      Height          =   315
      Index           =   21
      Left            =   4125
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   22
      Top             =   5475
      Width           =   315
   End
   Begin VB.PictureBox picClr 
      BackColor       =   &H00008000&
      Height          =   315
      Index           =   20
      Left            =   4125
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   21
      Top             =   5100
      Width           =   315
   End
   Begin VB.PictureBox picClr 
      BackColor       =   &H00C00000&
      Height          =   315
      Index           =   19
      Left            =   3750
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   20
      Top             =   5475
      Width           =   315
   End
   Begin VB.PictureBox picClr 
      BackColor       =   &H0000C000&
      Height          =   315
      Index           =   18
      Left            =   3750
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   19
      Top             =   5100
      Width           =   315
   End
   Begin VB.PictureBox picClr 
      BackColor       =   &H00FF0000&
      Height          =   315
      Index           =   17
      Left            =   3375
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   18
      Top             =   5475
      Width           =   315
   End
   Begin VB.PictureBox picClr 
      BackColor       =   &H0000FF00&
      Height          =   315
      Index           =   16
      Left            =   3375
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   17
      Top             =   5100
      Width           =   315
   End
   Begin VB.PictureBox picClr 
      BackColor       =   &H00FF8080&
      Height          =   315
      Index           =   15
      Left            =   3000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   16
      Top             =   5475
      Width           =   315
   End
   Begin VB.PictureBox picClr 
      BackColor       =   &H0080FF80&
      Height          =   315
      Index           =   14
      Left            =   3000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   15
      Top             =   5100
      Width           =   315
   End
   Begin VB.PictureBox picClr 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Index           =   13
      Left            =   2625
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   14
      Top             =   5475
      Width           =   315
   End
   Begin VB.PictureBox picClr 
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Index           =   12
      Left            =   2625
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   13
      Top             =   5100
      Width           =   315
   End
   Begin VB.PictureBox picClr 
      BackColor       =   &H00000040&
      Height          =   315
      Index           =   11
      Left            =   2250
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   12
      Top             =   5475
      Width           =   315
   End
   Begin VB.PictureBox picClr 
      BackColor       =   &H00000000&
      Height          =   315
      Index           =   10
      Left            =   2250
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   11
      Top             =   5100
      Width           =   315
   End
   Begin VB.PictureBox picClr 
      BackColor       =   &H00000080&
      Height          =   315
      Index           =   9
      Left            =   1875
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   10
      Top             =   5475
      Width           =   315
   End
   Begin VB.PictureBox picClr 
      BackColor       =   &H00404040&
      Height          =   315
      Index           =   8
      Left            =   1875
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   9
      Top             =   5100
      Width           =   315
   End
   Begin VB.PictureBox picClr 
      BackColor       =   &H000000C0&
      Height          =   315
      Index           =   7
      Left            =   1500
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   8
      Top             =   5475
      Width           =   315
   End
   Begin VB.PictureBox picClr 
      BackColor       =   &H00808080&
      Height          =   315
      Index           =   6
      Left            =   1500
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   7
      Top             =   5100
      Width           =   315
   End
   Begin VB.PictureBox picClr 
      BackColor       =   &H000000FF&
      Height          =   315
      Index           =   5
      Left            =   1125
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   6
      Top             =   5475
      Width           =   315
   End
   Begin VB.PictureBox picClr 
      BackColor       =   &H00C0C0C0&
      Height          =   315
      Index           =   4
      Left            =   1125
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   5
      Top             =   5100
      Width           =   315
   End
   Begin VB.PictureBox picClr 
      BackColor       =   &H008080FF&
      Height          =   315
      Index           =   3
      Left            =   750
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   4
      Top             =   5475
      Width           =   315
   End
   Begin VB.PictureBox picClr 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   2
      Left            =   750
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   5100
      Width           =   315
   End
   Begin VB.PictureBox picClr 
      BackColor       =   &H00C0C0FF&
      Height          =   315
      Index           =   1
      Left            =   375
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   5475
      Width           =   315
   End
   Begin VB.PictureBox picClr 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      Left            =   375
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   5100
      Width           =   315
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4590
      Left            =   375
      ScaleHeight     =   302
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   522
      TabIndex        =   0
      Top             =   375
      Width           =   7890
      Begin VB.Shape shp 
         DrawMode        =   6  'Mask Pen Not
         Height          =   75
         Index           =   0
         Left            =   1200
         Top             =   2250
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Line ln 
         DrawMode        =   6  'Mask Pen Not
         Visible         =   0   'False
         X1              =   80
         X2              =   115
         Y1              =   140
         Y2              =   140
      End
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Color:"
      Height          =   315
      Left            =   4800
      TabIndex        =   30
      Top             =   5550
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "Opacity:"
      Height          =   315
      Left            =   6750
      TabIndex        =   29
      Top             =   5175
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Border Size:"
      Height          =   315
      Left            =   4875
      TabIndex        =   27
      Top             =   5175
      Width           =   915
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This is just a sample form...

Private Sub Form_Load()
'Here we load some setings for drawing lines

'Thicknes
Combo1.AddItem 0
Combo1.AddItem 0.05
Combo1.AddItem 0.15
Combo1.AddItem 0.25
Combo1.AddItem 0.5
Combo1.AddItem 0.75
Combo1.AddItem 1
Combo1.AddItem 2
Combo1.AddItem 3
Combo1.AddItem 4
Combo1.AddItem 5
Combo1.Text = 1

'Opacity
Combo2.AddItem "00"
Combo2.AddItem "10"
Combo2.AddItem "20"
Combo2.AddItem "30"
Combo2.AddItem "40"
Combo2.AddItem "50"
Combo2.AddItem "60"
Combo2.AddItem "70"
Combo2.AddItem "80"
Combo2.AddItem "90"
Combo2.AddItem "100"

Combo2.Text = 100
End Sub

Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'This just draws a negative line - Paint shop etc. like

If Button = vbLeftButton Or Button = vbRightButton Then
    ln.X1 = x
    ln.Y1 = y
    
    shp(0).Top = y - Int(shp(0).Height / 2)
    shp(0).Left = x - Int(shp(0).Width / 2)
    
    shp(0).Visible = True
End If

End Sub

Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Or Button = vbRightButton Then
    ln.x2 = x
    ln.Y2 = y
    
    If shp.UBound < 1 Then
        Load shp(1)
    End If
    
    shp(1).Top = y - Int(shp(1).Height / 2)
    shp(1).Left = x - Int(shp(1).Width / 2)
    
    shp(1).Visible = True
    
    ln.Visible = True
End If
End Sub

Private Sub picCanvas_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Or Button = vbRightButton Then
    ln.Visible = False
    
    If shp.UBound > 0 Then Unload shp(1)
    shp(0).Visible = False
    
    
    'ONLY THE FOLOWING LINE DRAWS
    'THE ACTUAL LINE -> YOUT NEED JUST THIS
    'COMMAND AND THE MODULE
    AALine picCanvas.hdc, ln.X1, ln.Y1, ln.x2, ln.Y2, picCanvas.ForeColor, Combo1.Text + 0, Combo2.Text / 100
    'And that's it
    
End If

End Sub


Private Sub picClr_Click(Index As Integer)
'Selects the draw color
picCanvas.ForeColor = picClr(Index).BackColor
Picture1.BackColor = picClr(Index).BackColor

End Sub


