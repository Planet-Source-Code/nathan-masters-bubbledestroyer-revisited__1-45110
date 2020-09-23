VERSION 5.00
Begin VB.Form frmBack 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   FillColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picHeatBack 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   0
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   109
      TabIndex        =   6
      Top             =   600
      Width           =   1695
      Begin VB.PictureBox picHeatBar 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   7
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.PictureBox picLifeBack 
      BackColor       =   &H00FFFF00&
      Height          =   255
      Left            =   0
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   109
      TabIndex        =   0
      Top             =   0
      Width           =   1695
      Begin VB.PictureBox picLifeBar 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   1
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.Label lblInstruct 
      BackStyle       =   0  'Transparent
      Caption         =   "Pause: Esc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   10
      Left            =   0
      TabIndex        =   41
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblLevelUp 
      BackStyle       =   0  'Transparent
      Caption         =   "heat Recovery"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   14
      Left            =   10440
      TabIndex        =   40
      Top             =   6000
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblLevelUp 
      BackStyle       =   0  'Transparent
      Caption         =   "0: Quicker Over-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   12
      Left            =   9960
      TabIndex        =   39
      Top             =   5760
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   7080
      TabIndex        =   38
      Top             =   8400
      Width           =   3090
   End
   Begin VB.Label lblSpeed 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   7080
      TabIndex        =   37
      Top             =   8040
      Width           =   3090
   End
   Begin VB.Label lblInstruct 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click any of the instrucions to hide. Hit the ""i"" key to show them again."
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Index           =   9
      Left            =   120
      TabIndex        =   36
      Top             =   6720
      Width           =   2535
   End
   Begin VB.Label lblInstruct 
      BackStyle       =   0  'Transparent
      Caption         =   "Toggle Quad: ""1"" (NumPad)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   7
      Left            =   0
      TabIndex        =   35
      Top             =   6120
      Width           =   2535
   End
   Begin VB.Label lblInstruct 
      BackStyle       =   0  'Transparent
      Caption         =   "Quad Rotation: ""2"" (NumPad)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   8
      Left            =   0
      TabIndex        =   34
      Top             =   6480
      Width           =   2655
   End
   Begin VB.Label lblInstruct 
      BackStyle       =   0  'Transparent
      Caption         =   "Accelerate: Up Key"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   33
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblInstruct 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dampers: ""0"" on the Number Pad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Index           =   1
      Left            =   0
      TabIndex        =   32
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label lblInstruct 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00E0E0E0&
      Height          =   1335
      Index           =   2
      Left            =   0
      TabIndex        =   31
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label lblInstruct 
      BackStyle       =   0  'Transparent
      Caption         =   "Brake/Reverse: Down Key"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   30
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label lblInstruct 
      BackStyle       =   0  'Transparent
      Caption         =   "Fire: Ctrl Key"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   29
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label lblInstruct 
      BackStyle       =   0  'Transparent
      Caption         =   "Toggle Weapons: Shift Key"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   28
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Label lblInstruct 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reverse(Dampers On): Page Down Key"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Index           =   6
      Left            =   -120
      TabIndex        =   27
      Top             =   5520
      Width           =   2415
   End
   Begin VB.Label lblLevelUp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   360
      Index           =   13
      Left            =   11520
      TabIndex        =   26
      Top             =   1200
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label lblLevelUp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "The following apply only to ammunition fired from the curretly selected gun."
      ForeColor       =   &H00E0E0E0&
      Height          =   615
      Index           =   5
      Left            =   9840
      TabIndex        =   25
      Top             =   3480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblLevelUp 
      BackStyle       =   0  'Transparent
      Caption         =   "4: Faser ROF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   6
      Left            =   9960
      TabIndex        =   24
      Top             =   4200
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblLevelUp 
      BackStyle       =   0  'Transparent
      Caption         =   "5: Slower Heating"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   7
      Left            =   9960
      TabIndex        =   23
      Top             =   4440
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblLevelUp 
      BackStyle       =   0  'Transparent
      Caption         =   "7: Larger Bullet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   9
      Left            =   9960
      TabIndex        =   22
      Top             =   4920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblLevelUp 
      BackStyle       =   0  'Transparent
      Caption         =   "8: Higher Velocity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   10
      Left            =   9960
      TabIndex        =   21
      Top             =   5160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblLevelUp 
      BackStyle       =   0  'Transparent
      Caption         =   "6: Longer Range"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   8
      Left            =   9960
      TabIndex        =   20
      Top             =   4680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblLevelUp 
      BackStyle       =   0  'Transparent
      Caption         =   "9: Lower Velocity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   11
      Left            =   9960
      TabIndex        =   19
      Top             =   5400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblLevelUp 
      BackStyle       =   0  'Transparent
      Caption         =   "3: Regen Rate + 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   4
      Left            =   9960
      TabIndex        =   18
      Top             =   3120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblLevelUp 
      BackStyle       =   0  'Transparent
      Caption         =   "2: Max Heat +  5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   3
      Left            =   9960
      TabIndex        =   17
      Top             =   2880
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblLevelUp 
      BackStyle       =   0  'Transparent
      Caption         =   "1: Max Life + 10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   2
      Left            =   9960
      TabIndex        =   16
      Top             =   2640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblLevelUp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hit the number on your keyboard (not the number pad) that corresponds to the skill you wish to increase."
      ForeColor       =   &H00E0E0E0&
      Height          =   975
      Index           =   1
      Left            =   9840
      TabIndex        =   15
      Top             =   1680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblLevelUp 
      BackStyle       =   0  'Transparent
      Caption         =   "Level Up!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   0
      Left            =   9840
      TabIndex        =   14
      Top             =   1200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblQuadRotation 
      BackStyle       =   0  'Transparent
      Caption         =   "QuadRotation: Off"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   8640
      Width           =   2055
   End
   Begin VB.Label lblLevel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   9240
      TabIndex        =   12
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Level:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   8520
      TabIndex        =   11
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   9240
      TabIndex        =   10
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   8520
      TabIndex        =   9
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblGunHeat 
      BackStyle       =   0  'Transparent
      Caption         =   "Gun Heat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblLife 
      BackStyle       =   0  'Transparent
      Caption         =   "Life"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblGunMode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GunMode: Normal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   2760
      TabIndex        =   4
      Top             =   7920
      Width           =   2175
   End
   Begin VB.Label lblDampers 
      BackStyle       =   0  'Transparent
      Caption         =   "Dampers: Off"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Label lblQuadCannons 
      BackStyle       =   0  'Transparent
      Caption         =   "QuadCannons: Off"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   8280
      Width           =   2055
   End
End
Attribute VB_Name = "frmBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_GotFocus()
frmMain.Show
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Call frmMain.Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Call frmMain.Form_KeyUp(KeyCode, Shift)
End Sub


Private Sub Label3_Click()

End Sub

Private Sub lblDampers_Click()
Call frmMain.Form_KeyDown(96, 0)
End Sub

Private Sub lblGunMode_Click()
Call frmMain.Form_KeyDown(98, 0)
End Sub

Private Sub lblInstruct_Click(Index As Integer)
Dim i As Integer


For i = 0 To (lblInstruct.Count - 1)
    lblInstruct(i).Visible = False
Next i

    
End Sub

Private Sub lblQuadCannons_Click()
Call frmMain.Form_KeyDown(97, 0)
End Sub

Private Sub lblQuadRotation_Click()
Call frmMain.Form_KeyDown(100, 0)
End Sub

