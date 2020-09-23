VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6750
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   450
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgSaveLoad 
      Left            =   6240
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Saved Game Files(*.sav)|*.sav|"
      Flags           =   2
   End
   Begin VB.PictureBox picBackBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   6240
      ScaleHeight     =   449
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   449
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.Label lblCHEATED 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "***Cheated***"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   1830
      TabIndex        =   10
      Top             =   1320
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.Label lblDateSet 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/01"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label lblRecordHolder 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nobody"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   240
      TabIndex        =   8
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblHighScore 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label lblHS 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "High Score:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   0
      TabIndex        =   6
      Top             =   1680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblNewGame 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "New Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   2400
      TabIndex        =   5
      Top             =   5040
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Exit Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   2640
      TabIndex        =   4
      Top             =   6120
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label lblResume 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Resume Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   1635
      TabIndex        =   3
      Top             =   5520
      Visible         =   0   'False
      Width           =   3405
   End
   Begin VB.Label lblLoad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Load Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   2040
      TabIndex        =   2
      Top             =   3720
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.Label lblSave 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Save Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   2040
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.Label lblPaused 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Paused"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   2520
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   1725
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
  
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

Private Declare Function MoveToEx Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINT_TYPE) As Long

Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long



'The FileNumbers
Const SAVEGAMEFILE = 1
Const HIGHSCOREFILE = 2
Const ENCRYPTFILE = 3
Const RECORDSIZE = 42

Const ENCRYPTKEY = 85

Const FDELAY = 25
Const ROT_RATE = 8
Const SHIP_RAD = 8
Const QUADROTATIONSPEED = 0.3
Const PiBy2 = 3.14159 * 2
Const Pi180 = 3.14159 / 180
Const Pi3 = 3.14159 / 3
Const Pi2 = 3.14159 / 2
Const ACCRATE = 0.12
Const BRAKERATE = 0.12
Const EBRAKERATE = 5
Const BASEREGENRATE = 0.04
Const SHIPLNG = 5
Const SHIPFWDTH = 1.5
Const SHIPRWDTH = 3
Const TOTALBUBBLE = 10
Const MINBUBBLERAD = 8
Const MAXBUBBLERAD = 20
Const MAXBUBBLESPEED = 5
Const BASEBUBBLEREGENDELAY = 5000
Const MINDISTFROMSHIP = 200
Const ENEMYACCRATE = 0.08
Const ENEMYDESIREDDIST = 100
Const ENEMYDESIREDSPEED = 2


Const nmNumpadZero = 96
Const nmNumpadOne = 97
Const nmNumpadFour = 100

Dim ShipX1 As Single    'Coords of the ship
Dim ShipY1 As Single
Dim ShipX2 As Single
Dim ShipY2 As Single
Dim ShipX3 As Single
Dim ShipY3 As Single
Dim ShipX4 As Single
Dim ShipY4 As Single
Dim Trail(2) As TRAIL_TYPE
Dim CurTrail As Single


'The variables for gun style
Dim SHOTSPEED As Single
Dim BULLETRAD As Single
Dim SHOTDELAY As Single     ' The delay between shots
Dim BULLETLIFESPAN As Single    'How long each bullet will lBubble
Dim GUNCOOLRATE As Single   'The rate at which the gun heat (normally) decreases
Dim GUNHEATRATE As Single   'The rate at which the gun heat increases



Dim BonusLevels As Single
Dim NEWBUBBLEDELAY As Single
Dim Level As Single
Dim MainDC As Long
Dim Score As Double
Dim LeftKey As Boolean, RightKey As Boolean
Dim UpKey As Boolean, Reverse As Boolean
Dim Ebrakes As Boolean      'The End Button
Dim Brakes As Boolean       'The PDown Button
Dim Shoot As Boolean        'The Ctrl Button
Dim Dampers As Boolean      'The 0 (zero) Button
Dim QuadCannons As Boolean  'Use all four points on the Ship for shooting
Dim QuadRotation As Boolean   'Do the QuadCannons rotate?
Dim QuadSpread As Single
Dim GunMode As Byte         'The Current Style of the Ship
Dim CurHeat As Single
Dim HeatBarColor As Long    'Keeps track of the color, so it can flash when overheating
Dim Running As Boolean
Dim InGameTimer As Single
Dim ShipFacing As Single
Dim ShipHeading As Single
Dim ShipSpeed As Single
Dim ShipCenterX As Single, ShipCenterY As Single
Dim CurShotDel As Single    'keeps track of the mill.s since lBubble shot
Dim CurDelay As Single      ' The time in mil.sec. from lBubble frame
Dim CurLife As Single
Dim CurFrame As Single
Dim OverHeated As Boolean
Dim OverHeatTime As Single
Dim ColorChangedTime As Single
Dim BubbleRegen As Boolean
Dim GameTime As Long
Dim Paused As Boolean
Dim CHEATED As Boolean


Dim MAXHEAT As Single
Dim MAXLIFE As Single
Dim SHIPREGENRATE As Single
Dim OVERHEATWAITTIME As Single     'the recool time when the guns overheat


Dim CurCannon As Single
Dim Bullet() As BulletType                 'An array of bullets
Dim CurBullet As Single



Dim Bubble() As BubbleType
Dim Enemy() As ENEMY_TYPE


Dim GunStyle() As GUNSTYLE_TYPE
Dim EnemyBullet() As BulletType

Dim EnemyGun As GUNSTYLE_TYPE
Dim EnemyCurBullet As Single
Dim ENEMYLIFE As Single
Dim ENEMYRADIUS As Single

Dim ULTIMATEGUNACCESS As Byte
Private Type BulletType
    x As Single
    y As Single
    Heading As Single
    Speed As Single
    Radius As Single
    HitsTaken As Single
    Active As Boolean
    Lifespan As Single
    TimeInFlight As Single
    BulletName As String
End Type

Private Type TRAIL_TYPE
    x As Single
    y As Single
    Heading As Single
    Speed As Single
End Type

Private Type BubbleType
    x As Single
    y As Single
    Heading As Single
    Speed As Single
    Radius As Single
    Collision As Boolean
    Active As Boolean
    Debris(8) As TRAIL_TYPE
End Type

Private Type POINT_TYPE
  x As Long
  y As Long
End Type


Private Type ENEMY_TYPE
    x As Single
    y As Single
    Heading As Single
    Facing As Single
    Speed As Single
    CurShotDelay As Single
    Life As Single
    Active As Boolean
    Debris(2) As TRAIL_TYPE
    TimeSinceRespawn As Single
End Type

Private Type GUNSTYLE_TYPE
    Delay As Single
    CoolRate As Single
    HeatRate As Single
    Lifespan As Single
    Speed As Single
    StyleName As String * 20
    Radius As Single
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type



Private Sub Form_GotFocus()
RightKey = False
LeftKey = False
UpKey = False
Reverse = False
Ebrakes = False
Shoot = False
Brakes = False
End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If KeyCode = vbKeyRight Then RightKey = True
If KeyCode = vbKeyLeft Then LeftKey = True
If KeyCode = vbKeyUp Then UpKey = True
If KeyCode = vbKeyPageDown Then Reverse = True
If KeyCode = vbKeyEnd Then Ebrakes = True
If KeyCode = vbKeyControl Then Shoot = True
If KeyCode = vbKeyDown Then Brakes = True

If KeyCode = nmNumpadZero Then
    If Dampers Then
        Dampers = False
        frmBack.lblDampers = "Dampers: Off"
    Else:
        Dampers = True
        frmBack.lblDampers = "Dampers: On"
    End If
End If

If KeyCode = nmNumpadOne Then
    If QuadCannons Then
        QuadCannons = False
        frmBack.lblQuadCannons = "QuadCannons: Off"
    Else:
        QuadCannons = True
        frmBack.lblQuadCannons = "QuadCannons: On"
    End If
End If

If KeyCode = nmNumpadFour Then
    If QuadRotation Then
        QuadRotation = False
        frmBack.lblQuadRotation = "QuadRotation: Off"
    Else:
        QuadRotation = True
        frmBack.lblQuadRotation = "QuadRotation: On"
    End If
End If

If KeyCode = vbKeyShift Then
    GunMode = ((GunMode + 1) Mod (UBound(GunStyle) + ULTIMATEGUNACCESS))
    With GunStyle(GunMode)
        SHOTSPEED = .Speed
        BULLETRAD = .Radius
        GUNCOOLRATE = .CoolRate
        GUNHEATRATE = .HeatRate
        SHOTDELAY = .Delay
        BULLETLIFESPAN = .Lifespan
        frmBack.lblGunMode = "GunMode: " & .StyleName
    End With
End If


'Check to see if there are any BonusLevels. If so, let the player upgrade his ship.
If BonusLevels > 0 Then
    If KeyCode = vbKey1 Then
        MAXLIFE = MAXLIFE + 10
        CurLife = CurLife + 10
        BonusLevels = BonusLevels - 1
        frmBack.picLifeBack.Width = MAXLIFE
        frmBack.picLifeBack.Height = 20
        frmBack.picLifeBar.Width = MAXLIFE
        frmBack.picLifeBar.Height = 20
    End If
    
    If KeyCode = vbKey2 Then
        MAXHEAT = MAXHEAT + 5
        CurHeat = 0
        BonusLevels = BonusLevels - 1
        frmBack.picHeatBack.Width = MAXHEAT
        frmBack.picHeatBack.Height = 20
        frmBack.picHeatBar.Width = 0
        frmBack.picHeatBar.Height = 20
        HeatBarColor = vbYellow
    End If
    
    If KeyCode = vbKey3 Then
        SHIPREGENRATE = (SHIPREGENRATE + BASEREGENRATE * 0.5)
        BonusLevels = BonusLevels - 1
    End If
'Rof
    If KeyCode = vbKey4 Then
        If Not GunStyle(GunMode).Delay <= FDELAY Then
            GunStyle(GunMode).HeatRate = GunStyle(GunMode).HeatRate - 0.1
            GunStyle(GunMode).Delay = GunStyle(GunMode).Delay * 0.75
            BonusLevels = BonusLevels - 1
        End If
    End If
'Heat
    If KeyCode = vbKey5 Then
        If Not GunStyle(GunMode).HeatRate <= 1 Then
            GunStyle(GunMode).HeatRate = GunStyle(GunMode).HeatRate - 0.5
            BonusLevels = BonusLevels - 1
        End If
    End If
'range
    If KeyCode = vbKey6 Then
        GunStyle(GunMode).Lifespan = GunStyle(GunMode).Lifespan + 150
        BonusLevels = BonusLevels - 1
    End If
'size
    If KeyCode = vbKey7 Then
        If GunStyle(GunMode).Radius > 20 Then
            GunStyle(GunMode).Radius = GunStyle(GunMode).Radius + 1
            BonusLevels = BonusLevels - 1
        Else
            GunStyle(GunMode).Radius = GunStyle(GunMode).Radius + 0.5
            BonusLevels = BonusLevels - 1
        End If
    End If
'velocity
    If KeyCode = vbKey8 Then
        If Not GunStyle(GunMode).Speed >= 10 Then
            GunStyle(GunMode).Speed = GunStyle(GunMode).Speed + 2
            BonusLevels = BonusLevels - 1
        End If
    End If
'veloc Down
    If KeyCode = vbKey9 Then
        If Not GunStyle(GunMode).Speed <= 1 Then
            GunStyle(GunMode).Speed = GunStyle(GunMode).Speed - 2
            BonusLevels = BonusLevels - 1
        End If
    End If
'Overheat time
    If KeyCode = vbKey0 Then
        OVERHEATWAITTIME = OVERHEATWAITTIME - 500
        BonusLevels = BonusLevels - 1
    End If
    
'Reload the gun
    With GunStyle(GunMode)
        SHOTSPEED = .Speed
        BULLETRAD = .Radius
        GUNCOOLRATE = .CoolRate
        GUNHEATRATE = .HeatRate
        SHOTDELAY = .Delay
        BULLETLIFESPAN = .Lifespan
        frmBack.lblGunMode = "GunMode: " & .StyleName
    End With
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If KeyCode = vbKeyI Then
    For i = 0 To (frmBack.lblInstruct.Count - 1)
        frmBack.lblInstruct(i).Visible = True
    Next i
End If

If KeyCode = vbKeyEscape Then Pause True

'Programmers Shortcuts
If (KeyCode = vbKeyT) And (Shift = 7) Then GameTime = GameTime + 1000000
If (KeyCode = vbKeyP) And (Shift = 7) Then
    Score = Score + 5000
    CHEATED = True
End If
If (KeyCode = vbKeyU) And (Shift = 7) Then
    If Not ULTIMATEGUNACCESS = 1 Then ULTIMATEGUNACCESS = 1 'ReDim Preserve GunStyle(7)
    CHEATED = True
End If
End Sub

Public Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyRight Then RightKey = False
If KeyCode = vbKeyLeft Then LeftKey = False
If KeyCode = vbKeyUp Then UpKey = False
If KeyCode = vbKeyPageDown Then Reverse = False
If KeyCode = vbKeyEnd Then Ebrakes = False
If KeyCode = vbKeyControl Then Shoot = False
If KeyCode = vbKeyDown Then Brakes = False
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim NewBubbleTime As Single
Dim aaa As Long, aaa2 As Long, aaadif As Long



MAXHEAT = 100
MAXLIFE = 100
SHIPREGENRATE = BASEREGENRATE
OVERHEATWAITTIME = 5000

NEWBUBBLEDELAY = BASEBUBBLEREGENDELAY
QuadSpread = 0.5



' set the lifebars
frmBack.picLifeBack.Width = MAXLIFE
frmBack.picLifeBack.Height = 20
frmBack.picLifeBar.Width = MAXLIFE
frmBack.picLifeBar.Height = 20
CurLife = MAXLIFE

'Set the HeatBars
frmBack.picHeatBack.Width = MAXHEAT
frmBack.picHeatBack.Height = 20
frmBack.picHeatBar.Width = 0
frmBack.picHeatBar.Height = 20
HeatBarColor = vbYellow

MainDC = picBackBuffer.hdc
Running = True
InGameTimer = GetTickCount

'Initialyze the gunstyles
Call InitGunStyles

CurShotDel = 200 ' Makes sure the Ship can shoot right away
CurBullet = -1  ' Makes Sure that the first shot fired is #0 in the array
'Initialize the Bubble array

ShipCenterX = ScaleWidth / 2
ShipCenterY = ScaleHeight / 2

ReDim EnemyBullet(0)
ReDim Enemy(0)
ENEMYRADIUS = 10
ENEMYLIFE = 25
RespawnEnemies
ReDim Bubble(TOTALBUBBLE - 1)
ReDim Bullet(1)
For i = 0 To (TOTALBUBBLE \ 2)
    CreateNewBubble
Next
'Make Sure the Bubbles regenerate
BubbleRegen = True
frmBack.Show
frmMain.Show


'Start the game!
MainLoop
End Sub

Private Sub MainLoop()
On Error Resume Next
Dim HighScore As Double
Dim RecordHoldersName As String * 20
Dim DateSet As Variant
Dim i As Integer
Dim NewBubbleTime As Single
Dim TempFile As Integer
Dim FileData() As Byte
'TheFollowing line is for analyzing frame speed: enabled it and the texted marked below
'Open App.Path & "\aaa.txt" For Output As #1
Do While Running
    CurDelay = GetTickCount - InGameTimer
    
If GetTickCount - InGameTimer >= FDELAY Then
   
    If Not Paused Then
        NewBubbleTime = NewBubbleTime + CurDelay
        If (NewBubbleTime > NEWBUBBLEDELAY) And (BubbleRegen) Then
            CreateNewBubble
            NewBubbleTime = 0
        End If
        InGameTimer = GetTickCount
        ShipPhysics
        BubblePhysics
        EnemyPhysics
        RefreshBars
        BitBlt picBackBuffer.hdc, 0, 0, picBackBuffer.ScaleWidth, picBackBuffer.ScaleHeight, 0, 0, 0, vbBlackness
        DrawEnemies
        DrawBubbles
        DrawShip
        DrawBullets
        CollisionDetection
        
        'enable the following 3 lines and the "Open" line above to examine fram speed
        
        'aaa2 = GetTickCount
        'aaadif = aaa2 - InGameTimer
        'Write #1, aaadif, BulletsToShow.Count, CurFrame
        
        'The first number (in the text file) will be the time (in mils) it took that frame to execute,
        'the second number is how many bullets are being drawn nd collision-tested,
        'the third number is the current frame
        
        lblCHEATED.Visible = False
    Else
        If CHEATED Then
            lblCHEATED.Visible = True
        Else
            lblCHEATED.Visible = False
        End If
        TempFile = FreeFile
        'Update the pause-screen high score
        Open App.Path & "/HighScores.rec" For Binary Access Read Write As #HIGHSCOREFILE
        Open App.Path & "/A31.rec" For Binary Access Read Write As #ENCRYPTFILE
        If Not FileLen(App.Path & "/HighScores.rec") = 0 Then
            
    
            'Get the encrypted numbers
            ReDim FileData(RECORDSIZE)
            For i = 1 To RECORDSIZE
                Get HIGHSCOREFILE, i, FileData(i)
            Next i
            
            'Put the decrypted data into the temporary file
            For i = 1 To RECORDSIZE
                Put ENCRYPTFILE, i, FileData(i) Xor ENCRYPTKEY
            Next i
            
            'Get the decrypted data
            Get ENCRYPTFILE, 1, HighScore
            Get ENCRYPTFILE, , RecordHoldersName
            Get ENCRYPTFILE, , DateSet
            
            
        End If
        Close
        lblHighScore = HighScore
        lblRecordHolder = RecordHoldersName
        lblDateSet = DateSet
        
        'The following was riped right from RefreshBars. I do NOT recommend replacing it with a call to that subroutine.
        'Update the Score Label. Use an If statement so it does not flicker from constantly being set
        If Not frmBack.lblScore = Score Then frmBack.lblScore = Score
        'Check to see if there are any bonus levels. If there are, show the level up labels
        If BonusLevels > 0 Then
            'Show how many bonus levels there are
            If Not frmBack.lblLevelUp(13) = BonusLevels Then frmBack.lblLevelUp(13) = BonusLevels
            For i = 0 To frmBack.lblLevelUp.Count - 1
                'Make sure it does not flicker...
                If Not frmBack.lblLevelUp(i).Visible = True Then frmBack.lblLevelUp(i).Visible = True
            Next i
        Else
            'Hide them
            For i = 0 To frmBack.lblLevelUp.Count - 1
                If Not frmBack.lblLevelUp(i).Visible = False Then frmBack.lblLevelUp(i).Visible = False
            Next i
        End If
        
        DrawShip
        DrawEnemies
        DrawBubbles
        DrawBullets
    End If
End If

    DoEvents
Loop
End Sub

Private Sub Pause(Resumable As Boolean)
If Not Paused Then
    Paused = True
    Cls
    
    
    lblDateSet.Visible = True
    lblExit.Visible = True
    lblHighScore.Visible = True
    lblHS.Visible = True
    lblLoad.Visible = True
    lblNewGame.Visible = True
    lblPaused.Visible = True
    lblRecordHolder.Visible = True
    lblResume.Visible = True
    
    If Not Resumable Then
        lblResume.Enabled = False
    Else
        lblResume.Enabled = True
    End If
    
    lblSave.Visible = True
    

    
Else
    If Resumable Then
        Paused = False
        
        lblHighScore.Visible = False
        lblHS.Visible = False
        lblRecordHolder.Visible = False
        lblDateSet.Visible = False
        lblExit.Visible = False
        lblLoad.Visible = False
        lblNewGame.Visible = False
        lblPaused.Visible = False
        lblResume.Visible = False
        lblSave.Visible = False
        InGameTimer = GetTickCount
        CurDelay = 0
        '
    End If
End If
End Sub


Private Sub Form_LostFocus()
frmMain.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Running = False
    Paused = False
    Close
    End
End Sub

Private Sub BubblePhysics()
Dim i As Integer
Dim a As Integer
For i = 0 To UBound(Bubble)
    If Bubble(i).Active Then
        Bubble(i).x = Bubble(i).x + Bubble(i).Speed * Sin(Bubble(i).Heading)
        Bubble(i).y = Bubble(i).y - Bubble(i).Speed * Cos(Bubble(i).Heading)
        
        If Bubble(i).x > ScaleWidth Then 'Bubble(i).X = 0
            Bubble(i).x = Bubble(i).x Mod ScaleWidth
        End If
        If Bubble(i).x < 0 Then 'Bubble(i).X = ScaleWidth
            Bubble(i).x = (ScaleWidth + Bubble(i).x - 1) Mod ScaleWidth
        End If
        If Bubble(i).y > ScaleHeight Then 'Bubble(i).Y = 0
            Bubble(i).y = Bubble(i).y Mod ScaleHeight
        End If
        If Bubble(i).y < 0 Then 'Bubble(i).Y = ScaleHeight
             Bubble(i).y = (ScaleHeight + Bubble(i).y - 1) Mod ScaleHeight
        End If
    Else
        For a = 0 To UBound(Bubble(i).Debris)
            Bubble(i).Debris(a).x = Bubble(i).Debris(a).x + Bubble(i).Debris(a).Speed * Sin(Bubble(i).Debris(a).Heading)
            Bubble(i).Debris(a).y = Bubble(i).Debris(a).y + Bubble(i).Debris(a).Speed * Cos(Bubble(i).Debris(a).Heading)
        Next a
    End If
Next i
End Sub

Private Sub EnemyPhysics()
Dim XComp As Single
Dim YComp As Single
Dim AngleToShip As Single
Dim DistToShip As Single
Dim a As Integer
Dim i As Integer


For i = 0 To UBound(Enemy)
    If Enemy(i).Active Then
        With Enemy(i)
            .CurShotDelay = .CurShotDelay + CurDelay
            AngleToShip = (GetAngle(.x, .y, ShipCenterX, ShipCenterY))
            DistToShip = GetDist(.x, .y, ShipCenterX, ShipCenterY)
            If DistToShip > ENEMYDESIREDDIST Then
                
                .Facing = AngleToShip
                
                'Too far away. Accelerate
                XComp = .Speed * Sin(.Heading) + (ENEMYACCRATE) * Sin(.Facing)
                YComp = .Speed * Cos(.Heading) + (ENEMYACCRATE) * Cos(.Facing)
    
                .Speed = Sqr(XComp * XComp + YComp * YComp)
                If YComp > 0 Then .Heading = Atn(XComp / YComp)
                If YComp < 0 Then .Heading = Atn(XComp / YComp) + Pi
                
            Else
                'Circle the player
                If (i Mod 2) = 0 Then .Facing = AngleToShip + Pi2
                If (i Mod 2) = 1 Then .Facing = AngleToShip - Pi2
                .Speed = (ENEMYDESIREDSPEED + (i \ 2))
                
            End If
            .Heading = .Facing
            .x = .x + .Speed * Sin(.Heading)
            .y = .y - .Speed * Cos(.Heading)
            .TimeSinceRespawn = .TimeSinceRespawn + CurDelay
        End With
            
        If Enemy(i).CurShotDelay >= EnemyGun.Delay Then
            Enemy(i).CurShotDelay = 0
            EnemyCurBullet = 0
            For a = 1 To UBound(EnemyBullet)
                If Not EnemyBullet(a).Active Then
                    EnemyCurBullet = a
                    Exit For
                End If
            Next a
            If EnemyCurBullet = 0 Then
                EnemyCurBullet = UBound(EnemyBullet) + 1
                ReDim Preserve EnemyBullet(EnemyCurBullet)
            End If
            EnemyBullet(EnemyCurBullet).Active = True
            EnemyBullet(EnemyCurBullet).Heading = AngleToShip
            EnemyBullet(EnemyCurBullet).HitsTaken = 0
            EnemyBullet(EnemyCurBullet).Lifespan = EnemyGun.Lifespan
            EnemyBullet(EnemyCurBullet).Radius = EnemyGun.Radius
            EnemyBullet(EnemyCurBullet).Speed = EnemyGun.Speed + Enemy(i).Speed
            EnemyBullet(EnemyCurBullet).TimeInFlight = 0
            EnemyBullet(EnemyCurBullet).x = Enemy(i).x
            EnemyBullet(EnemyCurBullet).y = Enemy(i).y
        End If
        'Wrap the Enemy
        If Enemy(i).x > ScaleWidth Then 'Enemy(i).X = 0
            Enemy(i).x = ScaleWidth Mod ScaleWidth
        End If
        If Enemy(i).x < 0 Then 'Enemy(i).X = ScaleWidth
            Enemy(i).x = (ScaleWidth + Enemy(i).x - 1) Mod ScaleWidth
        End If
        If Enemy(i).y > ScaleHeight Then 'Enemy(i).Y = 0
            Enemy(i).y = Enemy(i).y Mod ScaleHeight
        End If
        If Enemy(i).y < 0 Then 'Enemy(i).Y = ScaleHeight
            Enemy(i).y = (ScaleHeight + Enemy(i).y - 1) Mod ScaleHeight
        End If
    Else
        For a = 0 To UBound(Enemy(i).Debris)
            Enemy(i).Debris(a).x = Enemy(i).Debris(a).x + Enemy(i).Debris(a).Speed * Sin(Enemy(i).Debris(a).Heading)
            Enemy(i).Debris(a).y = Enemy(i).Debris(a).y + Enemy(i).Debris(a).Speed * Cos(Enemy(i).Debris(a).Heading)
        Next a
    End If
    
    BitBlt frmMain.hdc, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight, picBackBuffer.hdc, 0, 0, vbSrcCopy
Next i


'Update the bullets
For i = 0 To UBound(EnemyBullet)
    If EnemyBullet(i).Active Then
        EnemyBullet(i).TimeInFlight = EnemyBullet(i).TimeInFlight + CurDelay
        If EnemyBullet(i).TimeInFlight > EnemyBullet(i).Lifespan Then
            EnemyBullet(i).Active = False
        End If
        EnemyBullet(i).x = EnemyBullet(i).x + EnemyBullet(i).Speed * Sin(EnemyBullet(i).Heading)
        EnemyBullet(i).y = EnemyBullet(i).y - EnemyBullet(i).Speed * Cos(EnemyBullet(i).Heading)
    End If
Next i
End Sub

Private Sub ShipPhysics()
Dim XComp As Single
Dim YComp As Single

Dim i As Integer
Dim a As Integer




If RightKey Then
    ShipFacing = ShipFacing + (ROT_RATE) * Pi180
    If ShipFacing > PiBy2 Then ShipFacing = ShipFacing - PiBy2
End If

If LeftKey Then
    ShipFacing = ShipFacing - (ROT_RATE) * Pi180
    If ShipFacing < 0 Then ShipFacing = ShipFacing + PiBy2
End If

If UpKey Then
    
    XComp = ShipSpeed * Sin(ShipHeading) + (ACCRATE) * Sin(ShipFacing)
    YComp = ShipSpeed * Cos(ShipHeading) + (ACCRATE) * Cos(ShipFacing)
    
    ShipSpeed = Sqr(XComp * XComp + YComp * YComp)
    If YComp > 0 Then ShipHeading = Atn(XComp / YComp)
    If YComp < 0 Then ShipHeading = Atn(XComp / YComp) + Pi
    
    If ShipHeading < 0 Then ShipHeading = ShipHeading + PiBy2
End If

If Ebrakes Then
    If ShipSpeed - EBRAKERATE >= 0 Then
        ShipSpeed = ShipSpeed - EBRAKERATE
    ElseIf ShipSpeed + EBRAKERATE <= 0 Then
        ShipSpeed = ShipSpeed + EBRAKERATE
    Else: ShipSpeed = 0
    End If
End If

If Reverse Then
    ShipSpeed = ShipSpeed - ACCRATE
End If

If Brakes Then
    If Dampers Then
        If ShipSpeed - BRAKERATE >= 0 Then
            ShipSpeed = ShipSpeed - BRAKERATE
        ElseIf ShipSpeed + BRAKERATE <= 0 Then
            ShipSpeed = ShipSpeed + BRAKERATE
        Else: ShipSpeed = 0
        End If
    Else
        XComp = ShipSpeed * Sin(ShipHeading) - (ACCRATE) * Sin(ShipFacing)
        YComp = ShipSpeed * Cos(ShipHeading) - (ACCRATE) * Cos(ShipFacing)
        
        ShipSpeed = Sqr(XComp * XComp + YComp * YComp)
        If YComp > 0 Then ShipHeading = Atn(XComp / YComp)
        If YComp < 0 Then ShipHeading = Atn(XComp / YComp) + Pi
        
        If ShipHeading < 0 Then ShipHeading = ShipHeading + PiBy2
    End If
End If

If Dampers Then     'The Dampers make the ship attempt to move in the direction it is facing
     ShipHeading = ShipFacing
End If


'Deal with the trail
CurTrail = (CurTrail + 1) Mod (UBound(Trail) + 1)
Trail(CurTrail).x = (ShipCenterX - SHIPLNG * Sin(ShipFacing))
Trail(CurTrail).y = (ShipCenterY + SHIPLNG * Cos(ShipFacing))
If Not ShipSpeed = 0 Then
    Trail(CurTrail).Heading = (ShipHeading + Pi - 0.5 + (Num(0, 1000) * 0.001))
Else
    Trail(CurTrail).Heading = (ShipFacing + Pi - 0.5 + (Num(0, 1000) * 0.001))
End If
Trail(CurTrail).Speed = (Num(50, 100) * 0.01)

For i = 0 To UBound(Trail)
    Trail(i).x = Trail(i).x + Trail(i).Speed * Sin(Trail(i).Heading)
    Trail(i).y = Trail(i).y - Trail(i).Speed * Cos(Trail(i).Heading)
Next i

'Move the Ship
ShipCenterX = ShipCenterX + ShipSpeed * Sin(ShipHeading)
ShipCenterY = ShipCenterY - ShipSpeed * Cos(ShipHeading)

CurShotDel = CurShotDel + CurDelay
If Shoot Then
    
    ' make sure the appropriate time has passed
    If (CurShotDel >= SHOTDELAY) And (CurHeat < MAXHEAT) Then
        'Heat the guns
        
        CurHeat = (CurHeat + (GUNHEATRATE))
        
        
        'Reset the shot delay
        CurShotDel = 0
        'Shoot
        CurCannon = ((CurCannon + 1) Mod 4)
        CurBullet = FreeBullet()
        
        
        If CurBullet = 0 Then
            CurBullet = UBound(Bullet) + 1
            ReDim Preserve Bullet(CurBullet)
        End If
        If QuadCannons Then
            With Bullet(CurBullet)
                If CurCannon = 1 Then
                    .x = ShipX2
                    .y = ShipY2
                ElseIf CurCannon = 2 Then
                    .x = ShipX1
                    .y = ShipY1
                ElseIf CurCannon = 3 Then
                    .x = ShipX3
                    .y = ShipY3
                ElseIf CurCannon = 0 Then
                    .x = ShipX4
                    .y = ShipY4
                End If
                .BulletName = GunStyle(GunMode).StyleName
                .Radius = BULLETRAD
                .HitsTaken = 0
                .Active = True
                .TimeInFlight = 0
                .Lifespan = BULLETLIFESPAN
                If (CurCannon = 1) Or (CurCannon = 2) Then
                    XComp = ShipSpeed * Sin(ShipHeading) + SHOTSPEED * Sin(ShipFacing)
                    YComp = ShipSpeed * Cos(ShipHeading) + SHOTSPEED * Cos(ShipFacing)
                ElseIf (CurCannon = 3) Then
                    XComp = ShipSpeed * Sin(ShipHeading) + SHOTSPEED * Sin(ShipFacing + QuadSpread)
                    YComp = ShipSpeed * Cos(ShipHeading) + SHOTSPEED * Cos(ShipFacing + QuadSpread)
                ElseIf (CurCannon = 0) Then
                    XComp = ShipSpeed * Sin(ShipHeading) + SHOTSPEED * Sin(ShipFacing - QuadSpread)
                    YComp = ShipSpeed * Cos(ShipHeading) + SHOTSPEED * Cos(ShipFacing - QuadSpread)
                    'Rotate the Cannons if necessary
                    If QuadRotation Then QuadSpread = QuadSpread + QUADROTATIONSPEED
                End If
                .Speed = Sqr(XComp * XComp + YComp * YComp)
                If YComp > 0 Then .Heading = Atn(XComp / YComp)
                If YComp < 0 Then .Heading = Atn(XComp / YComp) + Pi
            End With
            
        Else
             With Bullet(CurBullet)
                If CurCannon = 1 Then
                    .x = ShipX2
                    .y = ShipY2
                ElseIf CurCannon = 2 Then
                    .x = ShipX1
                    .y = ShipY1
                ElseIf CurCannon = 3 Then
                    .x = ShipX2
                    .y = ShipY2
                ElseIf CurCannon = 0 Then
                    .x = ShipX1
                    .y = ShipY1
                End If
                .BulletName = GunStyle(GunMode).StyleName
                .Radius = BULLETRAD
                .HitsTaken = 0
                .Active = True
                .TimeInFlight = 0
                .Lifespan = BULLETLIFESPAN
                XComp = ShipSpeed * Sin(ShipHeading) + SHOTSPEED * Sin(ShipFacing)
                YComp = ShipSpeed * Cos(ShipHeading) + SHOTSPEED * Cos(ShipFacing)
                .Speed = Sqr(XComp * XComp + YComp * YComp)
                If YComp > 0 Then .Heading = Atn(XComp / YComp)
                If YComp < 0 Then .Heading = Atn(XComp / YComp) + Pi
                
                
            End With
        End If
    End If
'If the player is not trying to shoot and is not overheated, let him regenerate
ElseIf Not OverHeated Then
    If Not CurLife >= MAXLIFE Then CurLife = CurLife + SHIPREGENRATE
    'Make sure we did not go over the ships maximum life rating
    If CurLife > MAXLIFE Then CurLife = MAXLIFE
End If

'Trim the Bullet Array, if possible
If Not Bullet(UBound(Bullet)).Active Then
    For i = UBound(Bullet) To 1 Step -1
        If Not Bullet(i).Active Then
            CurBullet = i
        Else
            Exit For
        End If
    Next i
    ReDim Preserve Bullet(i)
End If

For i = 1 To UBound(Bullet)
    If Bullet(i).Active Then
        
        'For the heat seeking gun
        If Bullet(i).BulletName = GunStyle(6).StyleName Then
            For a = 0 To UBound(Enemy)
                If Enemy(a).Active Then
                    Bullet(i).Heading = GetAngle(Bullet(i).x, Bullet(i).y, Enemy(a).x, Enemy(a).y)
                    Exit For
                End If
            Next a
        End If
        
        Bullet(i).TimeInFlight = Bullet(i).TimeInFlight + FDELAY
        If Bullet(i).TimeInFlight >= Bullet(i).Lifespan Then Bullet(i).Active = False
        Bullet(i).x = Bullet(i).x + Bullet(i).Speed * Sin(Bullet(i).Heading)
        Bullet(i).y = Bullet(i).y - Bullet(i).Speed * Cos(Bullet(i).Heading)
        
    End If
Next i

'Wrap the Ship
If ShipCenterX > ScaleWidth Then 'ShipCenterX = 0
    ShipCenterX = ScaleWidth Mod ScaleWidth
End If
If ShipCenterX < 0 Then 'ShipCenterX = ScaleWidth
    ShipCenterX = (ScaleWidth + ShipCenterX - 1) Mod ScaleWidth
End If
If ShipCenterY > ScaleHeight Then 'ShipCenterY = 0
    ShipCenterY = ShipCenterY Mod ScaleHeight
End If
If ShipCenterY < 0 Then 'ShipCenterY = ScaleHeight
    ShipCenterY = (ScaleHeight + ShipCenterY - 1) Mod ScaleHeight
End If



End Sub

Private Sub DrawShip()
Dim i As Integer
Dim TempPoint As POINT_TYPE
'    ShipX1 = ShipCenterX + Ship_RAD * Sin(ShipFacing)
'    ShipY1 = ShipCenterY - Ship_RAD * Cos(ShipFacing)
'    ShipX2 = ShipCenterX + Ship_RAD * Sin(ShipFacing + 1.5 * Pi3)
'    ShipY2 = ShipCenterY - Ship_RAD * Cos(ShipFacing + 1.5 * Pi3)
'    ShipX3 = ShipCenterX + Ship_RAD * Sin(ShipFacing + 4.5 * Pi3)
'    ShipY3 = ShipCenterY - Ship_RAD * Cos(ShipFacing + 4.5 * Pi3)
    
'    Line (ShipX1, ShipY1)-(ShipX2, ShipY2), vbRed
'    Line (ShipX2, ShipY2)-(ShipX3, ShipY3), vbWhite
'    Line (ShipX3, ShipY3)-(ShipX1, ShipY1), vbRed

    ShipX1 = ShipCenterX + (SHIPLNG * Sin(ShipFacing)) + (SHIPFWDTH * Sin(ShipFacing + Pi2))
    ShipY1 = ShipCenterY - (SHIPLNG * Cos(ShipFacing)) - (SHIPFWDTH * Cos(ShipFacing + Pi2))
    ShipX2 = ShipCenterX + (SHIPLNG * Sin(ShipFacing)) + (SHIPFWDTH * Sin(ShipFacing - Pi2))
    ShipY2 = ShipCenterY - (SHIPLNG * Cos(ShipFacing)) - (SHIPFWDTH * Cos(ShipFacing - Pi2))
    ShipX3 = ShipCenterX - (SHIPLNG * Sin(ShipFacing)) + (SHIPRWDTH * Sin(ShipFacing + Pi2))
    ShipY3 = ShipCenterY + (SHIPLNG * Cos(ShipFacing)) - (SHIPRWDTH * Cos(ShipFacing + Pi2))
    ShipX4 = ShipCenterX - (SHIPLNG * Sin(ShipFacing)) + (SHIPRWDTH * Sin(ShipFacing - Pi2))
    ShipY4 = ShipCenterY + (SHIPLNG * Cos(ShipFacing)) - (SHIPRWDTH * Cos(ShipFacing - Pi2))

    ' Draw The Ship
    picBackBuffer.ForeColor = vbRed
    MoveToEx MainDC, ShipX2, ShipY2, TempPoint 'Left
    LineTo MainDC, ShipX4, ShipY4
    MoveToEx MainDC, ShipX3, ShipY3, TempPoint 'Right
    LineTo MainDC, ShipX1, ShipY1
    MoveToEx MainDC, ShipX1, ShipY1, TempPoint 'Front
    LineTo MainDC, ShipX2, ShipY2
    picBackBuffer.ForeColor = vbWhite
    MoveToEx MainDC, ShipX4, ShipY4, TempPoint 'Rear
    LineTo MainDC, ShipX3, ShipY3
   
For i = 0 To UBound(Trail)
    SetPixel MainDC, Trail(i).x, Trail(i).y, vbRed
Next i
End Sub

Private Sub DrawBullets()
Dim i As Integer
Dim Temp As Integer     'where we'll put the current collection item for ease of use
Dim RectX1 As Single, RectY1 As Single  'Where we'll store the circle boundaries for Ellipse
Dim RectX2 As Single, RectY2 As Single
Dim TempPoint As POINT_TYPE

picBackBuffer.ForeColor = &H5555FF

For i = 1 To UBound(Bullet)
    If Bullet(i).Active Then
    
        'Wrap the Bullet
        If Bullet(i).x > ScaleWidth Then 'Bullet(i).X = 0
            Bullet(i).x = ScaleWidth Mod ScaleWidth
        End If
        If Bullet(i).x < 0 Then 'Bullet(i).X = ScaleWidth
            Bullet(i).x = (ScaleWidth + Bullet(i).x - 1) Mod ScaleWidth
        End If
        If Bullet(i).y > ScaleHeight Then 'Bullet(i).Y = 0
            Bullet(i).y = Bullet(i).y Mod ScaleHeight
        End If
        If Bullet(i).y < 0 Then 'Bullet(i).Y = ScaleHeight
            Bullet(i).y = (ScaleHeight + Bullet(i).y - 1) Mod ScaleHeight
        End If
        
        
        RectX1 = Bullet(i).x - Bullet(i).Radius
        RectY1 = Bullet(i).y + Bullet(i).Radius
        RectX2 = Bullet(i).x + Bullet(i).Radius
        RectY2 = Bullet(i).y - Bullet(i).Radius
        
        MoveToEx MainDC, RectX1, RectY1, TempPoint
        Ellipse MainDC, RectX1, RectY1, RectX2, RectY2
        
    End If
Next i

picBackBuffer.ForeColor = &HFFAA00
For i = 1 To UBound(EnemyBullet)
    If EnemyBullet(i).Active Then
    
        'Wrap the EnemyBullet
        If EnemyBullet(i).x > ScaleWidth Then 'EnemyBullet(i).X = 0
            EnemyBullet(i).x = ScaleWidth Mod ScaleWidth
        End If
        If EnemyBullet(i).x < 0 Then 'EnemyBullet(i).X = ScaleWidth
            EnemyBullet(i).x = (ScaleWidth + EnemyBullet(i).x - 1) Mod ScaleWidth
        End If
        If EnemyBullet(i).y > ScaleHeight Then 'EnemyBullet(i).Y = 0
            EnemyBullet(i).y = EnemyBullet(i).y Mod ScaleHeight
        End If
        If EnemyBullet(i).y < 0 Then 'EnemyBullet(i).Y = ScaleHeight
            EnemyBullet(i).y = (ScaleHeight + EnemyBullet(i).y - 1) Mod ScaleHeight
        End If
        
        
        RectX1 = EnemyBullet(i).x - EnemyBullet(i).Radius
        RectY1 = EnemyBullet(i).y + EnemyBullet(i).Radius
        RectX2 = EnemyBullet(i).x + EnemyBullet(i).Radius
        RectY2 = EnemyBullet(i).y - EnemyBullet(i).Radius
        
        MoveToEx MainDC, RectX1, RectY1, TempPoint
        Ellipse MainDC, RectX1, RectY1, RectX2, RectY2
        
    End If
Next i
picBackBuffer.ForeColor = vbWhite

End Sub

Private Sub DrawBubbles()
Dim i As Integer
Dim a As Integer
Dim Temp As Integer     'where we'll put the current collection item for ease of use
Dim RectX1 As Single, RectY1 As Single  'Where we'll store the circle boundaries for Ellipse
Dim RectX2 As Single, RectY2 As Single
Dim TempPoint As POINT_TYPE

picBackBuffer.ForeColor = vbWhite
For i = 1 To UBound(Bubble)
    If Bubble(i).Active Then
        RectX1 = Bubble(i).x - Bubble(i).Radius
        RectY1 = Bubble(i).y + Bubble(i).Radius
        RectX2 = Bubble(i).x + Bubble(i).Radius
        RectY2 = Bubble(i).y - Bubble(i).Radius
    
        MoveToEx MainDC, RectX1, RectY1, TempPoint
        Ellipse MainDC, RectX1, RectY1, RectX2, RectY2
    Else
        
        For a = 0 To UBound(Bubble(i).Debris)
            SetPixel MainDC, Bubble(i).Debris(a).x, Bubble(i).Debris(a).y, vbWhite
        Next a
    End If
Next i

End Sub

Private Sub CollisionDetection()
Dim i As Integer
Dim a As Integer
Dim j As Integer
Dim ShipWidthAverage As Single
Dim TempPoint As POINT_TYPE
'Bubbles
For a = 0 To UBound(Bubble)
    'If this Bubble is not active, then don't do anything!
    If (Bubble(a).Active) Then
        'Check for collision with the ship
        If (ColWithShip(Bubble(a).x, Bubble(a).y, Bubble(a).Radius)) Then
            If Not (Bubble(a).Collision) Then
                Bubble(a).Collision = True    'We check for Bubble(a).Collision so if the ship hits the Bubble and stays in it more than one frame it the ship is not utterly destroyed pretty much instantly
                CurLife = CurLife - Bubble(a).Radius
                Bubble(a).Radius = (Bubble(a).Radius * 0.5)
                If Bubble(a).Radius < MINBUBBLERAD Then
                    Bubble(a).Active = False
                    For j = 0 To UBound(Bubble(a).Debris)
                        Bubble(a).Debris(j).Heading = (Num(0, 628) * 0.01)
                        Bubble(a).Debris(j).Speed = Num(10, 500) * 0.01
                        Bubble(a).Debris(j).x = Bubble(a).x
                        Bubble(a).Debris(j).y = Bubble(a).y
                    Next j
                End If
            End If
        Else: Bubble(a).Collision = False
        End If

        'Bullets
        For i = 1 To UBound(Bullet)
            If (Bullet(i).Active) Then
                If CircCol(Bullet(i).x, Bullet(i).y, Bullet(i).Radius, Bubble(a).x, Bubble(a).y, Bubble(a).Radius) Then
                    Score = Score + 15
                    Bubble(a).Radius = Bubble(a).Radius - 1
                    If Bubble(a).Radius <= MINBUBBLERAD Then
                        Bubble(a).Active = False
                        Score = Score + 300
                        For j = 0 To UBound(Bubble(a).Debris)
                            Bubble(a).Debris(j).Heading = (Num(0, 628) * 0.01)
                            Bubble(a).Debris(j).Speed = Num(10, 500) * 0.01
                            Bubble(a).Debris(j).x = Bubble(a).x
                            Bubble(a).Debris(j).y = Bubble(a).y
                        Next j
                    End If
                    Bullet(i).HitsTaken = Bullet(i).HitsTaken + 1
                    If (Bullet(i).HitsTaken >= Bullet(i).Radius) Then Bullet(i).Active = False
                End If
            End If
        Next i
    End If
Next a

For a = 0 To UBound(Enemy)
    If Enemy(a).Active Then
        For i = 1 To UBound(Bullet)
            If (Bullet(i).Active) Then
                If CircCol(Bullet(i).x, Bullet(i).y, Bullet(i).Radius, Enemy(a).x, Enemy(a).y, ENEMYRADIUS) Then
                    Score = Score + 20
                    Bullet(i).HitsTaken = Bullet(i).HitsTaken + 1
                    Enemy(a).Life = Enemy(a).Life - 1
                    If Enemy(a).Life <= 0 Then
                        Enemy(a).Active = False
                        Score = Score + 400
                        'Give a bonus for killing it fast
                        If Not (Enemy(a).TimeSinceRespawn > 30000) Then Score = Score + ((30000 - Enemy(a).TimeSinceRespawn) \ 5000)
                        For j = 0 To UBound(Enemy(a).Debris)
                            Enemy(a).Debris(j).Heading = (Num(0, 628) * 0.01)
                            Enemy(a).Debris(j).Speed = Num(10, 500) * 0.01
                            Enemy(a).Debris(j).x = Enemy(a).x
                            Enemy(a).Debris(j).y = Enemy(a).y
                        Next j
                    End If
                    Bullet(i).HitsTaken = Bullet(i).HitsTaken + 1
                    If Bullet(i).HitsTaken >= Bullet(i).Radius Then Bullet(i).Active = False
                End If
            End If
        Next i
    End If
Next a

For i = 1 To UBound(EnemyBullet)
    If EnemyBullet(i).Active Then
        For a = 1 To UBound(Bullet)
            If CircCol(Bullet(a).x, Bullet(a).y, Bullet(a).Radius, EnemyBullet(i).x, EnemyBullet(i).y, EnemyBullet(i).Radius) Then
                Bullet(a).HitsTaken = Bullet(a).HitsTaken + EnemyBullet(i).Radius
                EnemyBullet(i).Active = False
                If Bullet(a).HitsTaken > Bullet(a).Radius Then Bullet(a).Active = False
            End If
        Next a
        If EnemyBullet(i).Active Then
            If ColWithShip(EnemyBullet(i).x, EnemyBullet(i).y, EnemyBullet(i).Radius) Then
                CurLife = CurLife - EnemyBullet(i).Radius
                EnemyBullet(i).Active = False
            End If
        End If
    End If
Next i


End Sub




Private Sub RefreshBars()
Dim i As Single
Dim EnemiesLeft As Boolean
Dim BubbleLeft As Boolean
Dim OldLevel As Single
Dim TimeHours As String
Dim TimeMinutes As String
Dim TimeSeconds As String


'LifeBars first
If CurLife > 0 Then
    frmBack.picLifeBar.Width = CurLife
Else
    frmBack.picLifeBar.Width = 0
End If

'Now the heatbars

'Since picturebox width must be higher than 0...
If CurHeat <= 0 Then CurHeat = 1

'To prevent the bar from wacking out when overheated...
If Not OverHeated Then
    'To make the bar rise more smoothly...
    If (CurHeat > frmBack.picHeatBar.Width + 2) Then
        frmBack.picHeatBar.Width = frmBack.picHeatBar.Width + 3
    ElseIf (CurHeat < frmBack.picHeatBar.Width) And (frmBack.picHeatBar.Width > 3) Then
        frmBack.picHeatBar.Width = frmBack.picHeatBar.Width - 1
    ElseIf CurHeat > 0 Then
         frmBack.picHeatBar.Width = CurHeat
    Else
        frmBack.picHeatBar.Width = 1
    End If
'And to make it rise all the way when overheated... (Without this it will simply stay where it is)
Else
    frmBack.picHeatBar.Width = MAXHEAT
End If
'Check to see if the guns just over heated this frame
If (Not OverHeated) And (CurHeat >= MAXHEAT) Then
    OverHeated = True
    OverHeatTime = 0 - CurDelay     'Set the overheattime to below zero because it will be set back to zero in the next if statement
End If

'if the guns are overheated, deal with them
If OverHeated Then
    OverHeatTime = OverHeatTime + CurDelay
    'Find the heatbarcolor
    ColorChangedTime = ColorChangedTime + CurDelay
    If ColorChangedTime > 300 Then      'Without this if, the rapid colorchanging would surely cause a siezure
        ColorChangedTime = 0
        If HeatBarColor = vbYellow Then
            HeatBarColor = vbRed
        Else: HeatBarColor = vbYellow
        End If
    End If
    'Set the heatbarcolor
    frmBack.picHeatBar.BackColor = HeatBarColor
    If OverHeatTime > OVERHEATWAITTIME Then
        OverHeated = False
        CurHeat = MAXHEAT * 0.6
        HeatBarColor = vbYellow
        frmBack.picHeatBar.BackColor = HeatBarColor
    End If
Else
    'Cool the guns
    If Not CurHeat > MAXHEAT Then CurHeat = CurHeat - GUNCOOLRATE   ' If overheating, the gun cools slower. See the RefreshBars Subroutine
End If


'Update the level and all related variables
OldLevel = Level
If (Score < 50000) Then
    Level = (Score \ 10000)  'A Forward Slash (\) = integer division. We don't want any decimals!
ElseIf (Score < 110000) Then
    Level = ((Score - 50000) \ 15000) + 5
Else
    Level = ((Score - 110000) \ 20000) + 9
End If
'If the player has cheated, don't give him nuthin'!
If Not CHEATED Then
    BonusLevels = BonusLevels + (Level - OldLevel)
End If
If Not frmBack.lblLevel = Level Then
    For i = 0 To UBound(Enemy)
        If Enemy(i).Active Then
            EnemiesLeft = True
            Exit For
        End If
    Next i
    
    If Not EnemiesLeft Then
        BubbleRegen = True
        frmBack.lblLevel = Level
        RespawnEnemies
        For i = 0 To (TOTALBUBBLE - 1)
            CreateNewBubble
        Next i
        
        NEWBUBBLEDELAY = BASEBUBBLEREGENDELAY - (Level * 100)

    Else
        BubbleRegen = False
    End If
End If
'Update the Score Label. Use an If statement so it does not flicker from constantly being set
If Not frmBack.lblScore = Score Then frmBack.lblScore = Score
'Check to see if there are any bonus levels. If there are, show the level up labels
If BonusLevels > 0 Then
    'Show how many bonus levels there are
    If Not frmBack.lblLevelUp(13) = BonusLevels Then frmBack.lblLevelUp(13) = BonusLevels
    For i = 0 To frmBack.lblLevelUp.Count - 1
        'Make sure it does not flicker...
        If Not frmBack.lblLevelUp(i).Visible = True Then frmBack.lblLevelUp(i).Visible = True
    Next i
Else
    'Hide them
    For i = 0 To frmBack.lblLevelUp.Count - 1
        If Not frmBack.lblLevelUp(i).Visible = False Then frmBack.lblLevelUp(i).Visible = False
    Next i
End If

'Check to see if there are any Bubbles left. if there are not, make some
If BubbleRegen Then
    For i = 0 To UBound(Bubble)
        If Bubble(i).Active Then
            BubbleLeft = True
            Exit For
        End If
    Next i
    
    If Not BubbleLeft Then
        For i = 0 To (TOTALBUBBLE \ 2 + Level)
            CreateNewBubble
        Next i
    End If
End If

'Refresh the speedometer
frmBack.lblSpeed = Fix(Abs(ShipSpeed * 28.8)) & " Km/H"
'5 pix = 1m
'm/s = speed/5 *40
'km/h = (speed / 5) * 40 * (3600 / 1000)

'Refresh the time and timer label
GameTime = GameTime + CurDelay
TimeHours = Format((GameTime \ 3600000), "######0")  'Milliseconds in an hour
TimeMinutes = Format(((GameTime \ 60000) Mod 60), "00") 'Milliseconds in a minute
TimeSeconds = Format(((GameTime \ 1000) Mod 60), "00")


frmBack.lblTime = "Time: " & TimeHours & ":" & TimeMinutes & ":" & TimeSeconds



'Check for death
If CurLife <= 0 Then
    SaveRecords
    Pause False
    
End If
End Sub


Private Function FreeBullet() As Single
Dim i As Single

For i = 1 To UBound(Bullet)
    If Not Bullet(i).Active Then
        FreeBullet = i
        Exit For
    End If
Next i

End Function

Private Sub CreateNewBubble()
Dim FreeBubble As Single
Dim i As Integer

For i = 1 To UBound(Bubble)
    If Not Bubble(i).Active Then
        FreeBubble = i
        Exit For
    End If
Next i

If FreeBubble = 0 Then
    FreeBubble = (UBound(Bubble) + 1)
    ReDim Preserve Bubble(FreeBubble)
End If



With Bubble(FreeBubble):
    .Radius = Num(MINBUBBLERAD, MAXBUBBLERAD)
    .Speed = ((Num(10, MAXBUBBLESPEED * 10)) * 0.1)
    .Heading = (Num(0, Pi * 100) / 100)
    .Active = True
End With

Do
    Bubble(FreeBubble).x = Num(1, ScaleWidth - 1)
    Bubble(FreeBubble).y = Num(1, ScaleHeight - 1)
    
Loop Until (GetDist(Bubble(FreeBubble).x, Bubble(FreeBubble).y, ShipCenterX, ShipCenterY) >= 100)


End Sub



Private Function ColWithShip(BubbleX As Single, BubbleY As Single, BubbleRadius As Single) As Boolean

If CircCol(ShipCenterX, ShipCenterY, (SHIPFWDTH + SHIPRWDTH) * 0.5, BubbleX, BubbleY, BubbleRadius) Then
    ColWithShip = True
    Exit Function
End If
If CircCol(ShipX1, ShipY1, 0, BubbleX, BubbleY, BubbleRadius) Then
    ColWithShip = True
    Exit Function
End If
If CircCol(ShipX2, ShipY2, 0, BubbleX, BubbleY, BubbleRadius) Then
    ColWithShip = True
    Exit Function
End If
If CircCol(ShipX3, ShipY3, 0, BubbleX, BubbleY, BubbleRadius) Then
    ColWithShip = True
    Exit Function
End If
If CircCol(ShipX4, ShipY4, 0, BubbleX, BubbleY, BubbleRadius) Then
    ColWithShip = True
    Exit Function
End If
End Function

Private Sub ResetStats()

Score = 0
LeftKey = False
RightKey = False
UpKey = False
Reverse = False
Ebrakes = False
Brakes = False
Shoot = False
Dampers = False
QuadCannons = False
GunMode = 0
CurHeat = 0
HeatBarColor = vbYellow
frmBack.picHeatBar.BackColor = HeatBarColor
Running = False
InGameTimer = 0
ShipFacing = 0
ShipHeading = 0
ShipSpeed = 0
CurCannon = 0
CurBullet = 1
Level = 0
BonusLevels = 0
CHEATED = False
GameTime = 0
OverHeated = False
OverHeatTime = 0
ColorChangedTime = 0
ULTIMATEGUNACCESS = 0
frmBack.lblDampers = "Dampers: Off"
frmBack.lblQuadCannons = "QuadCannons: Off"
frmBack.lblGunMode = "GunMode: Normal"
frmBack.picHeatBar.Width = 1
frmBack.lblLevel = "0"


End Sub

Private Sub RespawnEnemies()
Dim i As Single
ReDim Enemy(Level \ 4)

ENEMYLIFE = 20 + (Level)
'Make sure the enemies are not already too small before making them smaller
If Not (16 - (Level * 0.25)) <= 5 Then
    ENEMYRADIUS = 16 - (Level * 0.25)
Else
    ENEMYRADIUS = 5
End If

For i = 0 To UBound(Enemy)
    With Enemy(i)
        .Active = True
        .Life = ENEMYLIFE
        .CurShotDelay = i * Num(0, EnemyGun.Delay)
        Do
            .x = Num(1, ScaleWidth - 1)
            .y = Num(1, ScaleHeight - 1)
        Loop Until (GetDist(.x, .y, ShipCenterX, ShipCenterY) >= MINDISTFROMSHIP)

    End With
Next i



With EnemyGun
    .Delay = 1300 - (Level * 10)
    .Lifespan = 1000 + (Level * 10)
    .CoolRate = 1
    .HeatRate = 1
    .Radius = 2 + (Level * 0.2)
    If .Radius > 20 Then .Radius = 20
    .Speed = 3 + (Level * 0.1)
    .StyleName = "Enemy Gun"
End With

End Sub



Private Sub DrawEnemies()
Dim i As Single
Dim a As Single
Dim RectX1 As Single, RectY1 As Single
Dim RectX2 As Single, RectY2 As Single
Dim Point As POINT_TYPE

picBackBuffer.ForeColor = vbBlue
For i = 0 To UBound(Enemy)
    If Enemy(i).Active Then
        RectX1 = Enemy(i).x - ENEMYRADIUS
        RectY1 = Enemy(i).y + ENEMYRADIUS
        RectX2 = Enemy(i).x + ENEMYRADIUS
        RectY2 = Enemy(i).y - ENEMYRADIUS
        
        MoveToEx MainDC, Enemy(i).x, Enemy(i).y, Point
        Ellipse MainDC, RectX1, RectY1, RectX2, RectY2
    Else
        For a = 0 To UBound(Enemy(i).Debris)
            SetPixel MainDC, Enemy(i).Debris(a).x, Enemy(i).Debris(a).y, vbBlue
        Next a
    End If
Next i
picBackBuffer.ForeColor = vbWhite


End Sub

Public Sub Save()
Dim i As Integer
On Error GoTo ErrOut
dlgSaveLoad.ShowSave
Open dlgSaveLoad.FileName For Binary Access Write As #SAVEGAMEFILE
Open App.Path & "/A31.rec" For Binary Access Read Write As #ENCRYPTFILE



Put SAVEGAMEFILE, 1, ShipX1
Put SAVEGAMEFILE, , ShipY1
Put SAVEGAMEFILE, , ShipX2
Put SAVEGAMEFILE, , ShipY2
Put SAVEGAMEFILE, , ShipX3
Put SAVEGAMEFILE, , ShipY3
Put SAVEGAMEFILE, , ShipX4
Put SAVEGAMEFILE, , ShipY4
Put SAVEGAMEFILE, , ShipCenterX
Put SAVEGAMEFILE, , ShipCenterY
Put SAVEGAMEFILE, , Score

Put SAVEGAMEFILE, , SHOTSPEED
Put SAVEGAMEFILE, , BULLETRAD
Put SAVEGAMEFILE, , SHOTDELAY
Put SAVEGAMEFILE, , BULLETLIFESPAN
Put SAVEGAMEFILE, , GUNCOOLRATE
Put SAVEGAMEFILE, , GUNHEATRATE

Put SAVEGAMEFILE, , BonusLevels
Put SAVEGAMEFILE, , NEWBUBBLEDELAY
Put SAVEGAMEFILE, , Level
'Put SAVEGAMEFILE, , MainDC

Put SAVEGAMEFILE, , GunMode




Put SAVEGAMEFILE, , CurLife
Put SAVEGAMEFILE, , CurHeat
Put SAVEGAMEFILE, , MAXHEAT
Put SAVEGAMEFILE, , MAXLIFE
Put SAVEGAMEFILE, , SHIPREGENRATE

For i = 0 To UBound(GunStyle)
    Put SAVEGAMEFILE, , GunStyle(i)
Next i

Put SAVEGAMEFILE, , EnemyGun

Put SAVEGAMEFILE, , ENEMYLIFE
Put SAVEGAMEFILE, , ENEMYRADIUS

Put SAVEGAMEFILE, , GameTime
Put SAVEGAMEFILE, , CHEATED

Put SAVEGAMEFILE, , ShipFacing
Put SAVEGAMEFILE, , ShipHeading

Close
DeleteFile App.Path & "/A31.rec"
ErrOut:
    Close
    DeleteFile App.Path & "/A31.rec"
End Sub

Public Sub Load()
Dim i
Dim MajorVersion As Integer
Dim MinorVersion As Integer
On Error GoTo ErrOut
dlgSaveLoad.ShowOpen
Open dlgSaveLoad.FileName For Binary Access Read As SAVEGAMEFILE

Get SAVEGAMEFILE, 1, ShipX1
Get SAVEGAMEFILE, , ShipY1
Get SAVEGAMEFILE, , ShipX2
Get SAVEGAMEFILE, , ShipY2
Get SAVEGAMEFILE, , ShipX3
Get SAVEGAMEFILE, , ShipY3
Get SAVEGAMEFILE, , ShipX4
Get SAVEGAMEFILE, , ShipY4
Get SAVEGAMEFILE, , ShipCenterX
Get SAVEGAMEFILE, , ShipCenterY
Get SAVEGAMEFILE, , Score

Get SAVEGAMEFILE, , SHOTSPEED
Get SAVEGAMEFILE, , BULLETRAD
Get SAVEGAMEFILE, , SHOTDELAY
Get SAVEGAMEFILE, , BULLETLIFESPAN
Get SAVEGAMEFILE, , GUNCOOLRATE
Get SAVEGAMEFILE, , GUNHEATRATE

Get SAVEGAMEFILE, , BonusLevels
Get SAVEGAMEFILE, , NEWBUBBLEDELAY
Get SAVEGAMEFILE, , Level
'Get SAVEGAMEFILE, , MainDC

Get SAVEGAMEFILE, , GunMode




Get SAVEGAMEFILE, , CurLife
Get SAVEGAMEFILE, , CurHeat
Get SAVEGAMEFILE, , MAXHEAT
Get SAVEGAMEFILE, , MAXLIFE
Get SAVEGAMEFILE, , SHIPREGENRATE

For i = 0 To UBound(GunStyle)
    Get SAVEGAMEFILE, , GunStyle(i)
Next i

Get SAVEGAMEFILE, , EnemyGun

Get SAVEGAMEFILE, , ENEMYLIFE
Get SAVEGAMEFILE, , ENEMYRADIUS

Get SAVEGAMEFILE, , GameTime
Get SAVEGAMEFILE, , CHEATED

Get SAVEGAMEFILE, , ShipFacing
Get SAVEGAMEFILE, , ShipHeading

frmBack.picLifeBack.Width = MAXLIFE
frmBack.picLifeBack.Height = 20
frmBack.picLifeBar.Width = CurLife
frmBack.picLifeBar.Height = 20


'Set the HeatBars
frmBack.picHeatBack.Width = MAXHEAT
frmBack.picHeatBack.Height = 20
frmBack.picHeatBar.Width = 0
frmBack.picHeatBar.Height = 20
HeatBarColor = vbYellow


MainDC = picBackBuffer.hdc
Running = True
InGameTimer = GetTickCount
CurDelay = 0


frmBack.lblGunMode = "GunMode: " & GunStyle(GunMode).StyleName

CurShotDel = 200 ' Makes sure the Ship can shoot right away
CurBullet = -1  ' Makes Sure that the first shot fired is #0 in the array



ReDim EnemyBullet(0)


RespawnEnemies
ReDim Bubble(TOTALBUBBLE - 1)
ReDim Bullet(1)

'Initialize the Bubble array
For i = 0 To (TOTALBUBBLE \ 2)
    CreateNewBubble
Next i

lblResume.Enabled = True
'Make Sure the Bubbles regenerate
BubbleRegen = True
InGameTimer = GetTickCount
RefreshBars
frmBack.Show
frmMain.Show

Close

RefreshBars
ErrOut:
    Close
End Sub

Private Sub InitGunStyles()

ReDim GunStyle(8)

With GunStyle(0)
    .Delay = 200
    .Lifespan = 1000
    .CoolRate = 1
    .HeatRate = 0
    .Radius = 2
    .Speed = 4
    .StyleName = "Laser"
End With

With GunStyle(1)
    .Delay = 125
    .Lifespan = 1000
    .CoolRate = 1
    .HeatRate = 6.625
    .Radius = 1.5
    .Speed = 5
    .StyleName = "Machine Laser"
End With

With GunStyle(2)
    .Delay = 50
    .Lifespan = 700
    .CoolRate = 1
    .HeatRate = 4.05
    .Radius = 1
    .Speed = 3
    .StyleName = "Eradicator"
End With

With GunStyle(3)
    .Delay = 400
    .Lifespan = 800
    .CoolRate = 1
    .HeatRate = 10
    .Radius = 4
    .Speed = 5
    .StyleName = "Cannon"
End With

With GunStyle(4)
    .Delay = 50
    .Lifespan = 450
    .CoolRate = 1
    .HeatRate = 5
    .Radius = 3
    .Speed = 3
    .StyleName = "Auto Cannon"
End With

With GunStyle(5)
    .Delay = 600
    .Lifespan = 3000
    .CoolRate = 1
    .HeatRate = 25.5
    .Radius = 5
    .Speed = 5
    .StyleName = "Sniper Cannon"
End With

With GunStyle(6)
    .Delay = 600
    .Lifespan = 500
    .CoolRate = 1
    .HeatRate = 3
    .Radius = 2
    .Speed = 5
    .StyleName = "The Infamous"
End With

With GunStyle(7)
    .Delay = 3000
    .Lifespan = 10000
    .CoolRate = 1
    .HeatRate = 120
    .Radius = 20
    .Speed = 4
    .StyleName = "The Big One"
End With

With GunStyle(0)
    SHOTSPEED = .Speed
    BULLETRAD = .Radius
    GUNCOOLRATE = .CoolRate
    GUNHEATRATE = .HeatRate
    SHOTDELAY = .Delay
    BULLETLIFESPAN = .Lifespan
    frmBack.lblGunMode = "GunMode: " & .StyleName
End With


'Enemy Gun
With EnemyGun
    .Delay = 1000
    .Lifespan = 1000
    .CoolRate = 1
    .HeatRate = 1
    .Radius = 1
    .Speed = 5
    .StyleName = "Enemy Gun"
End With

'The Ultimate Gun!!!!!
With GunStyle(8)
    .Delay = 25
    .Lifespan = 2000    'I still limit this because my shibby 166 megahurtz computer will slow down if there are too many bullets on screen!
    .CoolRate = 1
    .HeatRate = 1
    .Radius = 5
    .Speed = 6
    .StyleName = "ULTIMATE"
End With
End Sub

Private Sub SaveRecords()
Dim HighScore As Double
Dim TempName As String * 20
Dim RecordHoldersName As String * 20
Dim DateSet As Variant
Dim RetVal As VbMsgBoxResult
Dim Temp As Byte
Dim i As Single


Dim EmptyScore As Double
Dim EmptyName As String * 20
Dim EmptyDate As Variant

Dim FileData() As Byte




If Not CHEATED Then
    Open App.Path & "/HighScores.rec" For Binary Access Read Write As #HIGHSCOREFILE
    Open App.Path & "/A31.rec" For Binary Access Read Write As #ENCRYPTFILE
    If Not FileLen(App.Path & "/HighScores.rec") = 0 Then
        'Get the encrypted numbers
        ReDim FileData(RECORDSIZE)
        For i = 1 To RECORDSIZE
            Get HIGHSCOREFILE, i, FileData(i)
        Next i
        
        'Put the decrypted data into the temporary file
        For i = 1 To RECORDSIZE
            Put ENCRYPTFILE, i, FileData(i) Xor ENCRYPTKEY
        Next i
        
        'Get the decrypted data
        
        Get ENCRYPTFILE, 1, HighScore
        Get ENCRYPTFILE, , RecordHoldersName
        Get ENCRYPTFILE, , DateSet
        
    End If
    
    If Score > HighScore Then
        If Not HighScore = 0 Then
            TempName = InputBox("Congratulations, you beat the High Score of " & HighScore & ", set by " & Trim(RecordHoldersName) & " on " & DateSet & ". Your Score: " & Score & ". Please enter your name.", "High Score", "Nobody")
        Else
            TempName = InputBox("Congratulations, you set the high score! Your score: " & Score & ". Please enter your name.", "High Score", "Nobody")
        End If
        If TempName = "" Then TempName = "Nobody"
        DateSet = Now
        RecordHoldersName = TempName
        'Put the unencrypted data back into the temporary file
        Put ENCRYPTFILE, 1, Score
        Put ENCRYPTFILE, , TempName
        Put ENCRYPTFILE, , Now
        
        'Get the data
        ReDim FileData(RECORDSIZE)
        For i = 1 To RECORDSIZE
            Get ENCRYPTFILE, i, FileData(i)
        Next i
        
        'Put the encrypted data back into the permanent file
        For i = 1 To RECORDSIZE
            Put HIGHSCOREFILE, i, FileData(i) Xor ENCRYPTKEY
        Next i
        End If
    
    

End If
    Close
    DeleteFile App.Path & "/A31.rec"
End Sub


Private Sub Restart()
SaveRecords

Paused = False
        
lblHighScore.Visible = False
lblHS.Visible = False
lblRecordHolder.Visible = False
lblDateSet.Visible = False
lblExit.Visible = False
lblLoad.Visible = False
lblNewGame.Visible = False
lblPaused.Visible = False
lblResume.Visible = False
lblSave.Visible = False

ResetStats
Form_Load

End Sub



Private Sub lblExit_Click()
Dim RetVal As VbMsgBoxResult

    RetVal = MsgBox("Are you sure you want to quit?", vbYesNo, "Quit Game")
    If RetVal = vbYes Then
        SaveRecords
        Form_Unload 0
    End If
End Sub

Private Sub lblLoad_Click()
    SaveRecords
    Load
End Sub

Private Sub lblNewGame_Click()
Dim RetVal As VbMsgBoxResult

    RetVal = MsgBox("Are you sure you want to restart?", vbYesNo, "Quit Game")
    If RetVal = vbYes Then
        Restart
    End If
End Sub

Private Sub lblResume_Click()
    Pause True
End Sub

Private Sub lblSave_Click()
    Save
End Sub
