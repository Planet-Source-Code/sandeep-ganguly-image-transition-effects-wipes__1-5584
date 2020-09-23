VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form WipesForm 
   BackColor       =   &H00000000&
   Caption         =   "Image Wipes"
   ClientHeight    =   8190
   ClientLeft      =   1800
   ClientTop       =   2640
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   9615
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Wipes.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Wipes.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Wipes.frx":0224
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Wipes.frx":0676
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Wipes.frx":0AC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Wipes.frx":0DE2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   1005
      ButtonWidth     =   2884
      ButtonHeight    =   953
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   20
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Make New Script"
            Key             =   "tlbNew"
            Object.ToolTipText     =   "Opens scripter. Select File->New to make a new script"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open & Edit Script"
            Key             =   "tlbOpen"
            Object.ToolTipText     =   "Opens an existing script for editing"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Execute Script"
            Key             =   "tlbExecute"
            Object.ToolTipText     =   "Executes the currently selected script. "
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Set Options"
            Key             =   "tlbSetDirectory"
            Object.ToolTipText     =   "Set Directory and other options for preset show"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Play Preset Animation"
            Key             =   "tlbRunPreset"
            Object.ToolTipText     =   "Runs preset show"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Play Music"
            Key             =   "tlbPlayMusic"
            Object.ToolTipText     =   "Play ""Every breath you take"" midi !"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      Begin MCI.MMControl MMControl1 
         Height          =   330
         Left            =   10440
         TabIndex        =   7
         Top             =   120
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         _Version        =   393216
         PlayEnabled     =   -1  'True
         PauseEnabled    =   -1  'True
         StopEnabled     =   -1  'True
         PrevVisible     =   0   'False
         NextVisible     =   0   'False
         BackVisible     =   0   'False
         StepVisible     =   0   'False
         RecordVisible   =   0   'False
         EjectVisible    =   0   'False
         DeviceType      =   ""
         FileName        =   "F:\my vb projs\Wipes\Sounds\Welcom98.wav"
      End
   End
   Begin VB.CommandButton cmdoundeSelect 
      Caption         =   "&Play"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Appearance      =   0  'Flat
      Caption         =   "Start Show"
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   120
      Pattern         =   "*.jpg"
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3450
      Left            =   2760
      ScaleHeight     =   230
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   502
      TabIndex        =   2
      Top             =   1920
      Width           =   7530
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         BorderWidth     =   3
         Index           =   0
         Visible         =   0   'False
         X1              =   0
         X2              =   264
         Y1              =   232
         Y2              =   232
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8880
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox ClearDestination 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   3030
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   3450
      Left            =   2760
      Picture         =   "Wipes.frx":1234
      ScaleHeight     =   226
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   498
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   7530
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   240
      Stretch         =   -1  'True
      Top             =   3720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuPicture 
         Caption         =   "Load Picture"
      End
      Begin VB.Menu mnuSetdir 
         Caption         =   "Set Directory"
      End
      Begin VB.Menu mnuScripter 
         Caption         =   "Launch Scripter"
      End
      Begin VB.Menu mnuLoadAndExecScript 
         Caption         =   "Execute Script"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuTransition 
      Caption         =   "Transition Effects"
      Begin VB.Menu mnuWipex 
         Caption         =   "WIPES"
         WindowList      =   -1  'True
         Begin VB.Menu mnuWipeLeft 
            Caption         =   "Wipe Left"
         End
         Begin VB.Menu mnuWipeRight 
            Caption         =   "Wipe Right"
         End
         Begin VB.Menu mnuWipeTop 
            Caption         =   "Wipe Top"
         End
         Begin VB.Menu mnuWipeBottom 
            Caption         =   "Wipe Bottom"
         End
         Begin VB.Menu mnuWipe4Corner 
            Caption         =   "Wipe From All 4 Corner"
         End
         Begin VB.Menu mnuWipeTLCorner 
            Caption         =   "Wipe From TL Corner"
         End
         Begin VB.Menu mnuWipeTRCorner 
            Caption         =   "Wipe From TR Corner"
         End
         Begin VB.Menu mnuWipeBLCorner 
            Caption         =   "Wipe from BL Corner"
         End
         Begin VB.Menu mnuWipeBRCorner 
            Caption         =   "Wipe BR Corner"
         End
         Begin VB.Menu mnuwipeRightLeft 
            Caption         =   "Wipe Right && Left"
         End
         Begin VB.Menu mnuWipeUpDown 
            Caption         =   "Wipe Up && Down"
         End
         Begin VB.Menu mnuWipeCentre 
            Caption         =   "Wipe Centre"
         End
      End
      Begin VB.Menu mnuStrectchex 
         Caption         =   "STRETCHES"
         Begin VB.Menu mnuStretchRight 
            Caption         =   "Stretch from Right"
         End
         Begin VB.Menu mnuStretchLeft 
            Caption         =   "Stretch from Left"
         End
         Begin VB.Menu mnuStretchBottom 
            Caption         =   "Stretch From Bottom"
         End
         Begin VB.Menu mnuStretchTop 
            Caption         =   "Stretch from Top"
         End
      End
      Begin VB.Menu mnuStripes 
         Caption         =   "STRIPES"
         Begin VB.Menu mnuHorizontalStripes 
            Caption         =   "Horizontal Stripes"
         End
         Begin VB.Menu mnuVerticalStripes 
            Caption         =   "Vertical Stripes"
         End
      End
      Begin VB.Menu mnuBlinds 
         Caption         =   "BLINDS"
         Begin VB.Menu mnuHorizontalBlinds 
            Caption         =   "Horizontal Blinds"
         End
         Begin VB.Menu mnuVerticalBlinds 
            Caption         =   "Vertical Blinds"
         End
      End
      Begin VB.Menu mnuOthers 
         Caption         =   "OTHERS"
         Begin VB.Menu mnuRandomBlocks 
            Caption         =   "Random Blocks"
         End
         Begin VB.Menu mnuCircularWipe 
            Caption         =   "Circular Wipe"
         End
         Begin VB.Menu mnuWipeFade 
            Caption         =   "Fade"
         End
         Begin VB.Menu mnuGrowCentre 
            Caption         =   "Grow from Centre-"
         End
         Begin VB.Menu mnuS_Click 
            Caption         =   "Show S"
         End
         Begin VB.Menu mnuReverseS 
            Caption         =   "Show Reverse S"
         End
         Begin VB.Menu mnuMaze 
            Caption         =   "Maze Out"
         End
      End
      Begin VB.Menu mnuSlidex 
         Caption         =   "SLIDES"
         Begin VB.Menu mnuSlideLeft 
            Caption         =   "Slide To Left"
         End
         Begin VB.Menu mnuSlideRight 
            Caption         =   "Slide To Right"
         End
         Begin VB.Menu mnuSlideTop 
            Caption         =   "Slide To Top"
         End
         Begin VB.Menu mnuSlideBottom 
            Caption         =   "Slide To Bottom"
         End
      End
   End
End
Attribute VB_Name = "WipesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Dim XCoords() As Integer
Dim YCoords() As Integer
Dim XCords(4, 4) As Integer
Dim YCords(4, 4) As Integer

Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Dim drawcode As Long
Dim DELAY_SECS As Integer
Dim RANDOM_BLOCKS As Integer
Dim FADE_STEPS As Integer
Dim SoundLoop As Boolean
Dim StopImmidiately As Boolean

Private Sub cmdoundeSelect_Click()
If MsgBox("Do you want the sound to loop infinitely ?", vbQuestion + vbYesNo, "Play sound once/infinitely?") = vbYes Then
    SoundLoop = True
Else
    SoundLoop = False
End If

MMControl1.FileName = App.Path & "\" & "Sounds\everybreath[1].mid"
MMControl1.Command = "Open"
MMControl1.Command = "Play"
End Sub

Private Sub cmdStart_Click()

Dim bDone As Boolean, bInfinite As Boolean, bClear As Boolean
Dim picname As String
Dim nFileno As Integer, nTransitNo As Integer

bDone = False
bInfinite = Form1.chkInfinite.Value
bClear = Form1.chkClearDest.Value
nFileno = 0
nTransitNo = 1
Dim delayTransit As Integer
delayTransit = val(Form1.txtDelay.Text)

DELAY_SECS = val(Form1.txtDelaySteps.Text)
RANDOM_BLOCKS = val(Form1.txtRandomBlocks.Text)
FADE_STEPS = val(Form1.txtFadeSteps.Text)

While bDone = False
    If StopImmidiately = True Then
        Picture1.Picture = LoadPicture()
        Picture2.Picture = LoadPicture()
        StopImmidiately = False
        Exit Sub
    End If
DoEvents
    If File1.ListCount = 0 Then
        Exit Sub
    End If
    picname = File1.Path & "\" & File1.List(nFileno)
    
    ShowTransit picname, nTransitNo, bClear, delayTransit
   nFileno = nFileno + 1
        If nFileno >= File1.ListCount Then
            nFileno = 0
            If bInfinite = False Then
                bDone = True
            Else
                bDone = False
            End If
        End If
            
            nTransitNo = nTransitNo + 1
            If nTransitNo > 31 Then
                nTransitNo = 1
            End If
Wend
End Sub
Private Sub Exit_Click()
    End
End Sub


Private Sub Form_Load()
Dim files As Integer
File1.Path = App.Path & "\" & "Pictures"
DELAY_SECS = 1
RANDOM_BLOCKS = 9
FADE_STEPS = 3
drawcode = vbSrcCopy
MMControl1.FileName = App.Path & "\" & "Sounds\everybreath[1].mid"
MMControl1.Command = "Open"
End Sub

Private Sub Form_Paint()
Gradient Me, 0, 255, 255, False
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Image1_Click()
Image1.Visible = False
Picture2.Visible = True
End Sub

Private Sub MMControl1_Done(NotifyCode As Integer)
Me.Caption = "Music loop called"

If ScriptMaker.lstLoopSound.Text = "Y" Or SoundLoop = True Then 'if infinite loop reqd
    If MMControl1.Position = MMControl1.Length Then
        MMControl1.Command = "Prev"
        MMControl1.Command = "play"
    End If
End If
End Sub

Private Sub mnuCircularWipe_Click()
On Error Resume Next

Dim ax As Integer, ay As Integer
Dim bx As Integer, by As Integer
Dim cx As Integer, cy As Integer
Dim dx As Integer, dy As Integer
Dim midx As Integer, midy As Integer
Dim lines As Integer

WipesForm.Refresh

'work out the coords
midx = Picture2.ScaleWidth / 2 'centre
midy = Picture2.ScaleHeight / 2
Point midx, midy

'ax = Picture2.Left 'top-left
ax = 0
'ay = Picture2.Top
ay = 0

bx = Picture2.ScaleWidth 'top-right
by = 0

cx = Picture2.ScaleWidth 'bottom-right
cy = Picture2.ScaleHeight

'dx = Picture2.Left 'bottom-left
dx = 0
dy = Picture2.ScaleHeight
Dim I As Integer
For I = 1 To Picture2.ScaleHeight
    'If i Mod 100 = 0 Then DoEvents
    Load Line1(I)
    Line1(I).X1 = midx
    Line1(I).Y1 = midy
    Line1(I).X2 = bx
    Line1(I).Y2 = by + I
    Line1(I).Visible = True
Next I
lines = I 'keep the number of lines
Dim x As Integer
x = 0
For I = cx To 1 Step -1
x = x + 1
    'If i Mod 100 = 0 Then DoEvents
    Load Line1(lines + x)
    Line1(lines + x).X1 = midx
    Line1(lines + x).Y1 = midy
    Line1(lines + x).X2 = cx - I
    Line1(lines + x).Y2 = cy
    Line1(lines + x).Visible = True
Next I
lines = lines + x
x = 0

'--
For I = Picture2.ScaleHeight To 1 Step -1
x = x + 1
    'If i Mod 100 = 0 Then DoEvents
    Load Line1(lines + x)
    Line1(lines + x).X1 = midx
    Line1(lines + x).Y1 = midy
    Line1(lines + x).X2 = dx
    Line1(lines + x).Y2 = dy - I
    Line1(lines + x).Visible = True
Next I
lines = lines + x
x = 0
'--

For I = 1 To Picture2.ScaleWidth
x = x + 1
    'If i Mod 100 = 0 Then DoEvents
    Load Line1(lines + x)
    Line1(lines + x).X1 = midx
    Line1(lines + x).Y1 = midy
    Line1(lines + x).X2 = ax + I
    Line1(lines + x).Y2 = ay
    Line1(lines + x).Visible = True
Next I
lines = lines + x
x = 0
DoEvents
Picture2.Picture = LoadPicture() 'mnuwipeRightLeft_Click
DoEvents
Picture2.Picture = Picture1.Picture
For I = 1 To lines
    Unload Line1(I)
Next I
End Sub

Private Sub mnuGrowCentre_Click()
On Error Resume Next
Dim midx As Integer
Dim midy As Integer
Dim setX As Integer
Dim setY As Integer
Dim setWidth As Integer
Dim setHeight As Integer
Dim rsetd As Integer

Picture1.Visible = False
Picture2.Visible = False
Picture2.Picture = Picture1.Picture
Image1.Picture = Picture1.Picture
midx = WipesForm.ScaleWidth / 2
midy = WipesForm.ScaleHeight / 2

rsteps = (WipesForm.ScaleWidth / 4) / 20

setX = midx
setY = midy
setWidth = 10
setHeight = 10
Image1.top = midy
Image1.Left = midx
Image1.width = 100
Image1.height = 100
Image1.Visible = True

While setX > 0 Or setY > 0
setX = setX - rsteps
    If setX < WipesForm.ScaleLeft Then
    delayNext 2
    Image1.Visible = False
    'Picture2.Cls
    Picture2.Visible = True
    Exit Sub
    End If

setY = setY - rsteps
If setY < WipesForm.ScaleTop Then
delayNext 2
Image1.Visible = False
Picture2.Visible = True
Exit Sub
End If


    setHeight = setHeight + (rsteps * 2)
    If setHeight >= WipesForm.ScaleHeight Then
    delayNext 2
    Image1.Visible = False
    Picture2.Visible = True
    Exit Sub
     End If

    
    setWidth = setWidth + (rsteps * 2)
    If setWidth >= WipesForm.ScaleWidth Then
    delayNext 2
    Image1.Visible = False
    Picture2.Visible = True
    Exit Sub

        End If

    
    delayNext 0.75
    Image1.Visible = False
    Image1.top = setY
    Image1.Left = setX
    Image1.width = setWidth
    Image1.height = setHeight
    Image1.Visible = True
    DoEvents
    DoEvents
Wend
delayNext 2
Image1.Visible = False
'Picture2.Cls
Picture2.Visible = True
Picture2.Cls

End Sub

Private Sub mnuHorizontalBlinds_Click()
Dim Stripes As Integer
Dim I As Integer, j As Integer
Dim StripeHeight As Integer

    If ClearDestination.Value Then Picture2.Picture = LoadPicture()
    StripeHeight = 20
    Stripes = Fix(Picture1.ScaleHeight / StripeHeight)
    On Error Resume Next
    For j = 1 To StripeHeight
        For I = 0 To Stripes
            Picture2.PaintPicture Picture1.Picture, 0, I * StripeHeight, _
            Picture1.ScaleWidth, j, _
            0, I * StripeHeight, _
            Picture1.ScaleWidth, j, drawcode
        Next
    Next
End Sub

Private Sub mnuHorizontalStripes_Click()
Dim TotalNoOfStripes As Integer 'max 50
Dim ImageWidth As Integer
Dim imageHeight As Integer
Dim EvenStripes_Stpoint(25) As Integer
Dim OddStripes_Stpoint(25) As Integer
Dim EvenStripes_Endpoint(25) As Integer
Dim OddStripes_Endpoint(25) As Integer
Dim I As Integer, valx As Integer, j As Integer
Dim StripeHeight As Integer


'TotalNoOfStripes = ??
Dim tm As String
tm = Str(Time) 'full time
tm = Right(tm, 5) ' returns SS AM/PM
tm = Left(tm, 2) 'return SS
TotalNoOfStripes = val(tm) ' to val

If TotalNoOfStripes > 50 Then
    TotalNoOfStripes = 50
End If
If TotalNoOfStripes < 12 Then
    TotalNoOfStripes = 12
End If
Me.Caption = "Stripes=" & TotalNoOfStripes

'Picture2.Cls
ImageWidth = Picture1.ScaleWidth
imageHeight = Picture1.ScaleHeight
StripeHeight = imageHeight / TotalNoOfStripes

valx = 0

EvenStripes_Stpoint(0) = 1
'For i = 1 To 4
For I = 1 To (TotalNoOfStripes - 2) / 2
    valx = I * (2 * StripeHeight) + 1
    EvenStripes_Stpoint(I) = valx
Next I

For I = 1 To TotalNoOfStripes / 2
    EvenStripes_Endpoint(I - 1) = EvenStripes_Stpoint(I - 1) + StripeHeight
Next I

For I = 1 To TotalNoOfStripes / 2
    OddStripes_Stpoint(I - 1) = EvenStripes_Endpoint(I - 1)
Next I

For I = 1 To TotalNoOfStripes / 2
    OddStripes_Endpoint(I - 1) = OddStripes_Stpoint(I - 1) + StripeHeight
Next I
I = 0

Dim s As Integer 'step
s = 0 'default step
For I = 1 To Picture1.ScaleWidth
    For j = 0 To TotalNoOfStripes / 2 - 1 'for even stripes 4
        y = EvenStripes_Stpoint(j)
        'Me.Caption = y
        Picture2.PaintPicture Picture1.Picture, Picture1.ScaleWidth - I, y, I, StripeHeight, 0, y, I, StripeHeight, drawcode
    Next j
    
    For w = 0 To TotalNoOfStripes / 2 - 1 'for odd stripes -4
        y = OddStripes_Stpoint(w)
        Picture2.PaintPicture Picture1.Picture, 0, OddStripes_Stpoint(w), I, StripeHeight, 0, OddStripes_Stpoint(w), I, StripeHeight, drawcode
     Next w
Next I

'Frame1.Visible = True
End Sub

Private Sub mnuLoadAndExecScript_Click()
Dim AnimationDone As Boolean
Dim SoundDone As Boolean
Dim aLoop As Boolean

AnimationDone = False
Load Selection
Load ScriptMaker
ScriptMaker.Show
Selection.Label1.Caption = "Load Script"
Selection.Image1.Visible = False
Selection.File1.Pattern = "*.tsf"
Selection.Show (vbModal)
'ScriptMaker.Hide

ScriptMaker.Data1.Refresh
ScriptMaker.Data1.Recordset.MoveLast
Dim recs As Integer
recs = ScriptMaker.Data1.Recordset.RecordCount
Dim I As Integer
Dim mPicname As String
Dim mTranstype As Integer
Dim mDrawmode As Long
Dim mUniversalNo As Integer
Dim mHoldImage As Integer
ScriptMaker.Data1.Recordset.MoveFirst
ScriptMaker.Data2.Refresh
ScriptMaker.Data2.Recordset.MoveFirst
ScriptMaker.Hide
If ScriptMaker.lstLoopAnimation.Text = "Y" Then
    aLoop = True
Else
    aLoop = False
End If
MMControl1.DeviceType = "Sequencer"
MMControl1.FileName = Variables.SelectedSoundName
MMControl1.Command = "Open"
MMControl1.Command = "Play"
While AnimationDone = False
    For I = 0 To recs - 1
    If StopImmidiately = True Then
        Picture1.Picture = LoadPicture()
        Picture2.Picture = LoadPicture()
        StopImmidiately = False
        Exit Sub
    End If
        mPicname = ScriptMaker.txtPicName.Text
        mTranstype = val(ScriptMaker.txtTransitionType.Text)
        If ScriptMaker.cmbDrawMode = "BLACKNESS" Then
            mDrawmode = vbBlackness
        End If
            
        If ScriptMaker.cmbDrawMode = "WHITENESS" Then
            mDrawmode = vbWhiteness
        End If
        
        If ScriptMaker.cmbDrawMode = "DSTINVERT" Then
            mDrawmode = vbDstInvert
        End If
        
        If ScriptMaker.cmbDrawMode = "SRCAND" Then
            mDrawmode = vbSrcAnd
        End If
        
        If ScriptMaker.cmbDrawMode = "MERGEPAINT" Then
            mDrawmode = vbMergePaint
        End If
        
        If ScriptMaker.cmbDrawMode = "SRCPAINT" Then
            mDrawmode = vbSrcPaint
        End If
        
        If ScriptMaker.cmbDrawMode = "NOTSRCERASE" Then
            mDrawmode = vbNotSrcErase
        End If
        
        If ScriptMaker.cmbDrawMode = "SRCINVERT (XOR PEN)" Then
            mDrawmode = vbSrcInvert
        End If
        
        If ScriptMaker.cmbDrawMode = "SRCCOPY   (COPY PEN)" Then
            mDrawmode = vbSrcCopy
        End If
        
        If ScriptMaker.cmbDrawMode = "SRCERASE" Then
            mDrawmode = vbSrcErase
        End If
        
        If ScriptMaker.cmbDrawMode = "NOTSRCCOPY" Then
            mDrawmode = vbNotSrcCopy
        End If
        
        If ScriptMaker.cmbDrawMode = "NOTSRCINVERT" Then
            mDrawmode = vbNotSrcCopy
        End If
            
        mUniveralNo = val(ScriptMaker.cmbTransitionSteps.Text)
        mHoldImage = val(ScriptMaker.cmbHoldImage.Text)
        
        DELAY_SECS = mUniveralNo
        RANDOM_BLOCKS = mUniveralNo
        FADE_STEPS = mUniveralNo
        ShowTransit2 mPicname, mTranstype, False, mHoldImage, mDrawmode
        If ScriptMaker.Data1.Recordset.EOF Then
        ScriptMaker.Data1.Recordset.MoveFirst
        Else
        ScriptMaker.Data1.Recordset.MoveNext
        End If
    Next I
    If aLoop = True Then
        AnimationDone = False
        I = 0
    Else
        AnimationDone = True
    End If
Wend
End Sub

Private Sub mnuMaze_Click()
Dim blockWidth As Integer, blockHeight As Integer
Dim xVal As Integer, yVal As Integer
Dim roew As Integer, col As Integer

blockWidth = Picture2.ScaleWidth / 5
blockHeight = Picture2.ScaleHeight / 5
yVal = 0
For row = 0 To 4
xVal = 0
    For col = 0 To 4
        XCords(row, col) = xVal
        xVal = xVal + blockWidth
        YCords(row, col) = yVal
    Next col
yVal = yVal + blockHeight
Next row
DrawMazeBlock 2, 2, blockWidth, blockHeight
delayNext 0.5
DrawMazeBlock 2, 3, blockWidth, blockHeight
delayNext 0.5
DrawMazeBlock 1, 3, blockWidth, blockHeight
delayNext 0.5
DrawMazeBlock 1, 2, blockWidth, blockHeight
delayNext 0.5
DrawMazeBlock 1, 1, blockWidth, blockHeight
delayNext 0.5
DrawMazeBlock 2, 1, blockWidth, blockHeight
delayNext 0.5
DrawMazeBlock 3, 1, blockWidth, blockHeight
delayNext 0.5
DrawMazeBlock 3, 2, blockWidth, blockHeight
delayNext 0.5
DrawMazeBlock 3, 3, blockWidth, blockHeight
delayNext 0.5
DrawMazeBlock 3, 4, blockWidth, blockHeight
delayNext 0.5
DrawMazeBlock 2, 4, blockWidth, blockHeight
delayNext 0.5
DrawMazeBlock 1, 4, blockWidth, blockHeight
delayNext 0.5
DrawMazeBlock 0, 4, blockWidth, blockHeight
delayNext 0.5
DrawMazeBlock 0, 3, blockWidth, blockHeight
delayNext 0.5
DrawMazeBlock 0, 2, blockWidth, blockHeight
delayNext 0.5
DrawMazeBlock 0, 1, blockWidth, blockHeight
delayNext 0.5
DrawMazeBlock 0, 0, blockWidth, blockHeight
delayNext 0.5
DrawMazeBlock 1, 0, blockWidth, blockHeight
delayNext 0.5
DrawMazeBlock 2, 0, blockWidth, blockHeight
delayNext 0.5
DrawMazeBlock 3, 0, blockWidth, blockHeight
delayNext 0.5
DrawMazeBlock 4, 0, blockWidth, blockHeight
delayNext 0.5
DrawMazeBlock 4, 1, blockWidth, blockHeight
delayNext 0.5
DrawMazeBlock 4, 2, blockWidth, blockHeight
delayNext 0.5
DrawMazeBlock 4, 3, blockWidth, blockHeight
delayNext 0.5
DrawMazeBlock 4, 4, blockWidth, blockHeight
delayNext 0.5
End Sub

Private Sub mnuPicture_Click()

    If ClearDestination.Value Then Picture2.Picture = LoadPicture()
   
    CommonDialog1.Filter = "Images|*.BMP;*.GIF;*.JPG"
    CommonDialog1.Action = 1
    If CommonDialog1.FileName = "" Then Exit Sub
    Picture1.Picture = LoadPicture(CommonDialog1.FileName)
    Picture2.width = Picture1.width
    Picture2.height = Picture1.height
    Picture2.Left = (WipesForm.width - Picture2.width) / 2
    Picture2.top = (WipesForm.height - Picture2.height) / 2
    

End Sub
Private Sub mnuRandomBlocks_Click()
If ClearDestination.Value Then Picture2.Picture = LoadPicture()
'there is a good chance that all the blocks
'will not be generated using rnd
'so draw the picture from 1 to 2 at the end
'of 5 seconds

Dim blockWidth As Integer, blockHeight As Integer
Dim no As Integer, num As Integer, TotalBlocks As Integer
Dim bDone() As Boolean
Dim AllDone As Boolean
Dim I As Integer, xVal As Integer, yVal As Integer

ReDim XCoords(1)
ReDim YCoords(1)
ReDim bDone(1)

Picture2.width = Picture1.width
Picture2.height = Picture1.height
AllDone = False

'uncomment to customize
'no = InputBox("Enter no of bloacks...", "How many blocks?", 9)

no = RANDOM_BLOCKS
TotalBlocks = (no * no) - 1

ReDim XCoords(no)
ReDim YCoords(no)
ReDim bDone(TotalBlocks)

For I = 1 To TotalBlocks
    bDone(I - 1) = False
Next I

xVal = 0
yVal = 0
blockWidth = Picture2.ScaleWidth / no
blockHeight = Picture2.ScaleHeight / no
Dim endtime As Integer
endtime = TotalBlocks / 15

If endtime > 10 Then
    endtime = 10
    blockWidth = blockWidth * 2
    blockHeight = blockHeight * 2
End If

    WipesForm.Caption = endtime

For I = 1 To no
    XCoords(I - 1) = xVal
    xVal = xVal + blockWidth
    
    YCoords(I - 1) = yVal
    yVal = yVal + blockHeight
Next I

AllDone = False

t1 = Timer
While AllDone = False
genNumber:
    
    If Timer - t1 > endtime Then 'if 7 seconds pass
        Picture2.Picture = Picture1.Picture
    Exit Sub
    End If
    
    num = Int(Rnd * (TotalBlocks - 1))
    
    If bDone(num) = True Then
        GoTo genNumber
    Else
        bDone(num) = True
        DrawBlocks num, no, blockWidth, blockHeight
        delayNext 0.01
    End If

    Dim togo As Integer
    togo = 0
    For I = 0 To TotalBlocks
        If bDone(I) = True Then
            togo = togo + 1
        End If
    Next I
    If togo = TotalBlocks Then
        AllDone = True
    End If
Wend
End Sub
Private Sub mnuReverseS_Click()
Dim blockWidth As Integer, blockHeight As Integer
Dim no As Integer, num As Integer, TotalBlocks As Integer
Dim I As Integer, xVal As Integer, yVal As Integer

ReDim XCoords(9)
ReDim YCoords(9)
If ClearDestination.Value Then Picture2.Picture = LoadPicture()
Picture2.width = Picture1.width
Picture2.height = Picture1.height
AllDone = False
no = 3
TotalBlocks = (no * no) - 1

xVal = 0
yVal = 0
blockWidth = Picture2.ScaleWidth / no
blockHeight = Picture2.ScaleHeight / no


For I = 1 To no
    XCoords(I - 1) = xVal
    xVal = xVal + blockWidth
    
    YCoords(I - 1) = yVal
    yVal = yVal + blockHeight
Next I
                DrawBlocks 5, no, blockWidth, blockHeight
        delayNext 1
                DrawBlocks 7, no, blockWidth, blockHeight
        delayNext 1
                DrawBlocks 6, no, blockWidth, blockHeight
        delayNext 1
                DrawBlocks 3, no, blockWidth, blockHeight
        delayNext 1
                DrawBlocks 4, no, blockWidth, blockHeight
        delayNext 1
                DrawBlocks 2, no, blockWidth, blockHeight
        delayNext 1
                DrawBlocks 8, no, blockWidth, blockHeight
        delayNext 1
                DrawBlocks 1, no, blockWidth, blockHeight
        delayNext 1
                DrawBlocks 0, no, blockWidth, blockHeight
        delayNext 1
End Sub

Private Sub mnuS_Click_Click()
Dim blockWidth As Integer, blockHeight As Integer
Dim no As Integer, num As Integer, TotalBlocks As Integer
Dim I As Integer, xVal As Integer, yVal As Integer

ReDim XCoords(9)
ReDim YCoords(9)
If ClearDestination.Value Then Picture2.Picture = LoadPicture()
Picture2.width = Picture1.width
Picture2.height = Picture1.height
AllDone = False
no = 3
TotalBlocks = (no * no) - 1

xVal = 0
yVal = 0
blockWidth = Picture2.ScaleWidth / no
blockHeight = Picture2.ScaleHeight / no

For I = 1 To no
    XCoords(I - 1) = xVal
    xVal = xVal + blockWidth
    
    YCoords(I - 1) = yVal
    yVal = yVal + blockHeight
Next I
        DrawBlocks 0, no, blockWidth, blockHeight
        delayNext 1

        DrawBlocks 1, no, blockWidth, blockHeight
        delayNext 1

        DrawBlocks 8, no, blockWidth, blockHeight
        delayNext 1
        
        DrawBlocks 2, no, blockWidth, blockHeight
        delayNext 1

        DrawBlocks 4, no, blockWidth, blockHeight
        delayNext 1
        
        DrawBlocks 3, no, blockWidth, blockHeight
        delayNext 1

        DrawBlocks 6, no, blockWidth, blockHeight
        delayNext 1

        DrawBlocks 7, no, blockWidth, blockHeight
        delayNext 1

        DrawBlocks 5, no, blockWidth, blockHeight
        delayNext 1

End Sub

Private Sub mnuScripter_Click()
ScriptMaker.Show
End Sub

Private Sub mnuSetdir_Click()
Form1.Show
End Sub

Private Sub mnuSlideBottom_Click()
On Error Resume Next
Dim I As Integer
For I = 0 To Picture1.ScaleHeight Step DELAY_SECS
    Picture2.PaintPicture Picture1.Picture, 0, 0, Picture1.ScaleWidth, I, 0, Picture2.ScaleHeight - I, Picture1.ScaleWidth, I, drawcode
Next I
End Sub

Private Sub mnuSlideLeft_Click()
On Error Resume Next
Dim I As Integer
For I = 0 To Picture1.ScaleWidth Step DELAY_SECS
    Picture2.PaintPicture Picture1.Picture, Picture1.ScaleWidth - I, 0, I, Picture1.ScaleHeight, Picture2.ScaleLeft, 0, I, Picture1.ScaleHeight, drawcode
Next I
delayNext 2
End Sub

Private Sub mnuSlideRight_Click()
On Error Resume Next
Dim I As Integer
For I = 0 To Picture1.ScaleWidth Step DELAY_SECS
    Picture2.PaintPicture Picture1.Picture, 0, 0, I, Picture1.ScaleHeight, Picture2.ScaleWidth - I, 0, I, Picture1.ScaleHeight, drawcode
Next I
End Sub

Private Sub mnuSlideTop_Click()
On Error Resume Next
Dim I As Integer
For I = 0 To Picture1.ScaleHeight Step DELAY_SECS
    Picture2.PaintPicture Picture1.Picture, 0, Picture1.ScaleHeight - I, Picture1.ScaleWidth, I, 0, 0, Picture1.ScaleWidth, I, drawcode
Next I
End Sub

Private Sub mnuStretchBottom_Click()
Dim x As Integer

    If ClearDestination.Value Then Picture2.Picture = LoadPicture()
    For x = 1 To Picture1.ScaleHeight Step DELAY_SECS
        Picture2.PaintPicture Picture1.Picture, 0, 0, _
        Picture1.ScaleWidth, Picture1.ScaleHeight, 0, 0, _
        Picture1.ScaleWidth, x, drawcode
    Next
End Sub

Private Sub mnuStretchLeft_Click()
Dim x

    If ClearDestination.Value Then Picture2.Picture = LoadPicture()
    For x = 1 To Picture1.ScaleWidth Step DELAY_SECS
        Picture2.PaintPicture Picture1.Picture, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture2.ScaleWidth - x, 0, x, Picture1.ScaleHeight, drawcode
    Next
End Sub

Private Sub mnuStretchRight_Click()
Dim x

    If ClearDestination.Value Then Picture2.Picture = LoadPicture()
    For x = 1 To Picture1.ScaleWidth Step DELAY_SECS
        Picture2.PaintPicture Picture1.Picture, 0, 0, _
        Picture1.ScaleWidth, Picture1.ScaleHeight, 0, 0, x, _
        Picture1.ScaleHeight, drawcode
    Next

End Sub

Private Sub mnuStretchTop_Click()
    If ClearDestination.Value Then Picture2.Picture = LoadPicture()
    For x = 1 To Picture1.ScaleHeight Step DELAY_SECS
        Picture2.PaintPicture Picture1.Picture, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, 0, Picture1.ScaleHeight - x, Picture1.ScaleWidth, x, drawcode
    Next
End Sub

Private Sub mnuVerticalBlinds_Click()
Dim Stripes As Integer
Dim I As Integer, j As Integer
Dim StripeHeight As Integer

    Picture2.Cls
    StripeWidth = 20
    Stripes = Picture1.ScaleWidth / StripeWidth
    On Error Resume Next
    For j = 1 To StripeWidth
        For I = 0 To Stripes
            Picture2.PaintPicture Picture1.Picture, I * StripeWidth, 0, _
                j, Picture1.ScaleHeight, _
                I * StripeWidth, 0, _
                j, Picture1.ScaleHeight, drawcode
        Next
    Next
End Sub

Private Sub mnuVerticalStripes_Click()
On Error Resume Next

Dim ImageWidth As Integer
Dim imageHeight As Integer
Dim EvenStripes_Stpoint(25) As Integer
Dim OddStripes_Stpoint(25) As Integer
Dim EvenStripes_Endpoint(25) As Integer
Dim OddStripes_Endpoint(25) As Integer
Dim I As Integer, valx As Integer, j As Integer
Dim StripeWidth As Integer
Dim TotalNoOfStripes As Integer 'max 50
Picture2.BorderStyle = 1
Dim tm As String

tm = Str(Time) 'full time
tm = Right(tm, 5) ' returns SS AM/PM
tm = Left(tm, 2) 'return SS
TotalNoOfStripes = val(tm) ' to val

If TotalNoOfStripes > 50 Then
    TotalNoOfStripes = 50
End If
If TotalNoOfStripes < 12 Then
    TotalNoOfStripes = 12
End If
Me.Caption = "Stripes=" & TotalNoOfStripes


'Picture2.Cls
ImageWidth = Picture1.ScaleWidth
imageHeight = Picture1.ScaleHeight
StripeWidth = (ImageWidth / TotalNoOfStripes) + 1

valx = 0

EvenStripes_Stpoint(0) = 0
'For i = 1 To 4
For I = 1 To (TotalNoOfStripes - 1) / 2 '2)/2
    valx = I * (2 * StripeWidth) + 1
    EvenStripes_Stpoint(I) = valx
Next I

For I = 1 To TotalNoOfStripes / 2
    EvenStripes_Endpoint(I - 1) = EvenStripes_Stpoint(I - 1) + StripeWidth
Next I

For I = 1 To TotalNoOfStripes / 2
    OddStripes_Stpoint(I - 1) = EvenStripes_Endpoint(I - 1)
Next I

For I = 1 To TotalNoOfStripes / 2
    OddStripes_Endpoint(I - 1) = OddStripes_Stpoint(I - 1) + StripeWidth
Next I
I = 0

Dim s As Integer 'step
s = 0 'default step

If Picture1.ScaleWidth Mod 3 = 0 Then
s = 3
End If

If Picture1.ScaleWidth Mod 5 = 0 Then
s = 5
End If

For I = 1 To Picture1.ScaleHeight + 1
    For j = 0 To TotalNoOfStripes / 2 - 1 'for even stripes 4
        y = EvenStripes_Stpoint(j)
        Picture2.PaintPicture Picture1.Picture, EvenStripes_Stpoint(j), 0, StripeWidth, I, EvenStripes_Stpoint(j), 0, StripeWidth, I, drawcode
    Next j
    
    For w = 0 To TotalNoOfStripes / 2 - 1 'for odd stripes -4
        y = OddStripes_Stpoint(w)
        Picture2.PaintPicture Picture1.Picture, OddStripes_Stpoint(w), Picture1.ScaleHeight - I, StripeWidth, I, OddStripes_Stpoint(w), Picture2.ScaleHeight - I, StripeWidth, I, drawcode
     Next w
Next I
Picture2.BorderStyle = 0
Picture2.Picture = Picture1.Picture

End Sub

Private Sub mnuWipe4Corner_Click()
'mid point
Dim midx As Integer, midy As Integer

'top left corner (tlc)
Dim tlcx As Integer, tlcy As Integer
Dim btlcx As Boolean, btlcy As Boolean
Dim tlcWidth As Integer, tlcHeight As Integer

'top right corner(trc)
Dim trcx As Integer, trcy As Integer
Dim btrcx As Boolean, btrcy As Boolean
Dim trcWidth As Integer, trcHeight As Integer

'bottom right corner(brc)
Dim brcx As Integer, brcy As Integer
Dim bbrcx As Boolean, bbtrcy As Boolean
Dim brcWidth As Integer, brcHeight As Integer

'bottom left corner(blc)
Dim blcx As Integer, blcy As Integer
Dim bblcx As Boolean, bblcy As Boolean
Dim blcWidth As Integer, blcHeight As Integer

'other variables
Dim stepX As Integer, stepy As Integer
Dim indexHeight As Integer, indexWidth As Integer

'mid points
midx = Picture1.ScaleWidth / 2
midy = Picture1.ScaleHeight / 2
stepX = 1
stepy = 1
'x:y growth ratio
If midx > midy Then 'width > height. width grow faster
    stepX = Int((midx / midy) * 10)
    stepy = 10
End If
If midy > midx Then
    stepy = Int((midy / midx) * 10)
    stepX = 10
End If
If midy = midx Then
    stepX = 10
    stepy = 10
End If

'var initialization
tlcx = 0
tlcy = 0
btlcx = False
btlcy = False

trcx = Picture1.ScaleWidth
trcy = 0
btrcx = False
btrcy = False

blcx = 0
blcy = Picture1.ScaleHeight
bblcx = False
bblcy = False

brcx = Picture1.ScaleWidth
brcy = Picture1.ScaleHeight
bbrcx = False
bbrcy = False

indexHeight = 1
indexWidth = 1
'prog starts here!
If ClearDestination.Value Then Picture2.Picture = LoadPicture()
While btlcx = False Or btlcy = False Or btrcx = False Or btrcy = False Or bblcx = False Or bblcy = False Or bbrcx = False Or bbrcy = False
'tlc - x & Y constant, h & W inc
tlcHeight = indexHeight * stepy
    If tlcHeight > midy Then
        tlcHeight = midy
        btlcy = True
    End If
tlcWidth = indexWidth * stepX
    If tlcWidth > midx Then
        tlcWidth = midx
        btlcx = True
    End If
Picture2.PaintPicture Picture1.Picture, tlcx, tlcy, tlcWidth, tlcHeight, tlcx, tlcy, tlcWidth, tlcHeight, drawcode

'trc : x reduce, y const, width & height inc
trcWidth = indexWidth * stepX
trcx = Picture2.ScaleWidth - trcWidth
    If trcx <= midx Then
        trcx = midx
        btrcx = True
    End If
trcHeight = indexHeight * stepy
If trcHeight >= midy Then
    trcHeight = midy
    btrcy = True
End If
Picture2.PaintPicture Picture1.Picture, trcx, trcy, trcWidth, trcHeight, trcx, trcy, trcWidth, trcHeight, drawcode
    
'blc (X same, y decrease, h&w=inc
blcWidth = indexWidth * stepX
If blcWidth >= midx Then
    blcWidth = midx
    bblcx = True
End If
blcHeight = indexHeight * stepy
blcy = Picture2.ScaleHeight - blcHeight
If blcy <= midy Then
    blcy = midy
    bblcy = True
End If
Picture2.PaintPicture Picture1.Picture, blcx, blcy, blcWidth, blcHeight, blcx, blcy, blcWidth, blcHeight, drawcode

'brc - x dec, y dec, h&w inc
brcWidth = indexWidth * stepX
brcx = Picture2.ScaleWidth - brcWidth
If brcx <= midx Then
    brcx = midx
    bbrcx = True
End If
brcHeight = indexHeight * stepy
brcy = Picture2.ScaleHeight - brcHeight
If brcy <= midy Then
    brcy = midy
    bbrcy = True
End If
Picture2.PaintPicture Picture1.Picture, brcx, brcy, brcWidth, brcHeight, brcx, brcy, brcWidth, brcHeight, drawcode

indexHeight = indexHeight + 1
indexWidth = indexWidth + 1
'uncomment to induce delay
delayNext 0.1
Wend

End Sub

Private Sub mnuWipeBLCorner_Click()
On Error Resume Next
If ClearDestination.Value Then Picture2.Picture = LoadPicture()
'blc
'top right corner(blc)
Dim blcx As Integer, blcy As Integer
Dim bblcx As Boolean, bblcy As Boolean
Dim blcWidth As Integer, blcHeight As Integer

'other variables
Dim stepX As Integer, stepy As Integer
Dim indexHeight As Integer, indexWidth As Integer

'x:y growth ratio
If Picture2.ScaleWidth > Picture2.ScaleHeight Then   'width > height. width grow faster
    stepX = Int((Picture2.ScaleWidth / Picture2.ScaleHeight) * 10)
    stepy = 10
End If
If Picture2.ScaleHeight > Picture2.ScaleWidth Then
    stepy = Int((Picture2.ScaleHeight / Picture2.ScaleWidth) * 10)
    stepX = 10
End If
If Picture2.ScaleHeight = Picture2.ScaleWidth Then
    stepX = 10
    stepy = 10
End If

'var initialization
blcx = 0
blcy = Picture2.ScaleHeight
bblcx = False
bblcy = False

indexHeight = 1
indexWidth = 1
'prog starts here!
While bblcx = False Or bblcy = False
'blc : x reduce, y const, width & height inc
blcWidth = indexWidth * stepX
    If blcWidth >= Picture2.ScaleWidth Then
        blcWidth = Picture2.ScaleWidth
        bblcx = True
    End If
blcHeight = indexHeight * stepy
blcy = Picture2.ScaleHeight - blcHeight
    If blcy <= 0 Then
        blcHeight = Picture2.ScaleHeight
        bblcy = True
    End If
Picture2.PaintPicture Picture1.Picture, blcx, blcy, blcWidth, blcHeight, blcx, blcy, blcWidth, blcHeight, drawcode
indexHeight = indexHeight + 1
indexWidth = indexWidth + 1
delayNext 0.1
Wend

End Sub

Private Sub mnuWipeBottom_Click()
On Error Resume Next
Dim x As Integer
Dim stepX As Integer
stepX = DELAY_SECS

    If ClearDestination.Value Then Picture2.Picture = LoadPicture()
    For x = 1 To Picture1.ScaleHeight Step stepX
        Picture2.PaintPicture Picture1.Picture, 0, Picture1.ScaleHeight - x, Picture1.ScaleWidth, x, 0, Picture2.ScaleHeight - x, Picture2.ScaleWidth, x, drawcode
    Next

End Sub

Private Sub mnuWipeBRCorner_Click()
On Error Resume Next
'blc
'top right corner(blc)
Dim brcx As Integer, brcy As Integer
Dim bbrcx As Boolean, bbrcy As Boolean
Dim brcWidth As Integer, brcHeight As Integer
If ClearDestination.Value Then Picture2.Picture = LoadPicture()
'other variables
Dim stepX As Integer, stepy As Integer
Dim indexHeight As Integer, indexWidth As Integer

'x:y growth ratio
If Picture2.ScaleWidth > Picture2.ScaleHeight Then   'width > height. width grow faster
    stepX = Int((Picture2.ScaleWidth / Picture2.ScaleHeight) * 10)
    stepy = 10
End If
If Picture2.ScaleHeight > Picture2.ScaleWidth Then
    stepy = Int((Picture2.ScaleHeight / Picture2.ScaleWidth) * 10)
    stepX = 10
End If
If Picture2.ScaleHeight = Picture2.ScaleWidth Then
    stepX = 10
    stepy = 10
End If

'var initialization
brcx = Picture2.ScaleWidth
brcy = Picture2.ScaleHeight
bbrcx = False
bbrcy = False

indexHeight = 1
indexWidth = 1
'prog starts here!
While bbrcx = False Or bbrcy = False
'blc : x reduce, y reduce, width & height inc
brcWidth = indexWidth * stepX
brcx = Picture2.ScaleWidth - brcWidth
    If brcx <= 0 Then
       brcx = 0
       bbrcx = True
    End If
brcHeight = indexHeight * stepy
brcy = Picture2.ScaleHeight - brcHeight
    If blcy <= 0 Then
        brcHeight = Picture2.ScaleHeight
        bbrcy = True
    End If
Picture2.PaintPicture Picture1.Picture, brcx, brcy, brcWidth, brcHeight, brcx, brcy, brcWidth, brcHeight, drawcode
indexHeight = indexHeight + 1
indexWidth = indexWidth + 1
delayNext 0.1
Wend

End Sub

Private Sub mnuWipeCentre_Click()
Dim PWidth As Integer, PHeight As Integer
Dim I As Integer

    If ClearDestination.Value Then Picture2.Picture = LoadPicture()
    If Picture1.ScaleWidth > Picture1.ScaleHeight Then
        PWidth = Picture1.ScaleWidth - Picture1.ScaleHeight
        PHeight = 1
    ElseIf Picture1.ScaleWidth < Picture1.ScaleHeight Then
        PWidth = 1
        PHeight = Picture1.ScaleHeight - Picture1.ScaleWidth
    Else
        PWidth = 1
        PHeight = 1
    End If

    For I = 1 To Picture1.ScaleWidth - PWidth
        Picture2.PaintPicture Picture1.Picture, _
        Int((Picture1.ScaleWidth - PWidth) / 2), Int((Picture1.ScaleHeight - PHeight) / 2), _
        PWidth, PHeight, _
        Int((Picture1.ScaleWidth - PWidth) / 2), Int((Picture1.ScaleHeight - PHeight) / 2), _
        PWidth, PHeight, drawcode
        PWidth = PWidth + 1
        PHeight = height + 1
    Next

End Sub

Private Sub mnuWipeFade_Click()
On Error Resume Next
Dim x As Integer
Dim r As Integer
Dim c As Integer
Dim steps As Integer
steps = FADE_STEPS
 For x = steps To 3 Step -1
    For r = x To Picture1.ScaleHeight Step x
        For c = x To Picture1.ScaleWidth Step x
        Picture2.PaintPicture Picture1.Picture, c - 1, r - 1, 3, 3, c - 1, r - 1, 3, 3, drawcode
'        Picture2.PaintPicture Picture1.Picture, c - 1, r, 1, 1, c - 1, r, 1, 1, drawcode
'        Picture2.PaintPicture Picture1.Picture, c + 1, r, 1, 1, c + 1, r, 1, 1, drawcode
'        Picture2.PaintPicture Picture1.Picture, c - 1, r - 1, 1, 1, c - 1, r - 1, 1, 1, drawcode
'        Picture2.PaintPicture Picture1.Picture, c + 1, r + 1, 1, 1, c + 1, r + 1, 1, 1, drawcode
        Next c
    Next r
Next x
Picture2.Picture = Picture1.Picture
End Sub

Private Sub mnuWipeLeft_Click()
On Error Resume Next
Dim x As Integer
Dim stepX As Integer
stepX = DELAY_SECS
    If ClearDestination.Value Then Picture2.Picture = LoadPicture()
    For x = 1 To Picture1.ScaleWidth Step stepX
        Picture2.PaintPicture Picture1.Picture, 0, 0, _
        x, Picture1.ScaleHeight, 0, 0, x, _
        Picture1.ScaleHeight, drawcode
    Next

End Sub

Private Sub mnuWipeRight_Click()
On Error Resume Next
Dim x As Integer
Dim stepX As Integer
stepX = DELAY_SECS


    If ClearDestination.Value Then Picture2.Picture = LoadPicture()
    For x = 1 To Picture1.ScaleWidth Step stepX
        Picture2.PaintPicture Picture1.Picture, _
        Picture1.ScaleWidth - x, 0, _
        x, Picture1.ScaleHeight, _
        Picture1.ScaleWidth - x, 0, _
        x, Picture1.ScaleHeight, drawcode
    Next

End Sub

Private Sub mnuwipeRightLeft_Click()
Dim PWidth As Integer, PHeight As Integer
Dim I As Integer

    If ClearDestination.Value Then Picture2.Picture = LoadPicture()
    PWidth = 1
    PHeight = Picture1.ScaleHeight
    For I = 1 To Picture1.ScaleWidth / 2
        Picture2.PaintPicture Picture1.Picture, _
        (Picture1.ScaleWidth - PWidth) / 2, 0, _
        PWidth, PHeight, _
        (Picture1.ScaleWidth - PWidth) / 2, 0, _
        PWidth, PHeight, drawcode
        PWidth = PWidth + 2
    Next

End Sub

Private Sub mnuWipeTLCorner_Click()
If ClearDestination.Value Then Picture2.Picture = LoadPicture()
If ClearDestination.Value Then Picture2.Picture = LoadPicture()
'blc
'top right corner(blc)
Dim tlcx As Integer, tlcy As Integer
Dim bblcx As Boolean, bblcy As Boolean
Dim tlcWidth As Integer, tlcHeight As Integer

'other variables
Dim stepX As Integer, stepy As Integer
Dim indexHeight As Integer, indexWidth As Integer

'x:y growth ratio
If Picture2.ScaleWidth > Picture2.ScaleHeight Then   'width > height. width grow faster
    stepX = Int((Picture2.ScaleWidth / Picture2.ScaleHeight) * 10)
    stepy = 10
End If
If Picture2.ScaleHeight > Picture2.ScaleWidth Then
    stepy = Int((Picture2.ScaleHeight / Picture2.ScaleWidth) * 10)
    stepX = 10
End If
If Picture2.ScaleHeight = Picture2.ScaleWidth Then
    stepX = 10
    stepy = 10
End If

'var initialization
tlcx = 0
tlcy = 0
btlcx = False
btlcy = False

indexHeight = 1
indexWidth = 1
'prog starts here!
While btlcx = False Or btlcy = False
'blc : x reduce, y const, width & height inc
tlcWidth = indexWidth * stepX
    If tlcWidth >= Picture2.ScaleWidth Then
        tlcWidth = Picture2.ScaleWidth
        btlcx = True
    End If
tlcHeight = indexHeight * stepy
    If tlcHeight >= Picture2.ScaleHeight Then
        tlcHeight = Picture2.ScaleHeight
        btlcy = True
    End If
Picture2.PaintPicture Picture1.Picture, tlcx, tlcy, tlcWidth, tlcHeight, tlcx, tlcy, tlcWidth, tlcHeight, drawcode
indexHeight = indexHeight + 1
indexWidth = indexWidth + 1
delayNext 0.1
Wend
End Sub

Private Sub mnuWipeTop_Click()
Dim x As Integer
Dim stepX As Integer
stepX = DELAY_SECS

    If ClearDestination.Value Then Picture2.Picture = LoadPicture()
    For x = 1 To Picture1.ScaleHeight Step stepX
        Picture2.PaintPicture Picture1.Picture, 0, 0, Picture2.ScaleWidth, x, 0, 0, Picture2.ScaleWidth, x, drawcode
    Next

End Sub

Private Sub mnuWipeTRCorner_Click()
On Error Resume Next
If ClearDestination.Value Then Picture2.Picture = LoadPicture()
'top right corner(trc)
Dim trcx As Integer, trcy As Integer
Dim btrcx As Boolean, btrcy As Boolean
Dim trcWidth As Integer, trcHeight As Integer

'other variables
Dim stepX As Integer, stepy As Integer
Dim indexHeight As Integer, indexWidth As Integer

'x:y growth ratio
If Picture2.ScaleWidth > Picture2.ScaleHeight Then   'width > height. width grow faster
    stepX = Int((Picture2.ScaleWidth / Picture2.ScaleHeight) * 10)
    stepy = 10
End If
If Picture2.ScaleHeight > Picture2.ScaleWidth Then
    stepy = Int((Picture2.ScaleHeight / Picture2.ScaleWidth) * 10)
    stepX = 10
End If
If Picture2.ScaleHeight = Picture2.ScaleWidth Then
    stepX = 10
    stepy = 10
End If

'var initialization
trcx = Picture1.ScaleWidth
trcy = 0
btrcx = False
btrcy = False

indexHeight = 1
indexWidth = 1
'prog starts here!
While btrcx = False Or btrcy = False
'trc : x reduce, y const, width & height inc
trcWidth = indexWidth * stepX
trcx = Picture2.ScaleWidth - trcWidth
    If trcx <= 0 Then
        trcx = endx
        btrcx = True
    End If
trcHeight = indexHeight * stepy
If trcHeight >= Picture2.ScaleHeight Then
    trcHeight = endy
    btrcy = True
End If
Picture2.PaintPicture Picture1.Picture, trcx, trcy, trcWidth, trcHeight, trcx, trcy, trcWidth, trcHeight, drawcode
indexHeight = indexHeight + 1
indexWidth = indexWidth + 1
delayNext 0.1
Wend
End Sub

Private Sub mnuWipeUpDown_Click()
Dim PWidth As Integer, PHeight As Integer
Dim I As Integer

    If ClearDestination.Value Then Picture2.Picture = LoadPicture()
    PWidth = Picture1.ScaleWidth
    PHeight = 1
    For I = 1 To Picture1.ScaleHeight / 2
        Picture2.PaintPicture Picture1.Picture, _
        0, (Picture1.ScaleHeight - PHeight) / 2, _
        PWidth, PHeight, _
        0, (Picture1.ScaleHeight - PHeight) / 2, _
        PWidth, PHeight, drawcode
        PHeight = PHeight + 2
    Next

End Sub

Private Sub Picture2_Click()
'Picture2.Picture = Picture1.Picture
StopImmidiately = True
'End
End Sub


Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button <> 2 Then Exit Sub
    CommonDialog1.Filter = "Images|*.BMP;*.GIF;*.JPG"
    CommonDialog1.Action = 1
    If CommonDialog1.FileName = "" Then Exit Sub
    Picture1.Picture = LoadPicture(CommonDialog1.FileName)
    Picture2.width = Picture1.width
    Picture2.height = Picture1.height
'    Picture1.Visible = False
'    Picture2.Top = 0
'    Picture2.Left = 0
'    Picture2.Visible = True
    
End Sub



Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    Picture2.Visible = False
    Picture1.Visible = True
End Sub

Private Sub WipeLeft_Click()
On Error Resume Next
Dim x As Integer

    If ClearDestination.Value Then Picture2.Picture = LoadPicture()
    For x = 1 To Picture1.ScaleWidth
        Picture2.PaintPicture Picture1.Picture, 0, 0, _
        x, Picture1.ScaleHeight, 0, 0, x, _
        Picture1.ScaleHeight, drawcode
    Next

End Sub

Public Sub delayNext(x As Single)
t1 = Timer
While Timer - t1 < x
Wend
End Sub

Public Sub DrawBlocks(blockno As Integer, rc_wh As Integer, blockWidth As Integer, blockHeight As Integer)
On Error Resume Next
Dim Xc As Integer, Yc As Integer
Yc = 0
Xc = 0

Yc = YCoords(blockno / rc_wh)
Xc = XCoords(blockno Mod rc_wh)

Picture2.PaintPicture Picture1.Picture, Xc, Yc, blockWidth, blockHeight, Xc, Yc, blockWidth, blockHeight, drawcode
End Sub


Public Sub DrawMazeBlock(row As Integer, col As Integer, width As Integer, height As Integer)
    Picture2.PaintPicture Picture1.Picture, XCords(row, col), YCords(row, col), width, height, XCords(row, col), YCords(row, col), width, height, drawcode
End Sub

Public Sub ShowTransit(picname As String, Transitno As Integer, clr As Boolean, Hold As Integer)
Picture1.Picture = LoadPicture(picname)
Picture2.height = Picture1.height
Picture2.width = Picture1.width
Picture2.top = (WipesForm.height - Picture2.height) / 2

If Picture2.top < WipesForm.ScaleTop Then
    Picture2.top = WipesForm.ScaleTop
End If
Picture2.Left = (WipesForm.width - Picture2.width) / 2
If Picture2.Left < WipesForm.ScaleLeft Then
    Picture2.Left = WipesForm.ScaleLeft
End If

If Picture2.height > WipesForm.ScaleHeight Then
    Picture2.height = WipesForm.ScaleHeight
End If
If Picture2.width > WipesForm.ScaleWidth Then
    Picture2.height = WipesForm.ScaleWidth
End If
If Form1.chkQuickClear.Value = 1 Then
Picture2.Cls
Picture2.Picture = LoadPicture()
End If

DoEvents
delayNext 1
DoEvents
Me.Refresh
DoEvents
Select Case Transitno
    Case 1
        drawcode = vbSrcCopy
        mnuWipeLeft_Click
        delayNext (Hold)
        If clr = True Then
            drawcode = vbBlackness
            mnuWipeLeft_Click
        End If
        
     Case 2
     drawcode = vbSrcCopy
        mnuWipeRight_Click
        delayNext (Hold)
        If clr = True Then
            drawcode = vbBlackness
            mnuWipeRight_Click
        End If

    Case 3
    drawcode = vbSrcCopy
        mnuWipeTop_Click
        delayNext (Hold)
        If clr = True Then
            drawcode = vbBlackness
            mnuWipeTop_Click
        End If
        
    Case 4
    drawcode = vbSrcCopy
        mnuWipeBottom_Click
        delayNext (Hold)
        If clr = True Then
            drawcode = vbBlackness
            mnuWipeBottom_Click
        End If
        
    Case 5
    drawcode = vbSrcCopy
        mnuWipeTLCorner_Click
        delayNext (Hold)
        If clr = True Then
            drawcode = vbBlackness
            mnuWipeTLCorner_Click
        End If
        
    Case 6
 drawcode = vbSrcCopy
        mnuWipeTRCorner_Click
        delayNext (Hold)
        If clr = True Then
            drawcode = vbBlackness
            mnuWipeTRCorner_Click
        End If
    
    Case 7
 drawcode = vbSrcCopy
        mnuWipeBLCorner_Click
        delayNext (Hold)
        If clr = True Then
            drawcode = vbBlackness
            mnuWipeBLCorner_Click
        End If
        
    Case 8
 drawcode = vbSrcCopy
        mnuWipeBRCorner_Click
        delayNext (Hold)
        If clr = True Then
            drawcode = vbBlackness
            mnuWipeBRCorner_Click
        End If
        
    Case 9
 drawcode = vbSrcCopy
        mnuwipeRightLeft_Click
        delayNext (Hold)
        If clr = True Then
            drawcode = vbBlackness
            mnuwipeRightLeft_Click
        End If
        
    Case 10
 drawcode = vbSrcCopy
        mnuWipeUpDown_Click
        delayNext (Hold)
        If clr = True Then
            drawcode = vbBlackness
            mnuWipeUpDown_Click
        End If
        
    Case 11
 drawcode = vbSrcCopy
        mnuWipeCentre_Click
        delayNext (Hold)
        If clr = True Then
            drawcode = vbBlackness
            mnuWipeCentre_Click
        End If
        
    Case 12
 drawcode = vbSrcCopy
        mnuRandomBlocks_Click
        delayNext (Hold)
        If clr = True Then
            drawcode = vbBlackness
            mnuRandomBlocks_Click
        End If
        
    Case 13
 drawcode = vbSrcCopy
        mnuStretchRight_Click
        delayNext (Hold)
       
    Case 14
 drawcode = vbSrcCopy
        mnuStretchLeft_Click
        delayNext (Hold)
        
    Case 15
 drawcode = vbSrcCopy
        mnuStretchBottom_Click
        delayNext (Hold)
       
    Case 16
 drawcode = vbSrcCopy
        mnuStretchTop_Click
        delayNext (Hold)
        
    Case 17
 drawcode = vbSrcCopy
        mnuHorizontalStripes_Click
        delayNext (Hold)
        If clr = True Then
            drawcode = vbBlackness
            mnuHorizontalStripes_Click
        End If
        
    Case 18
 drawcode = vbSrcCopy
        mnuVerticalStripes_Click
        delayNext (Hold)
        If clr = True Then
            drawcode = vbBlackness
            mnuVerticalStripes_Click
        End If
        
    Case 19
 drawcode = vbSrcCopy
        mnuHorizontalBlinds_Click
        delayNext (Hold)
        If clr = True Then
            drawcode = vbBlackness
            mnuHorizontalBlinds_Click
        End If
        
    Case 20
 drawcode = vbSrcCopy
        mnuVerticalBlinds_Click
        delayNext (Hold)
        If clr = True Then
            drawcode = vbBlackness
            mnuVerticalBlinds_Click
        End If
        
    Case 21
 drawcode = vbSrcCopy
        mnuCircularWipe_Click
        delayNext (Hold)
        
    Case 22
 drawcode = vbSrcCopy
        mnuWipeFade_Click
        delayNext (Hold)
        If clr = True Then
            drawcode = vbBlackness
            mnuWipeFade_Click
        End If
        
    Case 23
 drawcode = vbSrcCopy
        mnuGrowCentre_Click
        delayNext (Hold)
        
    Case 24
 drawcode = vbSrcCopy
        mnuS_Click_Click
        delayNext (Hold)
        If clr = True Then
            drawcode = vbBlackness
            mnuS_Click_Click
        End If
        
    Case 25
 drawcode = vbSrcCopy
        mnuReverseS_Click
        delayNext (Hold)
        If clr = True Then
            drawcode = vbBlackness
            mnuReverseS_Click
        End If
        
    Case 26
 drawcode = vbSrcCopy
        mnuSlideLeft_Click
        delayNext (Hold)
        If clr = True Then
            drawcode = vbBlackness
            mnuSlideLeft_Click
        End If
        
    Case 27
 drawcode = vbSrcCopy
        mnuSlideRight_Click
        delayNext (Hold)
        If clr = True Then
            drawcode = vbBlackness
            mnuSlideRight_Click
        End If
        
    Case 28
 drawcode = vbSrcCopy
        mnuSlideTop_Click
        delayNext (Hold)
        If clr = True Then
            drawcode = vbBlackness
            mnuSlideTop_Click
        End If
        
    Case 29
 drawcode = vbSrcCopy
        mnuSlideBottom_Click
        delayNext (Hold)
        If clr = True Then
            drawcode = vbBlackness
            mnuSlideBottom_Click
        End If
        
    Case 30
 drawcode = vbSrcCopy
        mnuMaze_Click
        delayNext (Hold)
        If clr = True Then
            drawcode = vbBlackness
            mnuMaze_Click
        End If
        
Case 31
drawcode = vbSrcCopy
    mnuWipe4Corner_Click
    delayNext (Hold)
        If clr = True Then
            drawcode = vbBlackness
            mnuWipe4Corner_Click
        End If
    
    End Select
    'If MsgBox("Done! Continue?", vbYesNo, "Continue?") = vbNo Then End
End Sub


Public Sub ShowTransit2(picname As String, Transitno As Integer, clr As Boolean, Hold As Integer, vbDrawMode As Long)
Picture1.Picture = LoadPicture(picname)
Picture2.height = Picture1.height
Picture2.width = Picture1.width
Picture2.top = (WipesForm.height - Picture2.height) / 2

If Picture2.top < WipesForm.ScaleTop Then
    Picture2.top = WipesForm.ScaleTop
End If
Picture2.Left = (WipesForm.width - Picture2.width) / 2
If Picture2.Left < WipesForm.ScaleLeft Then
    Picture2.Left = WipesForm.ScaleLeft
End If

If Picture2.height > WipesForm.ScaleHeight Then
    Picture2.height = WipesForm.ScaleHeight
End If
If Picture2.width > WipesForm.ScaleWidth Then
    Picture2.height = WipesForm.ScaleWidth
End If

DoEvents
delayNext 1
DoEvents
Me.Refresh
DoEvents
Select Case Transitno
    Case 1
        drawcode = vbDrawMode
        mnuWipeLeft_Click
        delayNext (Hold)
                
     Case 2
     drawcode = vbDrawMode
        mnuWipeRight_Click
        delayNext (Hold)
    
    Case 3
    drawcode = vbDrawMode
        mnuWipeTop_Click
        delayNext (Hold)
        
    Case 4
    drawcode = vbDrawMode
        mnuWipeBottom_Click
        delayNext (Hold)
                
    Case 5
    drawcode = vbDrawMode
        mnuWipeTLCorner_Click
        delayNext (Hold)
        
        
    Case 6
 drawcode = vbDrawMode
        mnuWipeTRCorner_Click
        delayNext (Hold)
        
    
    Case 7
 drawcode = vbDrawMode
        mnuWipeBLCorner_Click
        delayNext (Hold)
        
        
    Case 8
 drawcode = vbDrawMode
        mnuWipeBRCorner_Click
        delayNext (Hold)
        
        
    Case 9
 drawcode = vbDrawMode
        mnuwipeRightLeft_Click
        delayNext (Hold)
        
        
    Case 10
 drawcode = vbDrawMode
        mnuWipeUpDown_Click
        delayNext (Hold)
        
        
    Case 11
 drawcode = vbDrawMode
        mnuWipeCentre_Click
        delayNext (Hold)
       
        
    Case 12
 drawcode = vbDrawMode
        mnuRandomBlocks_Click
        delayNext (Hold)
        
        
    Case 13
 drawcode = vbDrawMode
        mnuStretchRight_Click
        delayNext (Hold)
'
        
    Case 14
 drawcode = vbDrawMode
        mnuStretchLeft_Click
        delayNext (Hold)
'
        
    Case 15
 drawcode = vbDrawMode
        mnuStretchBottom_Click
        delayNext (Hold)
'
        
    Case 16
 drawcode = vbDrawMode
        mnuStretchTop_Click
        delayNext (Hold)
        
    Case 17
 drawcode = vbDrawMode
        mnuHorizontalStripes_Click
        delayNext (Hold)
       
        
    Case 18
 drawcode = vbDrawMode
        mnuVerticalStripes_Click
        delayNext (Hold)
        
        
    Case 19
 drawcode = vbDrawMode
        mnuHorizontalBlinds_Click
        delayNext (Hold)
        
    Case 20
 drawcode = vbDrawMode
        mnuVerticalBlinds_Click
        delayNext (Hold)
        
        
    Case 21
 drawcode = vbDrawMode
        mnuCircularWipe_Click
        delayNext (Hold)
'        If clr = True Then
'            drawcode = vbBlackness
'            mnuCircularWipe_Click
'        End If
        
    Case 22
 drawcode = vbDrawMode
        mnuWipeFade_Click
        delayNext (Hold)
        
        
    Case 23
 drawcode = vbDrawMode
        mnuGrowCentre_Click
        delayNext (Hold)
'        If clr = True Then
'            drawcode = vbBlackness
'            mnuGrowCentre_Click
'        End If
        
    Case 24
 drawcode = vbDrawMode
        mnuS_Click_Click
        delayNext (Hold)
        
        
    Case 25
 drawcode = vbDrawMode
        mnuReverseS_Click
        delayNext (Hold)
        
        
    Case 26
 drawcode = vbDrawMode
        mnuSlideLeft_Click
        delayNext (Hold)
        
        
    Case 27
 drawcode = vbDrawMode
        mnuSlideRight_Click
        delayNext (Hold)
        
        
    Case 28
 drawcode = vbDrawMode
        mnuSlideTop_Click
        delayNext (Hold)
        
        
    Case 29
 drawcode = vbDrawMode
        mnuSlideBottom_Click
        delayNext (Hold)
        
        
    Case 30
 drawcode = vbDrawMode
        mnuMaze_Click
        delayNext (Hold)
        
    Case 31
drawcode = vbSrcCopy
    mnuWipe4Corner_Click
    delayNext (Hold)
        
    End Select
    'If MsgBox("Done! Continue?", vbYesNo, "Continue?") = vbNo Then End
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "tlbNew"
    Call ScriptMaker.Show

        
Case "tlbOpen"
    ScriptMaker.Show


Case "tlbExecute"
    mnuLoadAndExecScript_Click
    

Case "tlbSetDirectory"
    mnuSetdir_Click
    
Case "tlbRunPreset"
    cmdStart_Click

Case "tlbPlayMusic"
    cmdoundeSelect_Click
End Select
End Sub
