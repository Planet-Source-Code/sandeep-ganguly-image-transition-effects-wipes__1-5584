Attribute VB_Name = "Module11"
Public Const PLANES = 14 ' Number of planes
Public Const BITSPIXEL = 12 ' Number of bits per pixel


Type RECT
    Left As Long
    top As Long
    Right As Long
    Bottom As Long
End Type

Dim What As RECT


Declare Function CreateSolidBrush Lib "gdi32" _
    (ByVal crColor As Long) As Long


Declare Function DeleteObject Lib "gdi32" _
    (ByVal hObject As Long) As Long


Declare Function GetDeviceCaps Lib "gdi32" _
    (ByVal hDC As Long, ByVal nIndex As Long) As Long


Declare Function FillRect Lib "user32" _
    (ByVal hDC As Long, lpRect As RECT, _
    ByVal hBrush As Long) As Long
    
    Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Sub FadeForm(frmIn As Form, fadeStyle As Integer, RedVal As Integer, GreenVal As Integer)
    'fadeStyle = 0 produces diagonal gradient
    'fadeStyle = 1 produces vertical gradient
    'fadeStyle = 2 produces horizontal gradient
    'any other value produces solid medium-blue background
    Static ColorBits As Long
    Static RgnCnt As Integer
    Dim NbrPlanes As Long
    Dim BitsPerPixel As Long
    Dim AreaHeight As Long
    Dim AreaWidth As Long
    Dim BlueLevel As Long
    Dim prevScaleMode As Integer
    Dim IntervalY As Long
    Dim IntervalX As Long
    Dim I As Integer
    Dim r As Long
    Dim ColorVal As Long
    Dim FillArea As RECT
    Dim hBrush As Long
    'init code - performed only on the first pass through this routin
    '     e.


    If ColorBits = 0 Then
        
        'determine number of color bits supported.
        BitsPerPixel = GetDeviceCaps(frmIn.hDC, BITSPIXEL)
        NbrPlanes = GetDeviceCaps(frmIn.hDC, PLANES)
        ColorBits = (BitsPerPixel * NbrPlanes)
        'Calculate the number of regions that the screen will be divided
        '     o.
        'This is optimized for the current display's color depth. Why was
        '     te
        'time rendering 256 shades if you can only discern 32 or 64 of th
        '     em?


        Select Case ColorBits
            Case 32: RgnCnt = 256 '16M colors: 8 bits For blue
            Case 24: RgnCnt = 256 '16M colors: 8 bits For blue
            Case 16: RgnCnt = 256 '64K colors: 5 bits For blue
            Case 15: RgnCnt = 32 '32K colors: 5 bits For blue
            Case 8: RgnCnt = 64 '256 colors: 64 dithered blues
            Case 4: RgnCnt = 64 '16 colors : 64 dithered blues
            Case Else: ColorBits = 4
            RgnCnt = 64 '16 colors assumed: 64 dithered blues
        End Select


End If


'if solid then set and bail out


If fadeStyle = 3 Then
    frmIn.BackColor = &H7F0000 ' med blue
    Exit Sub
End If


prevScaleMode = frmIn.ScaleMode 'save the current scalemode
frmIn.ScaleMode = 3 'set to pixel
AreaHeight = frmIn.ScaleHeight 'calculate sizes
AreaWidth = frmIn.ScaleWidth
frmIn.ScaleMode = prevScaleMode 'reset to saved value

ColorVal = 256 / RgnCnt 'color diff between regions
IntervalY = AreaHeight / RgnCnt '# vert pixels per region
IntervalX = AreaWidth / RgnCnt '# horz pixels per region
'fill the client area from bottom/right
'to top/left except for top/left region
FillArea.Left = 0
FillArea.top = 0
FillArea.Right = AreaWidth
FillArea.Bottom = AreaHeight
BlueLevel = 0



For I = 0 To RgnCnt - 1
    'create a brush of the appropriate blue colour

 hBrush = CreateSolidBrush(RGB(RedVal, GreenVal, BlueLevel))
'hBrush = CreateSolidBrush(RGB(RedVal, GreenVal, BlueVal))
    
    If fadeStyle = 0 Then
    'diagonal gradient
    FillArea.top = FillArea.Bottom - IntervalY
    FillArea.Left = 0
    r = FillRect(frmIn.hDC, FillArea, hBrush)
    
    FillArea.top = 0
    FillArea.Left = FillArea.Right - IntervalX
    r = FillRect(frmIn.hDC, FillArea, hBrush)
    
    FillArea.Bottom = FillArea.Bottom - IntervalY
    FillArea.Right = FillArea.Right - IntervalX
    
ElseIf fadeStyle = 1 Then
    'horizontal gradient
    FillArea.top = FillArea.Bottom - IntervalY
    r = FillRect(frmIn.hDC, FillArea, hBrush)
    
    FillArea.Bottom = FillArea.Bottom - IntervalY
    
Else
    'vertical gradient
    FillArea.Left = FillArea.Right - IntervalX
    r = FillRect(frmIn.hDC, FillArea, hBrush)
    FillArea.Right = FillArea.Right - IntervalX
End If


'done with that brush, so delete
r = DeleteObject(hBrush)
'increment the value by the appropriate
'steps for the display colour depth
BlueLevel = BlueLevel + ColorVal

Next


'Fill any the remaining top/left holes of the
'client area with solid blue
FillArea.top = 0
FillArea.Left = 0

hBrush = CreateSolidBrush(RGB(0, 0, 255))
r = FillRect(frmIn.hDC, FillArea, hBrush)
r = DeleteObject(hBrush)

frmIn.Refresh
End Sub


Sub Gradient(TheObject As Object, RedVal&, GreenVal&, Blueval&, TopToBottom As Boolean)
    'TheObject can be any object that supports the Line method (like forms and pictures).
    'Redval, Greenval, and Blueval are the Red, Green, and Blue starting values from 0 to 255.
    'TopToBottom determines whether the gradient will draw down or up.
    Dim Step%, Reps%, FillTop%, FillLeft%, FillRight%, FillBottom%, HColor$
    'This will create 63 steps in the gradient. This looks smooth on 16-bit and 24-bit color.
    'You can change this, but be careful. You can do some strange-looking stuff with it...
    Step = (TheObject.height / 63)
    'This tells it whether to start on the top or the bottom and adjusts variables accordingly.
    If TopToBottom = True Then FillTop = 0 Else FillTop = TheObject.height - Step
    FillLeft = 0
    FillRight = TheObject.width
    FillBottom = FillTop + Step
    'If you changed the number of steps, change the number of reps to match it.
    'If you don't, the gradient will look all funny.
    For Reps = 1 To 63
        'This draws the colored bar.
        TheObject.Line (FillLeft, FillTop)-(FillRight, FillBottom), RGB(RedVal, GreenVal, Blueval), BF
        'This decreases the RGB values to darken the color.
        'Lower the value for "squished" gradients. Raise it for incomplete gradients.
        'Also, if you change the number of steps, you will need to change this number.
        RedVal = RedVal - 4
        GreenVal = GreenVal - 4
        Blueval = Blueval - 4
        'This prevents the RGB values from becoming negative, which causes a runtime error.
        If RedVal <= 0 Then RedVal = 0
        If GreenVal <= 0 Then GreenVal = 0
        If Blueval <= 0 Then Blueval = 0
        'More top or bottom stuff; Moves to next bar.
        If TopToBottom = True Then FillTop = FillBottom Else FillTop = FillTop - Step
        FillBottom = FillTop + Step
    Next
End Sub



