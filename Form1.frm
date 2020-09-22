VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "XPButton Stuff"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7125
   DrawMode        =   1  'Blackness
   LinkTopic       =   "Form1"
   ScaleHeight     =   341
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   475
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4095
      Left            =   120
      ScaleHeight     =   269
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   453
      TabIndex        =   2
      Top             =   120
      Width           =   6855
   End
   Begin Project1.XPButton XPButton1 
      Cancel          =   -1  'True
      Height          =   615
      Left            =   6000
      TabIndex        =   1
      Top             =   4440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1085
      AccessKey       =   " "
      Caption         =   "ESC TO QUIT"
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
   End
   Begin Project1.XPButton Btn1 
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4560
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AccessKey       =   " "
      Caption         =   "Begin"
      BorderStyle     =   2
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      ForeColour      =   16384
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private frame As Byte
Private Const DT_WORDBREAK = &H10
Private Const DT_CALCRECT = &H400
Private Const DT_CENTER = &H1
Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private Sub Btn1_Click()
    Dim Str1    As String
    Dim vh      As Long 'vertical height of text (wrapped)
    Dim hRect   As RECT 'text boundaries
    Static Y1 As Long
    
    'make sure that the Next button wont react until the text has finished.
    If Y1 = 0 Then
        If frame = 0 Then
            Btn1.Caption = "&Next"
            Refresh
            Str1 = "XPButton is a Windows XP Style command button, now available for Win95/NT3.1 and above. XPButton doesn't just have the standard borderstyle, but additional styles aswell, to enhance the appearance of your applications. Many of the standard bugs you will find in custom made button controls have been ironed out, and the button behaves almost the same as the Windows XP buttons."
            GoTo DrawTextNowPlease
        ElseIf frame = 1 Then
            Refresh
            Str1 = "The buttons contain default and cancel properties, enabling the buttons to have default focus, or act as a cancel button. In addition AccessKeys have been incorporated into the button, so if your button has a caption such as ""&Hello"" and the user pressed the Alt+H combo, this will cause the code inside the Click event to be executed."
            GoTo DrawTextNowPlease
        ElseIf frame = 2 Then
            Refresh
            Str1 = "Pretty amazing, isn't it. The XP Button is also superior to other XP buttons I have tried because it uses fast drawing methods for the borders and background gradient. While some buttons can redraw at 30 times a second, the XPButton can redraw at 2000 times a second (on my Pentium III 866 with gradient background turned on) It can redraw at at almost 4000 times a second without a gradient background. This makes it more usable for slower computers."
            GoTo DrawTextNowPlease
        ElseIf frame = 3 Then
            Refresh
            Str1 = "XPButton doesn't support pictures yet, but it will in the future. I will also add properties for text allignment. XPButton has MouseEnter and MouseExit events so you can easily change the appearance of the button according to whether the mouse is inside the usercontrol or not."
            GoTo DrawTextNowPlease
        ElseIf frame = 4 Then
            Refresh
            Str1 = "You can compile the XPButton project into an easy to use OCX, or you can add the button into a project, like I have done here. Remember to add the win.tlb file to the References aswell. If you import the button into your project you will require no additional runtime files. (Not even the win.tlb file)"
            GoTo DrawTextNowPlease
        ElseIf frame = 5 Then
            Btn1.Caption = "&Finish"
            Refresh
            Str1 = "Isn't the scrolling text nice aswell. I didn't even have to use bitblt and an image. Just the DrawText api. Very simple, infact I just came up with it now and I like the idea. It works quite well and doesn't seem to flicker much."
            GoTo DrawTextNowPlease
        ElseIf frame = 6 Then
            Unload Me
        End If
    End If
    
    Exit Sub
    
DrawTextNowPlease:

    SetRect hRect, 4, 0, Pic1.ScaleWidth - 4, Pic1.ScaleHeight
    vh = DrawText(Pic1.hdc, Str1, -1, hRect, DT_CALCRECT Or DT_CENTER Or DT_WORDBREAK)
    
    For Y1 = -vh To (Pic1.ScaleHeight - vh) / 2
        Pic1.Cls
        'set the rectangular area that the text can drawn onto'
        SetRect hRect, 4, Y1, l, Pic1.ScaleHeight
        'do a test draw (not actuallt drawn on screen) and find the height
        'the text occupies with text wrapping.
        'set the rectangular area such that the text is drawn
        'horizontally and vertically centered on the form
        SetRect hRect, 4, Y1, Pic1.ScaleWidth - 4, Pic1.ScaleHeight
        'now draw the text
        DrawText Pic1.hdc, Str1, -1, hRect, DT_WORDBREAK Or DT_CENTER
        DoEvents
        Sleep 7
    Next
    
    Y1 = 0
    frame = frame + 1
End Sub

Private Sub Btn1_MouseEnter()
    Btn1.ForeColour = vbRed
    Btn1.FontBold = True
End Sub

Private Sub Btn1_MouseExit()
    Btn1.ForeColour = &H4000&
    Btn1.FontBold = False
End Sub

Private Sub XPButton1_Click()
    End
End Sub
