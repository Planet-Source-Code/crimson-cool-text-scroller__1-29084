VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAbout 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer6 
      Left            =   3960
      Top             =   2880
   End
   Begin VB.Timer Timer5 
      Left            =   3960
      Top             =   2880
   End
   Begin VB.Timer Timer4 
      Left            =   3960
      Top             =   2880
   End
   Begin VB.Timer Timer3 
      Left            =   4080
      Top             =   2880
   End
   Begin VB.Timer Timer2 
      Left            =   3960
      Top             =   2880
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393217
      BackColor       =   -2147483642
      Enabled         =   0   'False
      MousePointer    =   1
      Appearance      =   0
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3960
      Top             =   2880
   End
   Begin RichTextLib.RichTextBox rtb2 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393217
      BackColor       =   -2147483642
      Enabled         =   0   'False
      MousePointer    =   1
      Appearance      =   0
      TextRTF         =   $"Form1.frx":0092
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtb3 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   393217
      BackColor       =   -2147483642
      Enabled         =   0   'False
      MousePointer    =   1
      Appearance      =   0
      TextRTF         =   $"Form1.frx":0124
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtb4 
      Height          =   375
      Left            =   285
      TabIndex        =   3
      Top             =   1560
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   661
      _Version        =   393217
      BackColor       =   -2147483642
      Enabled         =   0   'False
      MousePointer    =   1
      Appearance      =   0
      TextRTF         =   $"Form1.frx":01B6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtb5 
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   2040
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393217
      BackColor       =   0
      Enabled         =   0   'False
      ReadOnly        =   -1  'True
      MousePointer    =   1
      Appearance      =   0
      TextRTF         =   $"Form1.frx":0248
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblWeb 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   865
      TabIndex        =   8
      Top             =   2065
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblExit 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4020
      TabIndex        =   7
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sub7 Text Scroller Effect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   30
      Width           =   4215
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   4215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   4
      Height          =   3015
      Left            =   0
      Top             =   0
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   255
      Left            =   1485
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblRestart 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Restart"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1485
      TabIndex        =   5
      Top             =   2520
      Width           =   1095
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'After seeing how many people wanted the source code for a text
'scroller like the one in Subseven Defcon Edition - don't know
'about other editions - I decided to download Subseven to see
'what the fuss was all about...

'I figured out how to make a scroller like it and thought I'd
'post the source on Planet Source Code for others to use.

'No need to give me credit if you use this in your program since
'it was easy to make, plus I didn't think of it myself, of course.

'Enjoy
 


'   *****      **          *****      **   **
'   ***** **** ** **    ** **    **** ***  **
'   *     *  * ** ***  *** ***** *  * **** **
'   *     ***  ** ******** ***** *  * *******
'   ***** *  * ** ** ** **    ** *  * **  ***
'   ***** *  * ** **    ** ***** **** **   **


'I gave a label the function to open a website since I didn't take the time to figure out how to make the RichTextBox do it.  If you figure it out then you just improved the program :)
'Add and modify text constants to your needs
'More than one RichTextBoxes are used to prevent flickering
Option Explicit
Const text1 = "Welcome to Sub7 Text Scroller"
Const text2 = "Written by CrImSoN"
Const text3 = "Thanks to Subseven crew for giving me the idea"
Const text4 = "Have fun using my code PSC programmers"
Const text5 = "http://www.addwebsitehere.com"
Dim i As Single, X As Single, Y As Single, z As Single, a As Single

Private Sub Form_Load()
'Set singles to start at 0
i = 0
X = 0
Y = 0
z = 0
a = 0
MakeTopMost Me 'Make this form be on top of other windows
MakeCenter Me 'Make this form load in the center of the screen
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Change colors of label back to normal when mouse cursor is moved across form
lblRestart.BackColor = &H0& 'Black
lblRestart.ForeColor = &HFFFFFF 'White
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then MoveWithoutCap Me 'If left mouse button is held down over label1, the form will move with the mouse
End Sub

Private Sub lblExit_Click()
Unload Me 'Exit this form
End Sub

Private Sub lblRestart_Click() 'Restarts the the scrolling effects
'Set singles back to 0 and RichTextBoxes' text property to nothing
i = 0
X = 0
Y = 0
z = 0
a = 0
Timer1.Interval = 0
Timer2.Interval = 0
Timer3.Interval = 0
Timer4.Interval = 0
Timer5.Interval = 0
Timer6.Interval = 0
rtb.Text = ""
rtb2.Text = ""
rtb3.Text = ""
rtb4.Text = ""
rtb5.Text = ""
lblWeb.Visible = False
rtb5.Visible = True
rtb5.SelLength = 100
rtb5.SelUnderline = False
Timer1.Interval = 100
End Sub

Private Sub rtf_Change()

End Sub

Private Sub lblRestart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Change colors of label when mouse cursor is moved over it...  just for cool looks
lblRestart.BackColor = &HFFFFFF 'White
lblRestart.ForeColor = &H0& 'Black
End Sub

Private Sub lblWeb_Click()
Shell "start " & lblWeb.Caption  ' Opens a browser to the website, if any, in the lblWeb.caption(which is the same as rtb5.text)
End Sub

Private Sub Timer1_Timer()
i = i + 1
rtb.TextRTF = Left(text1, i) & "_"
rtb.SelStart = 0
rtb.SelLength = i
rtb.SelColor = vbRed
rtb.SelStart = i - 1
rtb.SelLength = 1
rtb.SelColor = vbWhite
rtb.SelStart = i
rtb.SelLength = 1
rtb.SelColor = vbRed
rtb.SelStart = 0
If rtb.Text = text1 & "_" Then
    rtb.Text = Mid(rtb.Text, 1, Len(rtb.Text) - 1)
    Timer1.Interval = 0
    Timer2.Interval = 10
End If
End Sub

Private Sub Timer2_Timer()
X = X + 1
rtb2.TextRTF = Left(text2, X) & "_"
rtb2.SelStart = 0
rtb2.SelLength = X
rtb2.SelColor = vbRed
rtb2.SelStart = X - 1
rtb2.SelLength = 1
rtb2.SelColor = vbWhite
rtb2.SelStart = X
rtb2.SelLength = 1
rtb2.SelColor = vbRed
rtb2.SelStart = 0
If rtb2.Text = text2 & "_" Then
    rtb2.Text = Mid(rtb2.Text, 1, Len(rtb2.Text) - 1)
    Timer2.Interval = 0
    Timer3.Interval = 10
End If
End Sub

Private Sub Timer3_Timer()
Y = Y + 1
rtb3.TextRTF = Left(text3, Y) & "_"
rtb3.SelStart = 0
rtb3.SelLength = Y
rtb3.SelColor = vbWhite
rtb3.SelStart = Y - 1
rtb3.SelLength = 1
rtb3.SelColor = vbRed
rtb3.SelStart = Y
rtb3.SelLength = 1
rtb3.SelColor = vbWhite
rtb3.SelStart = 0
If rtb3.Text = text3 & "_" Then
    rtb3.Text = Mid(rtb3.Text, 1, Len(rtb3.Text) - 1)
    Timer3.Interval = 0
    Timer4.Interval = 10
End If
End Sub

Private Sub Timer4_Timer()
z = z + 1
rtb4.TextRTF = Left(text4, z) & "_"
rtb4.SelStart = 0
rtb4.SelLength = z
rtb4.SelColor = vbRed
rtb4.SelStart = z - 1
rtb4.SelLength = 1
rtb4.SelColor = vbWhite
rtb4.SelStart = z
rtb4.SelLength = 1
rtb4.SelColor = vbRed
rtb4.SelStart = 0
If rtb4.Text = text4 & "_" Then
    rtb4.Text = Mid(rtb4.Text, 1, Len(rtb4.Text) - 1)
    Timer5.Interval = 100
    Timer4.Interval = 0
End If
End Sub

Private Sub Timer5_Timer()
a = a + 1
rtb5.TextRTF = Left(text5, a) & "_"
rtb5.SelStart = 0
rtb5.SelLength = a
rtb5.SelColor = vbRed
rtb5.SelStart = a - 1
rtb5.SelLength = 1
rtb5.SelColor = vbWhite
rtb5.SelStart = a
rtb5.SelLength = 1
rtb5.SelColor = vbRed
rtb5.SelStart = 0
If rtb5.Text = text5 & "_" Then
    rtb5.Text = Mid(rtb5.Text, 1, Len(rtb5.Text) - 1)
    rtb5.SelLength = Len(rtb5.Text)
    rtb5.SelColor = vbBlue
    rtb5.SelUnderline = True
    rtb5.Visible = False
    lblWeb.ForeColor = rtb5.SelColor
    lblWeb.Caption = rtb5.Text
    lblWeb.FontUnderline = rtb5.SelUnderline
    lblWeb.Visible = True
    'Replace the 8 lines above this comment with the next two comments, erase the 's of course, for another cool effect
    'rtb5.sellength = 1
    'Timer6.Interval = 100
    Timer5.Interval = 0
End If
End Sub

Private Sub Timer6_Timer()
rtb5.SelColor = vbBlue
rtb5.SelUnderline = True
rtb5.SelLength = rtb5.SelLength + 1
If rtb5.SelLength >= Len(rtb5.Text) Then
    rtb5.SelColor = vbBlue
    rtb5.SelUnderline = True
    rtb5.Visible = False
    lblWeb.ForeColor = rtb5.SelColor
    lblWeb.Caption = rtb5.Text
    lblWeb.FontUnderline = rtb5.SelUnderline
    lblWeb.Visible = True
    Timer6.Interval = 0
End If
End Sub
