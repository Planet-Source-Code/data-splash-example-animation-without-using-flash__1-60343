VERSION 5.00
Begin VB.Form frmIntro 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2595
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5430
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmIntro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrLabel 
      Interval        =   200
      Left            =   720
      Top             =   1920
   End
   Begin VB.Timer tmrImage 
      Interval        =   200
      Left            =   240
      Top             =   1920
   End
   Begin VB.OLE OLE2 
      Class           =   "SoundRec"
      Height          =   375
      Left            =   2640
      OleObjectBlob   =   "frmIntro.frx":000C
      TabIndex        =   7
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Latha"
         Size            =   36
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   2355
      Left            =   0
      TabIndex        =   6
      Top             =   -120
      Width           =   5655
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Designed By Data @ Yahell Pro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   960
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "yahell-pro.us"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1035
      Left            =   360
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Today"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1035
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Yahell Pro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1035
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Visit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1035
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.OLE OLE1 
      Class           =   "SoundRec"
      Height          =   375
      Left            =   2160
      OleObjectBlob   =   "frmIntro.frx":CA24
      TabIndex        =   0
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   2355
      Left            =   0
      Picture         =   "frmIntro.frx":5D63C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'I forgot where I got this for rounding form corners..it's not mine but works great...
Private Sub Form_Activate()
SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

    Dim hRgn As Long, x As Long
    Dim formWidth As Single, formHeight As Single
    Dim borderWidth As Single, titleHeight As Single
    ' Calculate the form area
    borderWidth = (Me.Width - Me.ScaleWidth) / 2
    titleHeight = Me.Height - Me.ScaleHeight - borderWidth
    ' Convert to Pixels
    borderWidth = ScaleX(borderWidth, vbTwips, vbPixels)
    titleHeight = ScaleY(titleHeight, vbTwips, vbPixels)
    formWidth = Me.ScaleX(Me.ScaleWidth + borderWidth, vbTwips, vbPixels)
    formHeight = Me.ScaleY(Me.ScaleHeight + titleHeight, vbTwips, vbPixels)
    ' Create a round rectangle region around the graphics area of the form
    hRgn = CreateRoundRectRgn(borderWidth, titleHeight, formWidth + borderWidth, _
                              formHeight + titleHeight, 30, 30)
    ' Set the clipping area of the window using the resulting region
    SetWindowRgn hWnd, hRgn, True
    ' Tidy up
    x = DeleteObject(hRgn)
    DoEvents
End Sub

Private Sub label8_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'enables you to move the form from it's position
MoveForm Me

End Sub
Private Sub Form_Load()
'shrinks the image to 15 then the image timer takes over
Image1.Height = 15

End Sub


Private Sub ProgressEnd()
' just a sub to kill everything that hasn't been killed to go to the next form
    frmIntro.Visible = False
frmMain.Visible = True
  OLE2.DoVerb 1
    frmMain.Show
Unload Me

End Sub

Private Sub tmrImage_Timer()
'this timer steps the image size by 20 and when it reaches it's normal size it ends the timer

    Image1.Height = Image1.Height + 20
        If Image1.Height >= 2355 Then tmrImage = False
        
End Sub

Private Sub tmrLabel_Timer()
'this timer just randomly hides and shows the stacked labels giving it the intro effect
OLE1.DoVerb
'I adjusted the pauses for each label change to fit to the sound being played
Pause 1
    Label2.Visible = True
Pause 5
    Label2.Visible = False
Pause 1
    Label3.Visible = True
Pause 5
    Label3.Visible = False
Pause 1
    Label4.Visible = True
Pause 5
    Label4.Visible = False
Pause 1
    Label5.Visible = True
Pause 2
    Label6.Visible = True
Pause 5
        Call ProgressEnd
    tmrLabel = False
End Sub
