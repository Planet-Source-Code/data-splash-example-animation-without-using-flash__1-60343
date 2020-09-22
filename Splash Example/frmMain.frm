VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Main Form"
   ClientHeight    =   2820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      Height          =   2775
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   3495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Data     http://www.yahell-pro.us"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMain.frx":0000
      ForeColor       =   &H000080FF&
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "frmMain"
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


Private Sub Label3_Click()

Unload Me

End Sub

Private Sub Form_Unload(cancle As Integer)
'Using End is gay...try avoid using!!!!
End

End Sub

Private Sub label4_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'enables you to move the form from it's position
MoveForm Me

End Sub
