VERSION 5.00
Begin VB.Form frmClock 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F468AE&
   BorderStyle     =   0  'None
   ClientHeight    =   4035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4020
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   3480
      Top             =   3480
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial MT Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   3120
      Width           =   3975
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "By Paul Crowdy."
      BeginProperty Font 
         Name            =   "Arial MT Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   2760
      Width           =   4095
   End
   Begin VB.Label lblDblClick 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Double-Click to Close"
      BeginProperty Font 
         Name            =   "Arial MT Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   2040
      Width           =   3975
   End
   Begin VB.Line lneSecond 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   2040
      X2              =   2640
      Y1              =   1920
      Y2              =   720
   End
   Begin VB.Line lneMinute 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   4
      X1              =   2040
      X2              =   2040
      Y1              =   1920
      Y2              =   720
   End
   Begin VB.Line lneHour 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   6
      X1              =   2040
      X2              =   2640
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label lblThree 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   570
      Left            =   3600
      TabIndex        =   3
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label lbNine 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   570
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label lblSix 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   570
      Left            =   1800
      TabIndex        =   1
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label lblTwelve 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   570
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Clock Written by Paul Crowdy. http://www.kmcpartnership.co.uk/
Private Sub Form_DblClick()
    End
End Sub

Private Sub Form_Load()
    'Assign a region on the screen in the shape of a circle
    lRegion = CreateEllipticRgn(0, 0, 270, 270)
    'Show the form in the region
    lResult = SetWindowRgn(Me.hWnd, lRegion, True)
    Me.Show
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  drag Me
End Sub

Private Sub lblDblClick_DblClick()
    End
End Sub

Private Sub lblDblClick_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    drag Me
End Sub

Private Sub lblName_DblClick()
    End
End Sub

Private Sub lblName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    drag Me
End Sub

Private Sub lblSix_DblClick()
    End
End Sub

Private Sub lblSix_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    drag Me
End Sub

Private Sub lblThree_DblClick()
    End
End Sub

Private Sub lblThree_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    drag Me
End Sub

Private Sub lblTime_DblClick()
    End
End Sub

Private Sub lblTime_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    drag Me
End Sub

Private Sub lblTwelve_DblClick()
    End
End Sub

Private Sub lblTwelve_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    drag Me
End Sub

Private Sub lbNine_DblClick()
    End
End Sub

Private Sub lbNine_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    drag Me
End Sub

Private Sub Timer1_Timer()
Dim hnd As Line
Dim sTime As String
Dim iHour As Integer
Dim iMinute As Integer
Dim iSecond As Integer
Dim iHandLength As Integer
Dim iAngle As Integer
Dim iUpMove As Integer
Dim iRightMove As Integer
Dim i As Integer

    sTime = Format(Time, "HH:MM:SS") 'Get Current Time
    
    iHour = Val(Mid$(sTime, 1, 2))
    iMinute = Val(Mid$(sTime, 4, 2))
    iSecond = Val(Mid$(sTime, 7, 2))
    
    If iHour > 12 Then iHour = iHour - 12 '12 hour clock
    
    For i = 0 To 2
        Select Case i
            Case 0
                'Set Hour hand
                iAngle = (iHour * 30) + (iMinute / 2)
                iHandLength = 720
                Set hnd = lneHour
            Case 1
                'Set Minute hand
                iAngle = (6 * iMinute)
                iHandLength = 1250
                Set hnd = lneMinute
            Case 2
                'Set Second hand
                iAngle = 6 * iSecond
                iHandLength = 1400
                Set hnd = lneSecond
        End Select
        
        'Set one end of hand to centre of the form
        hnd.x1 = Me.Width / 2
        hnd.y1 = Me.Height / 2
        
        'Send other end of hands to correct place
        hnd.x2 = hnd.x1 + (Sin(iAngle * (PI / 180)) * iHandLength)
        hnd.y2 = hnd.y1 - (Cos(iAngle * (PI / 180)) * iHandLength)
        
    Next i
    lblTime.Caption = Format(Time, "hh:mm:ss")
    Me.Refresh
End Sub
