VERSION 5.00
Begin VB.Form fMain 
   BackColor       =   &H00FF80FF&
   BorderStyle     =   4  'Festes Werkzeugfenster
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13635
   ForeColor       =   &H000000C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   582
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   909
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer tmrMagn 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   8025
      Top             =   7905
   End
   Begin VB.CommandButton btMagic 
      BackColor       =   &H000040C0&
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   1.5
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7080
      Left            =   90
      MouseIcon       =   "fMain.frx":0000
      MousePointer    =   99  'Benutzerdefiniert
      Style           =   1  'Grafisch
      TabIndex        =   0
      Top             =   -60
      Width           =   6900
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8310
      Left            =   7215
      ScaleHeight     =   8250
      ScaleWidth      =   6075
      TabIndex        =   2
      Top             =   180
      Width           =   6135
   End
   Begin VB.CommandButton btAgain 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Amazing, isn't it? Try again with a different number?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   7920
      MouseIcon       =   "fMain.frx":030A
      MousePointer    =   99  'Benutzerdefiniert
      Style           =   1  'Grafisch
      TabIndex        =   5
      Top             =   2325
      Width           =   4470
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   $"fMain.frx":0614
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1215
      Index           =   0
      Left            =   165
      TabIndex        =   4
      Top             =   7095
      Width           =   6750
   End
   Begin VB.Label lb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Example: 68 / 10 = 6 remainder 8  >>  68 - 6 - 8 = 54"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Index           =   1
      Left            =   150
      TabIndex        =   1
      Top             =   8295
      Width           =   5190
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
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
      Height          =   1215
      Index           =   2
      Left            =   180
      TabIndex        =   3
      Top             =   7110
      Width           =   6750
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Const SnipOff As Long = 16 'off the button corners
Private RandomChar As Long

Private Sub btAgain_Click()

    pic.Visible = True
    btMagic.Enabled = True

End Sub

Private Sub btMagic_Click()

    If tmrMagn.Enabled = False And pic.Visible Then
        tmrMagn.Enabled = True
        btMagic.Caption = Chr$(RandomChar)
    End If

End Sub

Private Sub Display(i As Long, Posn As Long)

    pic.CurrentX = Posn
    pic.Font = "Arial"
    pic.FontSize = 10
    pic.Print i;
    pic.Font = "Wingdings"
    pic.FontSize = 15
    pic.CurrentX = Posn + 600
    If i Mod 9 Then
        pic.Print Chr$(Rnd * 222 + 33);
      Else 'NOT I...
        pic.Print Chr$(RandomChar);
    End If

End Sub

Private Sub Form_Load()

  Dim i As Long
  Dim hRgn As Long

    Randomize -Timer
    lb(2) = lb(0)

    i = btMagic.Height - SnipOff
    hRgn = CreateRoundRectRgn(SnipOff, SnipOff, i, i, i, i)
    SetWindowRgn btMagic.hWnd, hRgn, True
    DeleteObject hRgn
    RandomChar = Rnd * 222 + 33
    pic.Cls
    For i = 1 To 97 Step 4
        Display i, 0
        Display i + 1, 1620
        Display i + 2, 3240
        Display i + 3, 4860
        pic.Print
    Next i

End Sub

Private Sub tmrMagn_Timer()

    btMagic.FontSize = btMagic.FontSize + 1
    If btMagic.FontSize >= 250 Then
        tmrMagn.Enabled = False
        btMagic.FontSize = 1
        btMagic.Caption = vbNullString
        pic.Visible = False
        Form_Load
    End If

End Sub

':) Ulli's VB Code Formatter V2.21.6 (2006-Sep-02 16:52)  Decl: 7  Code: 74  Total: 81 Lines
':) CommentOnly: 0 (0%)  Commented: 2 (2,5%)  Empty: 19 (23,5%)  Max Logic Depth: 2
