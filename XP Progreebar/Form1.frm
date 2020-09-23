VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "XP Progressbar Demo"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   2160
      Top             =   960
   End
   Begin Project1.XPProgressBar XP 
      Height          =   240
      Index           =   0
      Left            =   840
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   423
      Value           =   0.7
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   1920
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reset"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   2655
   End
   Begin Project1.XPProgressBar XP 
      Height          =   240
      Index           =   1
      Left            =   840
      Top             =   480
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   423
      Style           =   1
      Value           =   0.7
   End
   Begin Project1.XPProgressBar XP 
      Height          =   240
      Index           =   2
      Left            =   840
      Top             =   840
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   423
      Style           =   2
      Value           =   0.7
   End
   Begin Project1.XPProgressBar XP 
      Height          =   240
      Index           =   3
      Left            =   840
      Top             =   1200
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   423
      Style           =   3
      Value           =   0.7
   End
   Begin Project1.XPProgressBar XP 
      Height          =   240
      Index           =   4
      Left            =   840
      Top             =   1560
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   423
      Value           =   0.7
      Enabled         =   0   'False
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Disabled"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Blue"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Red"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Golden"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Green"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   435
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim r(4) As Single
Private Sub Command1_Click()
Form_Load
End Sub
Private Sub Command2_Click()
End
End Sub
Private Sub Form_Load()
Randomize Timer
Dim i As Integer
i = 0
Do While i < 4
    r(i) = Rnd * 0.94
    XP(i).Value = r(i)
    i = i + 1
Loop
End Sub
Private Sub Timer1_Timer()
Dim i As Integer
i = 0
Do While i < 4
    If r(i) < 0.95 Then
        r(i) = r(i) + 0.01 + (i / 200)
    Else
        r(i) = 0
    End If
    XP(i).Value = r(i)
    i = i + 1
Loop
End Sub
