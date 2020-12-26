VERSION 5.00
Begin VB.Form frmRecursion 
   Caption         =   "Simple Recursion"
   ClientHeight    =   2670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4260
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2670
   ScaleWidth      =   4260
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Algorithm"
      ForeColor       =   &H8000000D&
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2295
      Begin VB.OptionButton optAlgorithm 
         Caption         =   "Fibonacci"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   1140
         Width           =   1935
      End
      Begin VB.OptionButton optAlgorithm 
         Caption         =   "Greatest Common Divisor"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   660
         Width           =   1935
      End
      Begin VB.OptionButton optAlgorithm 
         Caption         =   "Factorial"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.TextBox txtInput1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   180
         TabIndex        =   6
         Text            =   "9"
         Top             =   1860
         Width           =   615
      End
      Begin VB.TextBox txtInput2 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1380
         TabIndex        =   5
         Text            =   "12"
         Top             =   1860
         Width           =   615
      End
      Begin VB.Label lblInput2 
         AutoSize        =   -1  'True
         Caption         =   "Input 2"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1380
         TabIndex        =   8
         Top             =   1620
         Width           =   615
      End
      Begin VB.Label Labels 
         AutoSize        =   -1  'True
         Caption         =   "Input 1"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   1620
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdRecurse 
      Caption         =   "Recurse Now"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   360
      Width           =   1515
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2873
      TabIndex        =   3
      Top             =   2100
      Width           =   795
   End
   Begin VB.Label lblResult 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   270
      Left            =   3180
      TabIndex        =   2
      Top             =   1440
      Width           =   180
   End
   Begin VB.Label Labels 
      AutoSize        =   -1  'True
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   2887
      TabIndex        =   1
      Top             =   960
      Width           =   780
   End
End
Attribute VB_Name = "frmRecursion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const cFAC% = 0
Private Const cGCD% = 1
Private Const cFIB% = 2

Private SelectedAlgorithm%
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub GetFactorial()
    
    If Not IsNumeric(txtInput1.Text) Then
        Beep
        Exit Sub
    End If
    
    lblResult = Factorial(txtInput1)

End Sub


Private Sub GetFibonacci()

    Me.MousePointer = vbHourglass
    If Not IsNumeric(txtInput1.Text) Then
        Beep
        Exit Sub
    End If
    
    lblResult = Fibonacci(CInt(txtInput1.Text))
    Me.MousePointer = vbDefault
    
End Sub

Private Sub GetGCD()

    If Not IsNumeric(txtInput1.Text) Or Not IsNumeric(txtInput2.Text) Then
        Beep
        Exit Sub
    End If
    
    lblResult = GreatestCommonDivisor(CInt(txtInput1.Text), CInt(txtInput2.Text))

End Sub



Private Sub cmdRecurse_Click()

    Select Case SelectedAlgorithm
        Case cFAC: GetFactorial
        Case cGCD: GetGCD
        Case cFIB: GetFibonacci
    End Select
    
End Sub


Private Sub optAlgorithm_Click(Index As Integer)
    txtInput2.Enabled = (Index = cGCD)
    lblInput2.Enabled = (Index = cGCD)
    SelectedAlgorithm = Index
End Sub


