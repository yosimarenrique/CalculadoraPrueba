VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calculator"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   293
   ScaleMode       =   0  'User
   ScaleWidth      =   242
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnNum0 
      Caption         =   "0"
      Enabled         =   0   'False
      Height          =   375
      Left            =   960
      TabIndex        =   10
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton btnequal 
      Caption         =   "="
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   16
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton btnAC 
      Caption         =   "AC"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton btndot 
      Caption         =   "."
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   15
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton btnADD 
      Caption         =   "+"
      Enabled         =   0   'False
      Height          =   855
      Left            =   2760
      TabIndex        =   17
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton btnC 
      Caption         =   "C"
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton btnSQR 
      Caption         =   "sqrt"
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton btnNum1 
      Caption         =   "1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton btnNum3 
      Caption         =   "3"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton btnNum2 
      Caption         =   "2"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton btnNum4 
      Caption         =   "4"
      Enabled         =   0   'False
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton btnNum5 
      Caption         =   "5"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton btnPrecent 
      Caption         =   "%"
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton btnNum6 
      Caption         =   "6"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton btnDEC 
      Caption         =   "-"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   18
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton btnOFF 
      Caption         =   "OFF"
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton btnSTAR 
      Caption         =   "*"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   19
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton btnNum9 
      Caption         =   "9"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton btnNum8 
      Caption         =   "8"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton btnNum7 
      Caption         =   "7"
      Enabled         =   0   'False
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton btnMminus 
      Caption         =   "M-"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   23
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton btnMRC 
      Caption         =   "MRC"
      Enabled         =   0   'False
      Height          =   375
      Left            =   960
      TabIndex        =   22
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton btnDIVslash 
      Caption         =   "/"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   20
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton btnMplus 
      Caption         =   "M+"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   21
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox txtLCD 
      Enabled         =   0   'False
      Height          =   315
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "Press on button to run..."
      ToolTipText     =   "LCD Display"
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label lblOn 
      Caption         =   "on"
      Height          =   195
      Left            =   480
      TabIndex        =   28
      Top             =   4080
      Width           =   180
   End
   Begin VB.Label lblLC 
      Caption         =   "LC-403LD"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1395
      TabIndex        =   27
      Top             =   840
      Width           =   840
   End
   Begin VB.Label lblEC 
      Caption         =   "ELECTRONIC CALCULATOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   525
      TabIndex        =   26
      Top             =   600
      Width           =   2580
   End
   Begin VB.Label lblCasio 
      Caption         =   "CASIO"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1365
      TabIndex        =   25
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Operator As String, Variable As String, Bring As Boolean

Private Sub focus()
    txtLCD.SetFocus
    txtLCD.SelStart = 0
    txtLCD.SelLength = Len(txtLCD.Text)
End Sub

Private Sub SetStatus(state As Boolean)
    txtLCD.Enabled = state
    btnOFF.Enabled = state
    btnDIVslash.Enabled = state
    btnPrecent.Enabled = state
    btnSTAR.Enabled = state
    btnSQR.Enabled = state
    btnDEC.Enabled = state
    btnC.Enabled = state
    btndot.Enabled = state
    btnADD.Enabled = state
    btnequal.Enabled = state
    btnNum0.Enabled = state
    btnNum1.Enabled = state
    btnNum2.Enabled = state
    btnNum3.Enabled = state
    btnNum4.Enabled = state
    btnNum5.Enabled = state
    btnNum6.Enabled = state
    btnNum7.Enabled = state
    btnNum8.Enabled = state
    btnNum9.Enabled = state
End Sub

Private Sub NumPressed(str As String)
    If txtLCD.Text = "0" Then txtLCD.Text = ""
    If Bring = False Then
        txtLCD.Text = txtLCD.Text + str
    Else
        Variable = txtLCD.Text
        txtLCD.Text = str
    End If
    Bring = False
    focus
End Sub

Private Sub OperatorPressed(CurrentOperator As String)
    If (Variable <> "") Then
        Select Case Operator
            Case "+"
                txtLCD.Text = CCur(Variable) + CCur(txtLCD.Text)
            Case "-"
                txtLCD.Text = CCur(Variable) - CCur(txtLCD.Text)
            Case "*"
                txtLCD.Text = CCur(Variable) * CCur(txtLCD.Text)
            Case "/"
                If (txtLCD.Text <> "0") Then
                    txtLCD.Text = CCur(Variable) / CCur(txtLCD.Text)
                End If
        End Select
    End If
    If (Operator <> "/") Or (txtLCD.Text <> "0") Then
        Operator = CurrentOperator
        Bring = True
        focus
    Else
        btnOFF_Click
        txtLCD.Text = "Error, Division by zero"
    End If
End Sub

Private Sub btnAC_Click()
    SetStatus (True)
    txtLCD.Text = "0"
    focus
End Sub

Private Sub btnADD_Click()
    OperatorPressed ("+")
End Sub

Private Sub btnC_Click()
    txtLCD.Text = "0"
    focus
End Sub

Private Sub btnDEC_Click()
    OperatorPressed ("-")
End Sub

Private Sub btnDIVslash_Click()
    OperatorPressed ("/")
End Sub

Private Sub btndot_Click()
    If (InStr(txtLCD.Text, ".") = 0) Then txtLCD.Text = txtLCD.Text + "."
    Bring = False
    focus
End Sub

Private Sub btnequal_Click()
    If (Variable <> "") Then
        OperatorPressed ("=")
    Else
        Select Case Operator
            Case "*"
                txtLCD.Text = CCur(txtLCD.Text) * CCur(txtLCD.Text)
            Case "/"
                txtLCD.Text = "1"
        End Select
        Variable = ""
        Operator = "="
        focus
    End If
End Sub

Private Sub btnNum0_Click()
    NumPressed ("0")
End Sub

Private Sub btnNum1_Click()
    NumPressed ("1")
End Sub

Private Sub btnNum2_Click()
    NumPressed ("2")
End Sub

Private Sub btnNum3_Click()
    NumPressed ("3")
End Sub

Private Sub btnNum4_Click()
    NumPressed ("4")
End Sub

Private Sub btnNum5_Click()
    NumPressed ("5")
End Sub

Private Sub btnNum6_Click()
    NumPressed ("6")
End Sub

Private Sub btnNum7_Click()
    NumPressed ("7")
End Sub

Private Sub btnNum8_Click()
    NumPressed ("8")
End Sub

Private Sub btnNum9_Click()
    NumPressed ("9")
End Sub

Private Sub btnOFF_Click()
    SetStatus (False)
    Operator = ""
    Variable = ""
    Bring = True
    txtLCD.Text = "Press on button to run..."
End Sub

Private Sub btnPrecent_Click()
    If (Variable <> "") Then
        Select Case Operator
            Case "+"
                txtLCD.Text = CCur(Variable) + (3 / 100)
            Case "-"
                txtLCD.Text = CCur(Variable) - (3 / 100)
            Case "*"
                txtLCD.Text = CCur(Variable) * (3 / 100)
            Case "/"
                txtLCD.Text = CCur(Variable) / (3 / 100)
        End Select
    End If
    Operator = "%"
    focus
End Sub

Private Sub btnSQR_Click()
    txtLCD.Text = Sqr(CCur(txtLCD.Text))
    Operator = "sqrt"
    focus
End Sub

Private Sub btnSTAR_Click()
    OperatorPressed ("*")
End Sub
