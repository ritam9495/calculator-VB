VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MyCalculator"
   ClientHeight    =   5235
   ClientLeft      =   5610
   ClientTop       =   3210
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   4485
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command24 
      Caption         =   "x!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   36
      Top             =   840
      Width           =   615
   End
   Begin VB.OptionButton OpD 
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   35
      ToolTipText     =   "Degree"
      Top             =   1920
      Width           =   615
   End
   Begin VB.OptionButton OpR 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   34
      ToolTipText     =   "Radian"
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command23 
      Caption         =   "tan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   33
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton Command22 
      Caption         =   "cos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   32
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Command21 
      Caption         =   "sin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   31
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   "ln"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   30
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton Command20 
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   27
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command19 
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   26
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Command18 
      Caption         =   "root"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   25
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton Command17 
      Caption         =   "M-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   24
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton Command16 
      Caption         =   "M+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   23
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Command15 
      Caption         =   "MR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   22
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton Command14 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   21
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command13 
      Caption         =   "AC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   20
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Command12 
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   19
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   18
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H8000000D&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   17
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "1/x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   16
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2280
      TabIndex        =   14
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   13
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   12
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   11
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   840
      TabIndex        =   8
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   1560
      TabIndex        =   7
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   840
      TabIndex        =   5
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   1560
      TabIndex        =   4
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   840
      TabIndex        =   2
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label lblDisp 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000B&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   28
      Top             =   840
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim decp, nter As Boolean
Dim num1, num2, mem, PI As Double
Dim op As Integer

Private Sub Command1_Click(Index As Integer)
    If nter = True Then
        lblDisp.Caption = ""
        nter = False
    End If
    If lblDisp.Caption = "0" Then
        lblDisp.Caption = ""
    End If
    lblDisp.Caption = lblDisp.Caption & Index
End Sub

Private Sub Command10_Click()
    If lblDisp <> 0 Then
        lblDisp.Caption = Str(Log(lblDisp.Caption))
    Else
        lblDisp.Caption = "NAN"
    End If
End Sub

Private Sub Command11_Click()
    s = MsgBox("Do you really want to exit?", vbYesNo, "Exit")
    If s = vbYes Then
        End
    End If
End Sub

Private Sub Command12_Click()
    lblDisp.Caption = ""
End Sub

Private Sub Command13_Click()
    Call Form_Load
End Sub

Private Sub Command14_Click()
    If num1 <> 0 And op = 4 Then
        lblDisp.Caption = Str(num1 - Val(lblDisp.Caption))
    End If
    num1 = Val(lblDisp.Caption)
    op = 4
    nter = True
    decp = False
End Sub

Private Sub Command15_Click()
    lblDisp.Caption = Str(mem)
    nter = True
End Sub

Private Sub Command16_Click()
    mem = Val(lblDisp.Caption) + mem
    If mem = 0 Then
        Label1.Caption = ""
        lblDisp.Caption = Str(mem)
    Else
        lblDisp.Caption = Str(mem)
        Label1.Caption = "M"
    End If
    nter = True
End Sub

Private Sub Command17_Click()
    mem = Val(lblDisp.Caption) - mem
    If mem = 0 Then
        Label1.Caption = ""
        lblDisp.Caption = Str(mem)
    Else
        lblDisp.Caption = Str(mem)
        Label1.Caption = "M"
    End If
    nter = True
End Sub

Private Sub Command18_Click()
    If Val(lblDisp.Caption) < 0 Then
        lblDisp.Caption = "Math error"
    Else
        lblDisp.Caption = Str(Sqr(Val(lblDisp.Caption)))
    End If
End Sub

Private Sub Command19_Click()
    l = Len(lblDisp.Caption)
    If l = 1 Or l = 0 Then
        lblDisp.Caption = ""
    Else
        Str1 = Mid(lblDisp.Caption, 1, l - 1)
        p = Mid(lblDisp.Caption, l - 1, 1)
        If p = "." Then
            decp = False
        End If
        lblDisp.Caption = Str1
    End If
End Sub

Private Sub Command2_Click()
    If nter = True Then
        lblDisp.Caption = ""
    End If
    If lblDisp.Caption = "" Then
        lblDisp.Caption = "0."
        decp = True
    End If
    If lblDisp.Caption = "0" And decp = False Then
        lblDisp.Caption = lblDisp.Caption & "."
        decp = True
        End If
    If decp = False Then
        lblDisp.Caption = lblDisp.Caption & "."
        decp = True
    End If
    nter = False
End Sub

Private Sub Command20_Click()
    mem = Val(lblDisp.Caption)
    If mem <> 0 Then
        Label1.Caption = "M"
        Else
        Label1.Caption = ""
    End If
    nter = True
End Sub

Private Sub Command21_Click()
    Dim x As Double
    x = lblDisp
    If OpD = True And x <> 0 Then
        x = PI * x / 180
    End If
    lblDisp.Caption = Str(Sin(x))
    nter = True
End Sub

Private Sub Command22_Click()
    Dim x As Double
    x = lblDisp
    If OpD = True And x <> 0 Then
        x = PI * x / 180
    End If
    lblDisp.Caption = Str(Cos(x))
    nter = True
End Sub

Private Sub Command23_Click()
    Dim x As Double
    x = lblDisp
    If OpD = True And x <> 0 Then
        x = PI * x / 180
    End If
    lblDisp.Caption = Str(Tan(x))
    nter = True
End Sub

Private Sub Command24_Click()
    Dim f As Double
    f = 1
    x = Val(lblDisp.Caption)
    While x > 0
        f = f * x
        x = x - 1
    Wend
    lblDisp.Caption = Str(f)
End Sub

Private Sub Command3_Click()
        lblDisp.Caption = Str(Val(lblDisp.Caption) * -1)
End Sub

Private Sub Command4_Click()
    If num1 <> 0 And op = 3 Then
        lblDisp.Caption = Str(num1 * Val(lblDisp.Caption))
    End If
    num1 = Val(lblDisp.Caption)
    op = 3
    nter = True
    decp = False
End Sub

Private Sub Command5_Click()
    If num1 <> 0 And op = 2 Then
        lblDisp.Caption = Str(num1 - Val(lblDisp.Caption))
    End If
    num1 = Val(lblDisp.Caption)
    op = 2
    nter = True
    decp = False
End Sub

Private Sub Command6_Click()
    If num1 <> 0 And op = 1 Then
        lblDisp.Caption = Str(num1 + Val(lblDisp.Caption))
    End If
    num1 = Val(lblDisp.Caption)
    op = 1
    nter = True
    decp = False
End Sub

Private Sub Command7_Click()
    lblDisp.Caption = Str(num1 * Val(lblDisp.Caption) / 100)
    nter = True
End Sub

Private Sub Command8_Click()
    If Val(lblDisp.Caption) = 0 Then
        lblDisp.Caption = "Divide by zero"
    Else
        lblDisp.Caption = Str(1 / Val(lblDisp.Caption))
    End If
    nter = True
End Sub

Private Sub Command9_Click()
    num2 = Val(lblDisp.Caption)
    If op = 1 Then
        lblDisp.Caption = Str(num1 + num2)
    ElseIf op = 2 Then
        lblDisp.Caption = Str(num1 - num2)
    ElseIf op = 3 Then
        lblDisp.Caption = Str(num1 * num2)
    ElseIf op = 4 Then
        If num2 = 0 Then
            lblDisp.Caption = "Divide by zero error"
        Else
            lblDisp.Caption = Str(num1 / num2)
        End If
    End If
    nter = True
End Sub

Private Sub Form_Load()
    decp = False
    nter = False
    mem = 0
    OpR = True
    op = 0
    lblDisp = "0"
    num1 = 0
    num2 = 0
    PI = 3.14159265
End Sub
