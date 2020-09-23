VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Derivation"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7515
   FontTransparent =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Form1.frx":1042
   ScaleHeight     =   4500
   ScaleWidth      =   7515
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtResult 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      TabIndex        =   5
      Top             =   1680
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Start Derivation"
      Height          =   975
      Left            =   960
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form1.frx":326A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox txtFormula 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   3360
      TabIndex        =   0
      Top             =   960
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00939A92&
      Caption         =   "Steps"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   3240
      TabIndex        =   4
      Top             =   600
      Width           =   4095
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   600
      ToolTipText     =   "Exit"
      Top             =   4200
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   120
      ToolTipText     =   "Minimize"
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   9
      Top             =   325
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   8
      Top             =   330
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Derivation Program"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   325
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   285
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   1680
      Width           =   285
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      Shape           =   3  'Circle
      Top             =   1680
      Width           =   285
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   285
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   960
      Width           =   285
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      Shape           =   3  'Circle
      Top             =   960
      Width           =   285
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Formula to Derivate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XPos As String 'The string of the Position of X
Dim PPos As Integer
Dim MPos As Integer
Dim Factor As String 'The factor of x
Dim Formula As String ' the complete formula
Dim Exp As String 'the exponent of x
Dim Derivation As String 'the result
Dim FirstNumberSign As String 'If the first number has a minus or plus
'This program works in the following way: It starts looking for a minus or
'plus sign. When the code finds it in the formula TextBox, it returns a value
'and its the number of digits from the start to the sign. Then the code sets
'the new formula as everything except the first set of number (eg: -12x2)
'then it starts working with the set of numbers, all before the x is the
'factor and all after the x is the exponent.
'Then it does the mathematical procedure to calculate de derivate.
'And finally, it loops and continues qith the rest of the formula
Private Sub Command1_Click()
List1.Clear 'clears the list to make an other derivation
Derivation = "" 'set derivation as nothing
Formula = txtFormula.Text 'set formula as the entry in text1
Formula = Trim(Formula) 'removes any space put by the user
Formula = LCase(Formula) 'set the formula as lower case to avoid X or x problems

If txtFormula.Text = "" Then
List1.AddItem "No formula to derivate"
txtFormula.SetFocus
Exit Sub
ElseIf Len(Formula) = 1 And Mid(Formula, 1, 1) = "+" Then
txtFormula.SetFocus
List1.AddItem "Dont put just the mathematical sign"
Exit Sub
ElseIf Len(Formula) = 1 And Mid(Formula, 1, 1) = "-" Then
txtFormula.SetFocus
List1.AddItem "Dont put just the mathematical sign"
Exit Sub
End If

List1.AddItem "The formula to derivate is: " & Formula

ViewComponents: 'the label to return afterwords
'First check if the first digit is a minus or plus
FirstNumberSign = Mid(Formula, 1, 1)
If FirstNumberSign = "+" Or FirstNumberSign = "-" Then
    FirstNumberSign = 2
Else
    FirstNumberSign = 1
End If

'searchs the first x in the formula
XPos = InStr(FirstNumberSign, Formula, "x")

If XPos = 0 Then GoTo EndFormula 'if there is no x then go to the endformula

Factor = Mid(Formula, 1, XPos - 1) 'set the factor as the text from the beggining of the formula up to the x
PPos = InStr(FirstNumberSign, Formula, "+") 'search the first +
MPos = InStr(FirstNumberSign, Formula, "-") 'search the first -

If PPos = 0 And MPos = 0 Then 'if there is no + nor -
List1.AddItem "The expresion to derivate is: " & Formula
Exp = Mid(Formula, XPos + 1, Len(Formula) - XPos) 'the exponent is the text from the x to the end of the formula
Formula = "No more Formula"
If Exp = "" Then Exp = 1
List1.AddItem "The factor is: " & Factor
List1.AddItem "The exponent is: " & Exp
GoTo Calculation
End If

If PPos = 0 Then PPos = Len(Formula) + 1 'if there is no + then set it to be the lengh of formula plus 1
If MPos = 0 Then MPos = Len(Formula) + 1 'if there is no - then set it to be the lengh of formula plus 1

If PPos < MPos Then 'if the first sign is + then do everything with the plus position
List1.AddItem "The expresion to derivate is: " & Mid(Formula, 1, PPos - 1)
Exp = Mid(Formula, XPos + 1, PPos - XPos - 1)
Formula = Right(Formula, Len(Formula) - PPos + 1) 'the formula is everything except the factor and exponent we just saw
If Exp = "" Then Exp = 1
List1.AddItem "The factor is: " & Factor
List1.AddItem "The exponent is: " & Exp
ElseIf MPos < PPos Then 'if the first sign is - then do everything with the minus position
List1.AddItem "The expresion to derivate is: " & Mid(Formula, 1, MPos - 1)
Exp = Mid(Formula, XPos + 1, MPos - XPos - 1)
Formula = Right(Formula, Len(Formula) - MPos + 1) 'the formula is everything except the factor and exponent we just saw
If Exp = "" Then Exp = 1
List1.AddItem "The factor is: " & Factor
List1.AddItem "The exponent is: " & Exp
End If

Calculation:

If Exp = "" Then Exp = 1 'if the exponent is none then set it to 1 (12x = 12x1)

If Exp = "1" Then 'if exponent is 1 then avoid 12x0 and put 12
Derivation = Derivation & Factor
GoTo ViewComponents 'go to see the rest of the formula
End If

Factor = Val(Factor) * Val(Exp) 'multiplicates the factor with the exponent as derivating
If Mid(Factor, 1, 1) <> "-" Then Factor = "+" & Factor
Derivation = Derivation & Factor
Exp = Val(Exp) - 1 'rests the exponent 1 as derivating

If Exp = 1 Then 'if exponent is 1 after resting 1 then avoid 12x1 and put 12x
Derivation = Derivation & "*" & "x"
GoTo ViewComponents 'go to see the rest of the formula
End If

Derivation = Derivation & "*" & "x" & "^" & Exp
GoTo ViewComponents 'go to see the rest of the formula

EndFormula:
If Mid(Derivation, 1, 1) = "+" Then Derivation = Mid(Derivation, 2, Len(Derivation) - 1)
txtResult.Text = Derivation
List1.AddItem "The derivation is: " & Derivation
Exit Sub
End Sub

Private Sub Form_Load()

End Sub

Private Sub Image1_Click()
Me.WindowState = 1
End Sub

Private Sub Image3_Click()
End
End Sub

Private Sub Label4_Click()
End
End Sub

Private Sub Label5_Click()
Me.WindowState = 1
End Sub
