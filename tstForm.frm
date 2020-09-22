VERSION 5.00
Begin VB.Form tstForm 
   Caption         =   "Form1"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Clear"
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add"
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   2280
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   1560
      TabIndex        =   8
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Original MsgBox"
      Height          =   495
      Left            =   2400
      TabIndex        =   7
      Top             =   4080
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Text            =   "Spot the real message box!"
      Top             =   1560
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Text            =   "This custom message box uses a parameter array to have any buttons you like"
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test MsgBox"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label6 
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Return:"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Buttons (Test MsgBox only)"
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Title"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Image"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Prompt"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "tstForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim CMBS As CMBStandardIconEnum
  Select Case Combo1.ListIndex + 1
    Case 1
      CMBS = IDI_NONE
    Case 2
      CMBS = IDI_QUESTION
    Case 3
      CMBS = IDI_EXCLAMATION
    Case 4
      CMBS = IDI_HAND
    Case 5
      CMBS = IDI_ASTERISK
    Case Else
      CMBS = IDI_NONE
  End Select
  Select Case List1.ListCount
    Case 0
      Label6.Caption = CustomMessageBox.CustomMessageBox(Text1.Text, CMBS, Text2.Text)
    Case 1
      Label6.Caption = CustomMessageBox.CustomMessageBox(Text1.Text, CMBS, Text2.Text, List1.List(0))
    Case 2
      Label6.Caption = CustomMessageBox.CustomMessageBox(Text1.Text, CMBS, Text2.Text, List1.List(0), List1.List(1))
    Case 3
      Label6.Caption = CustomMessageBox.CustomMessageBox(Text1.Text, CMBS, Text2.Text, List1.List(0), List1.List(1), List1.List(2))
    Case 4
      Label6.Caption = CustomMessageBox.CustomMessageBox(Text1.Text, CMBS, Text2.Text, List1.List(0), List1.List(1), List1.List(2), List1.List(3))
  End Select
  
End Sub

Private Sub Command2_Click()
Dim CMBS As VbMsgBoxStyle
  Select Case Combo1.ListIndex + 1
    Case 1
      CMBS = 0
    Case 2
      CMBS = vbQuestion
    Case 3
      CMBS = vbExclamation
    Case 4
      CMBS = vbCritical
    Case 5
      CMBS = vbInformation
    Case Else
      CMBS = 0
  End Select
  MsgBox Text1.Text, CMBS, Text2.Text
End Sub

Private Sub Command3_Click()
Dim tmpStr As String
  tmpStr = InputBox("Please enter command button text")
  List1.AddItem tmpStr
End Sub

Private Sub Command4_Click()
  List1.Clear
End Sub

Private Sub Form_Load()
  ' populate combo
  Combo1.AddItem "None"
  Combo1.AddItem "Question"
  Combo1.AddItem "Exclamation"
  Combo1.AddItem "Critical"
  Combo1.AddItem "Asterisk"
  Combo1.ListIndex = 0
  List1.AddItem "Ok"
End Sub
