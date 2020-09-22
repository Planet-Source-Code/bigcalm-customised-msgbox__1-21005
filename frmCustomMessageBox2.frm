VERSION 5.00
Begin VB.Form CustomMessageBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Message Box"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3465
   Icon            =   "frmCustomMessageBox2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   125
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   231
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdButton 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblCaption 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "CustomMessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0
Option Compare Text

' Code by Jonathan Daniel.  13/02/01
' Custom message box, but allows you to have any captions you like on the buttons.
' Should be very similar to Windows message box.

' Enum defines all possible Icons that can be used.
Public Enum CMBStandardIconEnum
    IDI_NONE = 0&   ' No Icon
    IDI_ASTERISK = 32516&       ' like vbInformation
    IDI_EXCLAMATION = 32515&    ' like vbExlamation
    IDI_HAND = 32513&           ' like vbCritical
    IDI_QUESTION = 32514&       ' like vbQuestion
End Enum


' API declarations to get a standard icon.
Private Declare Function LoadStandardIcon Lib "user32" Alias _
    "LoadIconA" (ByVal hInstance As Long, ByVal lpIconNum As _
    CMBStandardIconEnum) As Long
    
Private Declare Function DrawIcon Lib "user32" (ByVal hDC _
    As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal hIcon As Long) As Long
    
' API calls to disable close button on form
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const SC_CLOSE = &HF060&
Private Const MF_BYCOMMAND = &H0&

' Constants to make this look right
Private Const MaxMsgBoxWidth As Long = 500
Private Const MaxMsgBoxHeight As Long = 300
Private Const MinButtonWidth As Long = 75
' The returned argument
Private ClickedButton As String
' The icon handle.
Private hIcon As Long
Private mImg As CMBStandardIconEnum

' Call THIS function on this form to get a custom message box.
Public Function CustomMessageBox(Prompt As String, Img As CMBStandardIconEnum, Title As String, ParamArray ButtonCaptions() As Variant) As String
Dim tmpVar As Variant
Dim i As Long
Dim LargestButtonWidth As Long
Dim CaptionWidth As Long
Dim hSysMenu As Long

  ' Initialise
  Me.ScaleMode = 3 ' pixel
  LargestButtonWidth = 0
  mImg = Img
  
  ' Check for rubbish being passed
  If UBound(ButtonCaptions) < 0 Then
    Exit Function
  End If
  If UBound(ButtonCaptions) > 5 Then
    Exit Function
  End If
  If Len(Title) = 0 Then
    Me.Caption = App.Title
  Else
    Me.Caption = Title
  End If
  
  
  ' Placement and size (the hard bit)
  ' Ok, let's find the largest button caption
  For Each tmpVar In ButtonCaptions
    If Me.TextWidth(tmpVar) > LargestButtonWidth Then
      LargestButtonWidth = Me.TextWidth(tmpVar)
    End If
  Next
  ' Need to leave enough space to put text actually inside the button...
  LargestButtonWidth = LargestButtonWidth + 10
  If LargestButtonWidth < MinButtonWidth Then
    LargestButtonWidth = MinButtonWidth
  End If
  
  ' Is this button width too wide? (do we have to foreshorten it and make the height of the button larger)
  Do While (LargestButtonWidth + 20) * UBound(ButtonCaptions) > MaxMsgBoxWidth
    LargestButtonWidth = LargestButtonWidth / 2
    cmdButton(0).Height = cmdButton(0).Height + Me.TextHeight(ButtonCaptions(0)) + 1
  Loop
  cmdButton(0).Width = LargestButtonWidth
  
  ' How wide is the caption?
  CaptionWidth = Me.TextWidth(Prompt)
  lblCaption.AutoSize = False
  
  ' Width of the form and size of the label
  If Img = IDI_NONE Then
    lblCaption.Left = 10
    If CaptionWidth + 20 > MaxMsgBoxWidth Then
      lblCaption.Width = MaxMsgBoxWidth - 20
      lblCaption.Height = ((Me.TextHeight(Prompt) + 2) * (CaptionWidth / (MaxMsgBoxWidth - 50))) + 10
      Me.Width = Screen.TwipsPerPixelX * MaxMsgBoxWidth
    Else
      If CaptionWidth + 20 > (LargestButtonWidth + 20) * (UBound(ButtonCaptions) + 1) Then
        Me.Width = Screen.TwipsPerPixelX * (CaptionWidth + 20)
        lblCaption.Width = CaptionWidth
        lblCaption.Height = Me.TextHeight(Prompt) + 5
      Else
        Me.Width = Screen.TwipsPerPixelX * ((LargestButtonWidth + 20) * (UBound(ButtonCaptions) + 1))
        lblCaption.Width = ((LargestButtonWidth + 20) * (UBound(ButtonCaptions) + 1) - 20)
        lblCaption.Height = Me.TextHeight(Prompt) + 5
      End If
    End If
  Else
    lblCaption.Left = 52
    If CaptionWidth + 62 > MaxMsgBoxWidth Then
      lblCaption.Width = MaxMsgBoxWidth - 62
      Me.Width = Screen.TwipsPerPixelX * MaxMsgBoxWidth
      lblCaption.Height = ((Me.TextHeight(Prompt) + 2) * (CaptionWidth / (MaxMsgBoxWidth - 50))) + 10
    Else
      If CaptionWidth + 62 > (LargestButtonWidth + 20) * (UBound(ButtonCaptions) + 1) Then
        Me.Width = Screen.TwipsPerPixelX * (CaptionWidth + 62)
        lblCaption.Width = CaptionWidth
        lblCaption.Height = Me.TextHeight(Prompt) + 5
        lblCaption.Top = 16
      Else
        Me.Width = Screen.TwipsPerPixelX * ((LargestButtonWidth + 20) * (UBound(ButtonCaptions) + 1))
        lblCaption.Width = Screen.TwipsPerPixelX * (CaptionWidth + 62)
        lblCaption.Height = Me.TextHeight(Prompt) + 5
        lblCaption.Top = 16
      End If
    End If
  End If
  
  ' Form height
  lblCaption.Caption = Prompt
  If lblCaption.Height < 32 And Img <> IDI_NONE Then
    cmdButton(0).Top = 52
    Me.Height = 0
    Me.Height = Me.Height + (Screen.TwipsPerPixelY * (cmdButton(0).Height + 62))
  Else
    cmdButton(0).Top = lblCaption.Top + lblCaption.Height + 10
    Me.Height = 0
    Me.Height = Me.Height + (Screen.TwipsPerPixelY * (cmdButton(0).Height + lblCaption.Height + 30))
  End If

  ' Load icon
  If Img <> IDI_NONE Then
    hIcon = LoadStandardIcon(0&, Img)
  End If
  
  ' Disable close on form
  hSysMenu = GetSystemMenu(Me.hwnd, False)
  RemoveMenu hSysMenu, SC_CLOSE, MF_BYCOMMAND
  
  
  ' Set up the buttons
  i = 0
  For Each tmpVar In ButtonCaptions
    If i = 0 Then
      cmdButton(i).Caption = tmpVar
      If UBound(ButtonCaptions) = 0 Then
        ' only one button, so centre the button
        cmdButton(i).Left = (Me.ScaleWidth / 2) - (cmdButton(i).Width / 2)
      End If
    Else
      Load cmdButton(i)
      cmdButton(i).Caption = tmpVar
      cmdButton(i).Left = (Me.ScaleWidth - 10 - cmdButton(0).Width) * i / UBound(ButtonCaptions)
      cmdButton(i).Top = cmdButton(i - 1).Top
      cmdButton(i).Width = cmdButton(i - 1).Width
      cmdButton(i).Height = cmdButton(i - 1).Height
      If tmpVar = "Cancel" Then
        cmdButton(i).Cancel = True
      End If
      If tmpVar = "Ok" Or tmpVar = "Yes" Then
        cmdButton(i).Default = True
      End If
      cmdButton(i).Visible = True
    End If
    i = i + 1
  Next
  
  ' Show the form
  ClickedButton = ""
  Me.Show vbModal
  
  ' Clean up
  hIcon = 0
  
  CustomMessageBox = ClickedButton
End Function

Private Sub cmdButton_Click(Index As Integer)
  ClickedButton = cmdButton(Index).Caption
  Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim tmpStyle As VbMsgBoxStyle
  If Button = vbLeftButton And Shift = (vbShiftMask Or vbCtrlMask) Then
    Select Case mImg
      Case CMBStandardIconEnum.IDI_ASTERISK
        tmpStyle = vbInformation
      Case CMBStandardIconEnum.IDI_EXCLAMATION
        tmpStyle = vbExclamation
      Case CMBStandardIconEnum.IDI_HAND
        tmpStyle = vbCritical
      Case CMBStandardIconEnum.IDI_NONE
        tmpStyle = 0
      Case CMBStandardIconEnum.IDI_QUESTION
        tmpStyle = vbQuestion
      Case Else
        tmpStyle = 0
    End Select
    MsgBox lblCaption.Caption, tmpStyle, Me.Caption
  End If
End Sub

Private Sub Form_Paint()
    If hIcon <> 0 Then
      Call DrawIcon(Me.hDC, 10&, 10&, hIcon)
    End If
End Sub
