VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "API Text Wrapper Demo"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9135
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   6960
      TabIndex        =   7
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4680
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Normal:"
      Height          =   195
      Left            =   6960
      TabIndex        =   8
      Top             =   120
      Width           =   540
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   8895
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Lowercase:"
      Height          =   195
      Left            =   4680
      TabIndex        =   5
      Top             =   120
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Uppercase:"
      Height          =   195
      Left            =   2400
      TabIndex        =   4
      Top             =   120
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Numbers Only:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1035
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Text1DefaultSyle As Long
Dim Text2DefaultSyle As Long
Dim Text3DefaultSyle As Long

Private Sub Form_Load()
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

Text1DefaultSyle = NumbersOnly(Text1)
Text2DefaultSyle = UpperCaseOnly(Text2)
Text3DefaultSyle = LowerCaseOnly(Text3)


End Sub


Private Sub Text1_GotFocus()
Dim LabelText  As String
Label4.Caption = ""

LabelText = "TextBox1:" & vbCrLf
LabelText = LabelText & "Style Number: " & GetStyle(Text1) & vbCrLf
LabelText = LabelText & "Style Name: " & StyleNumberToText(Text1) & vbCrLf
Label4.Caption = LabelText


End Sub


Private Sub Text2_GotFocus()
Dim LabelText  As String
Label4.Caption = ""

LabelText = "TextBox2:" & vbCrLf
LabelText = LabelText & "Style Number: " & GetStyle(Text2) & vbCrLf
LabelText = LabelText & "Style Name: " & StyleNumberToText(Text2) & vbCrLf
Label4.Caption = LabelText
End Sub


Private Sub Text3_GotFocus()
Dim LabelText  As String
Label4.Caption = ""

LabelText = "TextBox3:" & vbCrLf
LabelText = LabelText & "Style Number: " & GetStyle(Text3) & vbCrLf
LabelText = LabelText & "Style Name: " & StyleNumberToText(Text3) & vbCrLf
Label4.Caption = LabelText
End Sub


Private Sub Text4_GotFocus()
Dim LabelText  As String
Label4.Caption = ""

LabelText = "TextBox4:" & vbCrLf
LabelText = LabelText & "Style Number: " & GetStyle(Text4) & vbCrLf
LabelText = LabelText & "Style Name: " & StyleNumberToText(Text4) & vbCrLf
Label4.Caption = LabelText
End Sub


