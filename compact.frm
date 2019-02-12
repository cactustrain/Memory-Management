VERSION 5.00
Begin VB.Form temp 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "quit"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "compact"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "show memory"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   1575
      Left            =   960
      TabIndex        =   0
      Top             =   960
      Width           =   6135
   End
End
Attribute VB_Name = "temp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem Code to compact memory for OS Memory Management Simulation
Rem M. Russell April 2000, HND group, year 2
Option Explicit
Dim result As Boolean

Private Sub Command1_Click()
Text1.Text = ""
For n = 3 To 20
If memory(n) = 0 Then
    Text1.Text = Text1.Text & "0,"
Else
    Text1.Text = Text1.Text & memory(n) & ","
End If
Next n
End Sub

Private Sub Command2_Click()
result = compact()
If result = False Then
    MsgBox ("compaction did not take place")
Else
    MsgBox ("compaction took place")
End If
End Sub

Private Sub Command3_Click()
End
End Sub


Private Sub Form_Load()
memory(4) = 2
memory(5) = 2
memory(6) = 2
memory(8) = 3
memory(20) = 4
End Sub
