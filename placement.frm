VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   1920
      TabIndex        =   10
      Top             =   3240
      Width           =   4815
      Begin VB.OptionButton Option3 
         Caption         =   "Worst"
         Height          =   375
         Left            =   720
         TabIndex        =   13
         Top             =   1680
         Width           =   3615
      End
      Begin VB.OptionButton Option2 
         Caption         =   "best"
         Height          =   375
         Left            =   720
         TabIndex        =   12
         Top             =   1080
         Width           =   3135
      End
      Begin VB.OptionButton Option1 
         Caption         =   "first"
         Height          =   375
         Left            =   720
         TabIndex        =   11
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   7680
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      Top             =   5880
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "find location"
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   6840
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   1800
      TabIndex        =   4
      Top             =   2040
      Width           =   5535
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   1920
      TabIndex        =   0
      Top             =   1080
      Width           =   6495
   End
   Begin VB.Label Label5 
      Caption         =   "placement policy"
      Height          =   735
      Left            =   240
      TabIndex        =   14
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "process size"
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Partitions at"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Simulation type"
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "memory map"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem Test for code which decides upon process placement
Rem M. Russell April 2000, HND group 2, year 1


Private Sub Command1_Click()
If Val(Text4) < 1 Then
    MsgBox ("Process size incorrect")
    Exit Sub
End If
If Val(Text4) / 50 <> Int(Val(Text4) / 50) Then
    MsgBox ("Process size incorrect")
    Exit Sub
End If
If Option1 = True Then policy = 0 'first fit
If Option2 = True Then policy = 1 ' best fit
If Option3 = True Then policy = 2 'worst fit
result = place(Val(Text4))
If result = 0 Then
    MsgBox ("Placement failed")
Else
    MsgBox ("Process Allocation to" & result)
End If
Text4.SetFocus
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
Form1.Show
Text4.SetFocus
Option1 = True
fixed = True
memory(3) = 1
memory(4) = 1
memory(6) = 4
memory(10) = 10
memory(15) = 15
If fixed = True Then
    part(0) = 100
    part(1) = 300
    part(2) = 600
    part(3) = 650
    part(4) = 800
    part(5) = 900
    part(6) = 1000
End If
Rem Display Values
For n = 3 To 20
Text1.Text = Text1.Text & Val(memory(n)) & ", "
Next n
If fixed = True Then
    For n = 0 To 6
    Text3.Text = Text3.Text & Val(part(n)) & ","
    Next n
End If
If fixed = False Then
    Text2.Text = "Variable"
Else
    Text2.Text = "Fixed"
End If
End Sub


Private Sub Option1_Click()
policy = 0
End Sub

Private Sub Option2_Click()
    policy = 1
End Sub

Private Sub Option3_Click()
policy = 3
End Sub
