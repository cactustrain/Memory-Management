VERSION 5.00
Begin VB.Form amend_params 
   Caption         =   "Memory Management Simulation"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6885
   LinkTopic       =   "Form2"
   ScaleHeight     =   6795
   ScaleWidth      =   6885
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   720
      TabIndex        =   13
      Top             =   3000
      Width           =   5655
      Begin VB.CheckBox Check1 
         Caption         =   "Compaction"
         Height          =   375
         Left            =   480
         TabIndex        =   16
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton Option7 
         Caption         =   "When process cannot load due to fragmentation"
         Height          =   375
         Left            =   2280
         TabIndex        =   15
         Top             =   600
         Width           =   3015
      End
      Begin VB.OptionButton Option6 
         Caption         =   "When process finishes"
         Height          =   255
         Left            =   2280
         TabIndex        =   14
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   720
      TabIndex        =   10
      Top             =   4200
      Width           =   5655
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   4
         Left            =   4800
         TabIndex        =   22
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   3
         Left            =   4080
         TabIndex        =   21
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   2
         Left            =   3360
         TabIndex        =   20
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   1
         Left            =   2640
         TabIndex        =   19
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   0
         Left            =   1920
         TabIndex        =   11
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Partition locations  (multiples of 50k)"
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Main Menu"
      Height          =   615
      Left            =   2400
      TabIndex        =   9
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   3840
      TabIndex        =   4
      Top             =   720
      Width           =   2055
      Begin VB.OptionButton Option5 
         Caption         =   "Worst fit"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Best Fit"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "First Fit"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Placement Policy"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   2295
      Begin VB.OptionButton Option2 
         Caption         =   "Variable Partition"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Fixed Partition"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame5 
      Height          =   6120
      Left            =   360
      TabIndex        =   17
      Top             =   480
      Width           =   6375
      Begin VB.CheckBox Check2 
         Caption         =   "Halt simulation after each event"
         Height          =   255
         Left            =   1800
         TabIndex        =   18
         Top             =   4800
         Width           =   2775
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Amend Parameters"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "amend_params"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem Code for Amend Parameters screen
Rem Written by M. Russell April-May 2000
Option Explicit
Dim smallest, temp, position, pos As Integer

Rem Toggle Compaction selection
Private Sub Check1_Click()
If Check1.Value = 0 Then
    Option6.Enabled = False
    Option7.Enabled = False
Else
    Option6.Enabled = True
    Option7.Enabled = True
End If
End Sub

Rem Exit to main menu option selected
Private Sub Command1_Click()
If Option1.Value = True Then
    fixed = True
    temp = 0
    For n = 0 To 4
    If Int(Val(Text1(n)) / 50) <> Val(Text1(n)) / 50 Then
        MsgBox ("Partition sizes must be in multiples of 50k")
        Exit Sub
    End If
    If Val(Text1(n)) > 950 Then
        MsgBox ("Partition too large")
        Exit Sub
    End If
    Next n
    For n = 0 To 4
    If Val(Text1(n)) > 100 Then
        temp = temp + 1
    End If
    Next n
    If temp = 0 Then
        MsgBox ("Correct values needed for partition sizes")
        Exit Sub
    Else
        'Assign values to partition array, may be in any order
        position = 1
        Do
            smallest = 999
            For n = 0 To 4
            If Val(Text1(n)) > 100 And Val(Text1(n)) < smallest Then
                smallest = Val(Text1(n))
                pos = n
            End If
            Next n
            Text1(pos).Text = ""
            part(position) = smallest
            temp = temp - 1
            position = position + 1
        Loop Until temp = 0
        part(position) = 1000
        part(0) = 100
    End If
Else
    fixed = False
End If
If Check2.Value = 1 Then
    halt = True
Else
    halt = False
End If
If Option3.Value = True Then policy = 0
If Option4.Value = True Then policy = 1
If Option5.Value = True Then policy = 2
If Check1.Value = 1 And Check1.Enabled = True And Option6.Value = False And Option7.Value = False Then
    MsgBox ("Please enter a compaction method")
    Exit Sub
End If
If Check1.Value = 0 Or Option1.Value = True Then
    compaction = 0
Else
    If Option6.Value = True Then
        compaction = 1
    Else
        compaction = 2
    End If
End If
Unload Me
End Sub

Rem Restore values to screen
Private Sub Form_Load()
For n = 1 To 5
If part(n) < 1000 And part(n) > 0 Then Text1(n - 1).Text = part(n)
Next n
If fixed = True Then
    Option1.Value = True
    Option2.Value = False
Else
    Option1.Value = False
    Option2.Value = True
End If
Select Case policy
Case 0
    Option3.Value = True
    Option4.Value = False
    Option5.Value = False
Case 1
    Option3.Value = False
    Option4.Value = True
    Option5.Value = False
Case 2
    Option3.Value = False
    Option4.Value = False
    Option5.Value = True
End Select
Select Case compaction
Case 0
    Check1.Value = 0
    Option6.Enabled = False
    Option7.Enabled = False
Case 1
    Check1.Value = 1
    Option6.Value = True
    Option7.Value = False
Case 2
    Check1.Value = 1
    Option6.Value = False
    Option7.Value = True
End Select
If halt = True Then
    Check2.Value = 1
Else
    Check2.Value = 0
End If

End Sub

Rem Fixed Partition selected
Private Sub Option1_Click()
Check1.Enabled = False
Option6.Enabled = False
Option7.Enabled = False
Label3.Enabled = True
For n = 0 To 4
Text1(n).Enabled = True
Next n
End Sub

Rem Variable Partition selected
Private Sub Option2_Click()
Check1.Enabled = True
If Check1.Value = 1 Then
    Option6.Enabled = True
    Option7.Enabled = True
End If
Label3.Enabled = False
For n = 0 To 4
Text1(n).Enabled = False
Next n
End Sub
