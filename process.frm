VERSION 5.00
Begin VB.Form process_setup 
   Caption         =   "Memory Management Simulation"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "clear list"
      Height          =   375
      Left            =   3480
      TabIndex        =   15
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4440
      TabIndex        =   12
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit to main menu"
      Height          =   615
      Left            =   2400
      TabIndex        =   3
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "delete"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton add 
      Caption         =   "add"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3960
      ItemData        =   "process.frx":0000
      Left            =   240
      List            =   "process.frx":0002
      TabIndex        =   0
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Process Setup Screen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   14
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "process"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "runtime"
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "arrival time"
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "size"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "runtime"
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "arrival time"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "size"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   1080
      Width           =   375
   End
End
Attribute VB_Name = "process_setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem Code to input list of processes for OS Simulation
Rem M. Russell April 2000, Group 2, year 1
Dim position As Integer
Option Explicit

Rem Add a new process to the list
Private Sub add_Click()
Dim newsize, newarrival, newrun As Integer
If process >= 26 Then
    MsgBox ("Sorry, can't have more than 26 processes")
    Exit Sub
End If
newsize = Val(Text1.Text)
newarrival = Val(Text2.Text)
If Val(Text3.Text) < 1 Or Val(Text3.Text) > 9999 Then
    MsgBox ("Error in run time")
    Exit Sub
End If
newrun = Val(Text3.Text)
If newsize < 50 Or newsize > 900 Then
    MsgBox ("Error in value of process size")
    Exit Sub
End If
If newarrival < 0 Or newarrival > 9999 Then
    MsgBox ("Error in arrival time")
    Exit Sub
End If
If newarrival + newrun > 9999 Then
    MsgBox ("Process will exceed maximum clock setting")
    Exit Sub
End If
If newsize / 50 <> Int(newsize / 50) Then
    MsgBox ("Process must be a multiple of 50k")
    Exit Sub
End If
Rem Insert new process into list
Select Case process
Case 0
    size(1) = newsize
    arrival(1) = newarrival
    run(1) = newrun
Case 1
    If newarrival < arrival(1) Then
        arrival(2) = arrival(1)
        run(2) = run(1)
        size(2) = size(1)
        size(1) = newsize
        arrival(1) = newarrival
        run(1) = newrun
    Else
        arrival(2) = newarrival
        run(2) = newrun
        size(2) = newsize
    End If
Case Else  'Must be in middle or at end of list
    position = process + 1
    For n = 1 To process
        If arrival(n) > newarrival Then
            position = n
            n = process + 1
        End If
    Next n
    For n = process To position Step -1
        size(n + 1) = size(n)
        arrival(n + 1) = arrival(n)
        run(n + 1) = run(n)
    Next n
    size(position) = newsize
    arrival(position) = newarrival
    run(position) = newrun
End Select
process = process + 1
printlist
End Sub

Rem Clear entire list
Private Sub Command1_Click()
process = 0
printlist
End Sub

Rem Delete the selected list item
Private Sub Command2_Click()
position = List1.ListIndex + 1
If position = 0 Then Exit Sub
If position <> process Then
    For n = position + 1 To process
        size(n - 1) = size(n)
        arrival(n - 1) = arrival(n)
        run(n - 1) = run(n)
    Next n
End If
If process > 0 Then process = process - 1
printlist
End Sub

Rem Subroutine to update the List box with the processes
Public Sub printlist()
List1.Clear
For n = 1 To process
    List1.AddItem " " & Chr$(64 + n) & "    " & Right$("  " & size(n), 3) & "   " & Right$("   " & arrival(n), 4) & "   " & Right$("   " & run(n), 4)
Next n
End Sub

Private Sub Form_Load()
printlist
End Sub

Rem Exit to main menu
Private Sub Command3_Click()
Unload Me
End Sub

Rem Return listbox item to data entry boxes
Private Sub List1_Click()
position = List1.ListIndex + 1
If position <> 0 Then
    Text1.Text = size(position)
    Text2.Text = arrival(position)
    Text3.Text = run(position)
End If
End Sub
