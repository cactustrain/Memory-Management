VERSION 5.00
Begin VB.Form Help_Form 
   Caption         =   "Memory Management Simulation"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   11655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Return to Main Menu"
      Height          =   615
      Left            =   4680
      TabIndex        =   0
      Top             =   7560
      Width           =   2175
   End
End
Attribute VB_Name = "Help_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem Help screen code
Rem Written by M. Russell April-May 2000
Option Explicit

Rem Exit selected
Private Sub Command1_Click()
Unload Me
End Sub

Rem On start up, enter text into textbox
Private Sub Form_Load()
With Text1
    .Text = "This program simulates how an operating system copes with memory allocation in a multitasking environment. "
    .Text = .Text & "Normally, processes start at random intervals and are of a random length. "
    .Text = .Text & "However, so that you can test out the various systems available, you can set up the process list yourself. "
    .Text = .Text & "Up to 26 processes can be set-up. "
    .Text = .Text & "It is recommended that you use the same process list and alter the other parameters so that you can compare the efficiency of the systems. "
    .Text = .Text & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10)
    .Text = .Text & "You can choose either a fixed partition or a variable partition arrangement for the memory. "
    .Text = .Text & "Paged, segmented and virtual memory systems are not covered by this simulation. "
    .Text = .Text & "Locations 0 to 100k are occupied by the operating system itself, and so locations 100 to 1000k are available for process allocation. "
    .Text = .Text & "Processes must be in multiples of 50k. "
    .Text = .Text & "The processes arrive at the start time that you specify, but if you enter two or more with the same start time, then the order that you enter them takes precedence. "
    .Text = .Text & "When the simulation commences, the clock is set to the earliest start time in your process list. "
    .Text = .Text & "The clock is incremented about once every second and you should bear this in mind when setting up the process list. "
    .Text = .Text & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10)
    .Text = .Text & "Compaction is a technique available in variable partition systems whereby processes can be moved in memory to free up unused areas. "
    .Text = .Text & "This system is efficient, but does necessitate the processor spending time relocating code, along with all the difficulties that that entails. "
    .Text = .Text & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10)
    .Text = .Text & "You should consult the User Guide for more detailed instructions."
    .Text = .Text & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10)
    .Text = .Text & "If you wish to know more about this subject, then I recommend you read chapter 5 of 'Operating Systems', by Colin Ritchie and published by Letts. (See chapter 2 regarding the problems of memory relocation)."
End With
End Sub
