VERSION 5.00
Begin VB.Form simulation 
   Caption         =   "Main Screen"
   ClientHeight    =   3195
   ClientLeft      =   5940
   ClientTop       =   4395
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Pause"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   26
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   240
   End
   Begin VB.PictureBox partition 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7500
      Left            =   10920
      ScaleHeight     =   7440
      ScaleWidth      =   675
      TabIndex        =   24
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5160
      TabIndex        =   23
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3240
      TabIndex        =   22
      Top             =   6120
      Width           =   4935
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5160
      TabIndex        =   21
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5160
      TabIndex        =   20
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5160
      TabIndex        =   19
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5160
      TabIndex        =   18
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5160
      TabIndex        =   17
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2280
      TabIndex        =   16
      Top             =   2280
      Width           =   5895
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2280
      TabIndex        =   15
      Top             =   1800
      Width           =   5895
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2280
      TabIndex        =   14
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   1
      Top             =   7440
      Width           =   1215
   End
   Begin VB.PictureBox Memorymap 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7500
      Left            =   8640
      ScaleHeight     =   7440
      ScaleWidth      =   1935
      TabIndex        =   0
      Top             =   480
      Width           =   2000
   End
   Begin VB.Label Label14 
      Caption         =   "/ 1000 sec"
      Height          =   375
      Left            =   3120
      TabIndex        =   27
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "Partition"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10920
      TabIndex        =   25
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label12 
      Caption         =   "Storage placement policy:"
      Height          =   375
      Left            =   840
      TabIndex        =   13
      Top             =   6720
      Width           =   2895
   End
   Begin VB.Label Label11 
      Caption         =   "Compaction policy:"
      Height          =   375
      Left            =   840
      TabIndex        =   12
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Label Label10 
      Caption         =   "Number of processes sucessfully run:"
      Height          =   375
      Left            =   840
      TabIndex        =   11
      Top             =   5520
      Width           =   3855
   End
   Begin VB.Label Label9 
      Caption         =   "Number of active processes:"
      Height          =   375
      Left            =   840
      TabIndex        =   10
      Top             =   4920
      Width           =   3735
   End
   Begin VB.Label Label8 
      Caption         =   "No. of processes awaiting execution:"
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Top             =   4320
      Width           =   3975
   End
   Begin VB.Label Label7 
      Caption         =   "Number of processes dropped:"
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   3720
      Width           =   3375
   End
   Begin VB.Label Label6 
      Caption         =   "Current fragmentation:"
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "Action:"
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Occurence:"
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Time:"
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Memory Management Simulation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3120
      TabIndex        =   3
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Memory Map"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "simulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem Code to display memory map for OS simulation
Rem M. Russell April 2000, HND Group 2, year 1
Option Explicit
Dim step As Double
Dim entry As Integer
Dim result As Integer
Dim p_left, dropped, active, finished As Integer
Dim finish(1 To 26) As Integer
Dim memlocs As Integer
Dim total_height As Integer
Dim total_width As Integer
Dim colour(0 To 26) As Long

Rem This code displays the memory map in the picture box
Private Sub Memory_map()
Dim block As Double
Dim xpos As Integer
Dim ypos As Integer
Dim n As Integer
Memorymap.Cls
xpos = 0
ypos = total_height
block = step
For n = 4 To 20
If memory(n) <> memory(n - 1) Then
    Memorymap.Line (xpos, ypos)-Step(total_width - 800, -block), colour(memory(n - 1)), BF
    Memorymap.Print ((n - 1) * 50)
    Rem Print process letter
    ypos = ypos - block / 2
    If memory(n - 1) <> 0 Then
        Memorymap.Print Chr$(64 + memory(n - 1))
    End If
    ypos = ypos - block / 2
    block = step
Else
    block = block + step
End If
Next n
Memorymap.Line (xpos, ypos)-Step(total_width - 800, -block), colour(memory(n - 1)), BF
Memorymap.Print ((n - 1) * 50)
ypos = ypos - block / 2
If memory(n - 1) <> 0 Then
    Memorymap.Print Chr$(64 + memory(n - 1))
End If
End Sub

Rem Exit simulation button selected
Private Sub Command1_Click()
Unload Me
End Sub

Rem Pause button selected. This is a 'toggle' button
Private Sub Command2_Click()
If Command2.Caption = "Resume" Then
    Command2.Caption = "Pause"
    Timer1.Enabled = True
Else
    Timer1.Enabled = False
    Command2.Caption = "Resume"
End If
End Sub

Rem When form loaded, reset screen
Private Sub Form_Activate()
Timer1.Enabled = False  ' Turn timer off
Initialise  ' Setup variables etc
Memory_map  ' Display memory map
If fixed = True Then    ' If fixed partition, display partition list
    n = 1
    Do Until part(n) = 1000
    partition.Line (0, total_height - (((part(n) / 50 - 2) * step)))-Step(300, 0)
    partition.Print (part(n))
    n = n + 1
    Loop
End If
Timer1.Enabled = True   ' Turn timer on
End Sub

Rem Setup variables for new run
Public Sub Initialise()
Load report_form    ' loaded so can be updated, but not displayed yet
total_height = 7500 ' Height of memory map picturebox
total_width = 2000  ' width of memory map picturebox
Memorymap.Height = total_height
Memorymap.Width = total_width
memlocs = 18
step = total_height / memlocs
Command2.Caption = "Pause"
' Initialise report form textbox
With report_form.Text1
    .Text = "Simulation run on " & Date & " at " & Time & new_line & new_line
    .Text = report_form.Text1.Text & "Partition type: "
    If fixed = True Then
        .Text = .Text & "Fixed, at location(s) "
        j = 1
        Do Until part(j) = 1000
        If part(j + 1) <> 1000 Then
            .Text = .Text & part(j) & ", "
        Else
            .Text = .Text & part(j) & "."
        End If
        j = j + 1
        Loop
    Else
        .Text = .Text & "Variable"
    End If
    .Text = .Text & new_line
    .Text = .Text & "Compaction policy: "
End With
Select Case compaction
Case 0
    If fixed = False Then
        report_form.Text1.Text = report_form.Text1.Text & "None" & new_line
            Text9.Text = "None"
    Else
        report_form.Text1.Text = report_form.Text1.Text & "N/A" & new_line
            Text9.Text = "N/A"
    End If
Case 1
    Text9.Text = "Compact when process finishes"
    report_form.Text1.Text = report_form.Text1.Text & "When process finishes" & new_line
Case 2
    Text9.Text = "Compact when process unable to load due to fragmentation"
    report_form.Text1.Text = report_form.Text1.Text & "When process unable to load due to fragmentation" & new_line
Case Else
    Text9.Text = "** ERROR **"
End Select
report_form.Text1.Text = report_form.Text1.Text & "Placement policy: "
Select Case policy
Case 0
    Text10.Text = "First fit"
    report_form.Text1.Text = report_form.Text1.Text & "First fit" & new_line
Case 1
    Text10.Text = "Best fit"
    report_form.Text1.Text = report_form.Text1.Text & "Best fit" & new_line
Case 2
    Text10.Text = "Worse fit"
    report_form.Text1.Text = report_form.Text1.Text & "Worse fit" & new_line
Case Else
    Text10.Text = "** ERROR **"
End Select
report_form.Text1.Text = report_form.Text1.Text & new_line & new_line
report_form.Text1.Text = report_form.Text1.Text & "No. frag clock Event" & new_line
For n = 1 To process
finish(n) = run(n) + arrival(n)
Next n
' Initialise variables
p_left = process
dropped = 0
entry = 0
active = 0
finished = 0
clock = arrival(1)  ' Set clock to start time of earliest process
For n = 3 To 20
memory(n) = 0
Next n
comp_count = 0
Rem Need 26 unique colours, 1 for each process
colour(0) = RGB(255, 255, 255)
colour(1) = RGB(0, 255, 255)
colour(2) = RGB(85, 85, 0)
colour(3) = RGB(255, 0, 0)
colour(4) = RGB(0, 125, 0)
colour(5) = RGB(0, 255, 0)
colour(6) = RGB(50, 50, 125)
colour(7) = RGB(0, 125, 255)
colour(8) = RGB(125, 125, 0)
colour(9) = RGB(255, 125, 0)
colour(10) = RGB(125, 125, 125)
colour(11) = RGB(125, 125, 255)
colour(12) = RGB(255, 85, 255)
colour(13) = RGB(255, 255, 0)
colour(14) = RGB(0, 0, 255)
colour(15) = RGB(125, 0, 125)
colour(16) = RGB(255, 0, 255)
colour(17) = RGB(170, 170, 0)
colour(18) = RGB(50, 85, 0)
colour(19) = RGB(0, 85, 170)
colour(20) = RGB(0, 170, 85)
colour(21) = RGB(125, 0, 75)
colour(22) = RGB(170, 85, 0)
colour(23) = RGB(125, 0, 0)
colour(24) = RGB(0, 85, 0)
colour(25) = RGB(170, 0, 0)
colour(26) = RGB(170, 170, 170)
End Sub

Rem This code is executed every time the clock pulses
Private Sub Timer1_Timer()
Dim i, mem_temp As Integer
Dim comp As Boolean
Text1.Text = Right$("0000" & clock, 4)
If p_left = 0 Then  ' Exit if no more processes left
    Timer1.Enabled = False
    With report_form.Text1
        .Text = .Text & new_line & new_line
        .Text = .Text & "Number of processes successfully run: " & finished & new_line
        .Text = .Text & "Number of processes dropped: " & dropped & new_line
        .Text = .Text & "Maximum fragmentation: " & frag_max & "% at event " & Right$("000" & frag_clock, 4) & new_line
        If comp_count <> 0 Then .Text = .Text & "Number of times memory compacted: " & comp_count
    End With
    report_form.Text1.Enabled = False   ' Ensure user cannot amend before displaying
    report_form.Show
    Unload Me   ' Simulation exits here when returning from report_form
End If
For n = 1 To process
' First check for processes finishing
If finish(n) = clock And run(n) > 0 Then
    result = 1
    p_left = p_left - 1
    finished = finished + 1
    active = active - 1
    For i = 3 To 20
    If memory(i) = n Then memory(i) = 0
    Next i
    Call Update("Process " & Chr$(64 + n) & " has terminated.", "Memory freed for re-allocation.")
    If compaction = 1 Then
        If compact = True Then
            Call Update("Process finished.", "Memory compacted")
            comp_count = comp_count + 1
        End If
    End If
End If
' Now check to see if any processes waiting to start
    If arrival(n) = clock And run(n) > 0 Then
        If compaction = 2 Then
            comp = True ' This variable ensures that compaction only called once
        Else            ' can't load process due to fragmentation
            comp = False
        End If
        Do
            result = place(size(n))
            If result = 0 And comp = True Then
                comp = compact
                Call Update("Process " & Chr$(64 + n) & ", " & size(n) & "k, due to start, but has placement problem.", "Memory compacted")
                comp_count = comp_count + 1
            Else
                comp = False
            End If
        Loop Until comp = False
        If result = 0 Then
            finish(n) = 0
            dropped = dropped + 1
            p_left = p_left - 1
            Call Update("Process " & Chr$(64 + n) & ", " & size(n) & "k, due to start.", "Dropped - no room in memory. ")
        Else
            mem_temp = result
            result = result / 50 + 1
            For i = 1 To size(n) / 50
                memory(result) = n
                result = result + 1
            Next i
            active = active + 1
            Call Update("Process " & Chr$(64 + n) & ", " & size(n) & "k, due to start.", "Allocated to memory starting at location " & mem_temp)
        End If
    End If
Next n
clock = clock + 1 ' Increment clock on exiting subroutine
End Sub

Rem This function is given a process size and returns with a place in memory or null
Public Function place(size As Integer) As Integer
Dim diff, held, pointer, location As Integer
place = 0
If policy = 1 Then held = 9999
If fixed = True Then
    Do
        If memory((part(pointer) / 50) + 1) = 0 Then
            location = part(pointer + 1) - part(pointer)
            If location >= size Then
                Select Case policy
                Case 0
                    If held = 0 Then
                        place = part(pointer)
                        held = 1
                    End If
                Case 1
                    diff = location - size
                    If diff <= held Then
                        place = part(pointer)
                        held = diff
                    End If
                Case 2
                    diff = location - size
                    If diff >= held Then
                        place = part(pointer)
                        held = diff
                    End If
                End Select
            End If
        End If
        pointer = pointer + 1
    Loop Until part(pointer) = 1000
    Exit Function
Else    ' Variable partition
    pointer = 3
    Do
        Do Until pointer = 21
            If memory(pointer) = 0 Then Exit Do
            pointer = pointer + 1
        Loop
        If pointer = 21 Then Exit Function
        location = pointer
        Do Until location >= 20
            If memory(location) <> 0 Then Exit Do
            location = location + 1
        Loop
        If location = 20 Then location = 21
        If size <= (location - pointer) * 50 Then
                Select Case policy
                Case 0
                    If held = 0 Then
                        place = (pointer - 1) * 50
                        held = 1
                    End If
                Case 1
                    diff = (location - pointer) * 50 - size
                    If diff <= held Then
                        place = (pointer - 1) * 50
                        held = diff
                    End If
                Case 2
                    diff = (location - pointer) * 50 - size
                    If diff >= held Then
                        place = (pointer - 1) * 50
                        held = diff
                    End If
                End Select
        End If
        pointer = location
        Loop Until pointer >= 20
End If
End Function

Rem Compaction function. Returns true if compaction took place
Public Function compact() As Boolean
Dim free, nxt As Integer
free = 3 ' Set to bottom of memory
compact = True
Do While memory(free) <> 0 And free < 20
    free = free + 1
Loop
If free = 20 Then
    compact = False
    Exit Function
End If
nxt = free + 1
Do While memory(nxt) = 0 And nxt < 20
    nxt = nxt + 1
Loop
If nxt = 20 And memory(nxt) = 0 Then
    compact = False
    Exit Function
End If
Do
    memory(free) = memory(nxt)
    memory(nxt) = 0
    free = free + 1
    Do While memory(nxt) = 0 And nxt < 20
        nxt = nxt + 1
    Loop
    If nxt = 20 And memory(nxt) <> 0 Then
        memory(free) = memory(nxt)
        memory(nxt) = 0
    End If
Loop Until nxt = 20
End Function

Rem An even has happened. Update the screen with new parameters
Public Sub Update(occur As String, action As String)
Dim frag As Integer
Dim free_locs, total_locs, found As Integer
Memory_map
Text2.Text = occur
Text3.Text = action
' Calculate fragmentation
frag = 0
free_locs = 0
total_locs = 0
found = 0
If fixed = True Then
    i = 0
    Do Until part(i) = 1000
        If memory(part(i) / 50 + 1) > 0 Then
            For j = part(i + 1) / 50 To part(i) / 50 + 1 Step -1
                If memory(j) = 0 Then free_locs = free_locs + 1
                total_locs = total_locs + 1
            Next j
        End If
    i = i + 1
    Loop
    If free_locs <> 0 Then
        frag = Int((free_locs / total_locs) * 10000) / 100
    End If
Else    'Variable partition
    i = 20
    Do
        If memory(i) > 0 Then found = i
        i = i - 1
    Loop Until i = 2 Or found <> 0
    i = 0
    j = found
    If j > 3 Then
        Do
            If memory(j) = 0 Then i = j
            j = j - 1
        Loop Until j = 2 Or i <> 0
        If i <> 0 Then
            i = 0
            For j = found To 3 Step -1
                If memory(j) = 0 Then i = i + 1
            Next j
            frag = Int(i / (found - 2) * 10000) / 100
        End If
    End If
End If
Text4.Text = frag & "%"
If frag > frag_max Then
    frag_max = frag
    frag_clock = entry
End If
Text5.Text = dropped
Text6.Text = p_left - active
Text7.Text = active
Text8.Text = finished
' Update run log
report_form.Text1.Text = report_form.Text1.Text & Right$("00" & entry, 3) & "  " & Right$("00" & frag, 2) & "%  " & Right$("000" & clock, 4) & " " & occur & " " & action & new_line
entry = entry + 1
If halt = True Then MsgBox ("Halted...")
End Sub
