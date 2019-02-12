VERSION 5.00
Begin VB.Form Splash 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Memory Management Simulation"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   1935
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enter"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Operating System Memory Management Simulation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   840
      TabIndex        =   1
      Top             =   3120
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   2805
      Left            =   240
      Picture         =   "Splash.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3000
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem Code to Display Splash screen
Rem Written by M. Russell April-May 2000
Option Explicit

Rem Exit button selected
Private Sub Command1_Click()
Unload Me
End Sub

Rem Update textbox with program info
Private Sub Form_Load()
With Text1
    .Text = new_line & "    Written by M. Russell" & new_line
    .Text = .Text & "   as part of a HND course" & new_line
    .Text = .Text & "   at Bradford College" & new_line & new_line
    .Text = .Text & "   Version 1.0" & new_line
    .Text = .Text & "   May 2000"
End With
End Sub
