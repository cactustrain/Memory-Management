VERSION 5.00
Begin VB.Form menu 
   Caption         =   "Operating System Memory Management Simualtion"
   ClientHeight    =   6600
   ClientLeft      =   120
   ClientTop       =   690
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   6885
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Help"
      Height          =   615
      Left            =   2760
      TabIndex        =   4
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Amend Parameters"
      Height          =   735
      Left            =   2760
      TabIndex        =   3
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Quit"
      Height          =   615
      Left            =   2760
      TabIndex        =   2
      Top             =   5400
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start Simulation"
      Height          =   735
      Left            =   2760
      TabIndex        =   1
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update Process List"
      Height          =   735
      Left            =   2760
      TabIndex        =   0
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   5
      Top             =   360
      Width           =   4815
   End
End
Attribute VB_Name = "menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem Code for main menu
Rem Written by M. Russell April-May 2000
Option Explicit

Rem Process setup selected
Private Sub Command1_Click()
Load process_setup
process_setup.Show 1
End Sub

Rem Start Simulation selected
Private Sub Command2_Click()
If process = 0 Then
    MsgBox ("You have not set up any processes")
    Exit Sub
End If
simulation.Show
End Sub

Rem Exit program selected
Private Sub Command3_Click()
End
End Sub

Rem Amend Parameters screen selected
Private Sub Command4_Click()
Load amend_params
amend_params.Show 1
End Sub

Rem Help screen selected
Private Sub Command5_Click()
Help_Form.Show
End Sub
