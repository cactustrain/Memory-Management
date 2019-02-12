VERSION 5.00
Begin VB.Form report_form 
   Caption         =   "Memory Management Simulation"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Exit to Main Menu"
      Height          =   495
      Left            =   6240
      TabIndex        =   3
      Top             =   7800
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   7800
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   720
      Width           =   11535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Simulation Results Log"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "report_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem Code to display report form at end of simulation run
Rem Written by M. Russell April-May 2000
Option Explicit

Rem Print report list selected
Private Sub Command1_Click()
Printer.Print Text1.Text
Printer.EndDoc
End Sub

Rem Exit report screen selected
Private Sub Command2_Click()
Text1.Enabled = True    ' Re-enable textbox so can be updated on new run
Unload Me
End Sub
