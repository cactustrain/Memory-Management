Attribute VB_Name = "OSsim"
Rem Global variable declarations
Rem Written by M. Russell April-May 2000
Option Explicit
Public frag_max As Integer
Public new_line As String
Public n, i, j As Integer
Public clock, frag_clock As Integer
Public compaction, comp_count As Integer
Public fixed As Boolean
Public policy As Integer
Public halt As Boolean
Public memory(3 To 20) As Integer
Public part(6) As Integer
Public process As Integer
Public size(1 To 26) As Integer
Public arrival(1 To 26) As Integer
Public run(1 To 26) As Integer

Rem This is the first code executed when the program is loaded
Sub main()
new_line = Chr$(13) & Chr$(10)  ' Setup constant
Splash.Show 1   ' Display splash screen as modal form
menu.Show
End Sub


