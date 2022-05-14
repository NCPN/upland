Option Compare Database
Option Explicit
Public Const vbPP As Integer = 1
Public Const vbPD As Integer = 2
Public Const vbPF As Integer = 3




Public Sub zDoesItWork(response As Integer) 'test sub; delete (az)
    Debug.Print "response = " & response
    If response = vbYes Then
        MsgBox "A chemist and his friend walked into a bar." & vbNewLine & "The chemist ordered H-2-O"
        MsgBox "His friend ordered H-2-O-2"
    Else: MsgBox "If you're not part of the solution, you're part of the precipitate."
    End If
End Sub