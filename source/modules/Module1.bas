Option Compare Database
Option Explicit

Public Function Get_Old_Forms()      'new function (az)
    On Error GoTo Err_Handler
    
    Dim st As String
    Dim ButtonType As Integer
    Dim r As Integer
    Dim FormToOpen As String
    
    st = Screen.ActiveControl.Name
     If st = "btnOT" Then
        FormToOpen = "frm_Select_Overstory_Revisit"
     ElseIf st = "btnPR" Then
        FormToOpen = "frm_Select_Plot_Establishment"
     ElseIf st = "btnSP" Then
        FormToOpen = "frm_Species_Report_Select"
     Else
         MsgBox "Error"
    End If
        DoCmd.OpenForm FormToOpen       'open form
        
    r = Int(100 * Rnd)
    If r = 33 Then Call zDoesItWork(vbYes)      'call test sub
    
Exit_Handler:
    Exit Function

Err_Handler:
    MsgBox Err.Description
    Resume Exit_Handler
    
End Function

Public Sub zzDoesItWork(response As Integer) 'test sub; delete (az)
    Debug.Print "response = " & response
    If response = vbYes Then
        MsgBox "A chemist and his friend walked into a bar." & vbNewLine & "The chemist ordered H-2-O"
        MsgBox "His friend ordered H-2-O-2"
    Else: MsgBox "If you're not part of the solution, you're part of the precipitate."
    End If
End Sub