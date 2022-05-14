Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Function GetReports()      'new function (az)
    On Error GoTo Err_SelectPlots
    Dim st As String
    Dim EventType As Integer
    Dim r As Integer
    
'    st = Screen.ActiveControl.Name
'     If st = "btnLoadInfestEvents" Then
'        EventType = vbYes               ' Load infestation events
'     ElseIf st = "btnLoadTransEvents" Then
        EventType = vbNo                ' Load transect events
 '    Else
 '       MsgBox "Error"
 '   End If
    Debug.Print EventType
'    Call LoadEvents(EventType)         'call production sub
    r = Int(100 * Rnd)
    If r = 33 Then Call zDoesItWork(EventType)      'call test sub
    
Exit_SelectPlots:
    Exit Function

Err_SelectPlots:
    MsgBox Err.Description
    Resume Exit_SelectPlots
    
End Function

Public Sub zDoesItWork(response As Integer) 'test sub; delete (az)
    Debug.Print "response = " & response
    If response = vbYes Then
        MsgBox "A chemist and his friend walked into a bar." & vbNewLine & "The chemist ordered H-2-O"
        MsgBox "His friend ordered H-2-O-2"
    Else: MsgBox "If you're not part of the solution, you're part of the precipitate."
    End If
End Sub