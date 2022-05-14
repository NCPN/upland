Option Compare Database
Option Explicit

' SUB:          MkMydir
' Description:  Creates directories and sub-directories
' Assumptions:
' Parameters:   spath (the desired directory)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  AZ, March 2022
' Adapted:      Daniel Pineault -
'               https://www.devhut.net/vba-create-directory-structurecreate-multiple-directories/
' Revisions:
'   AZ  - 3/25/2022  - initial version
' ---------------------------------

Public Sub MkMyDir(spath As String)

Dim iStart As Integer
Dim aDirs As Variant
Dim sCurDir As String
Dim i As Integer
    Debug.Print "called"
    If spath <> "" Then
        aDirs = Split(spath, "\")
        If Left(spath, 2) = "\\" Then
            iStart = 3
        Else
            iStart = 1
        End If
 
        sCurDir = Left(spath, InStr(iStart, spath, "\"))
 
        For i = iStart To UBound(aDirs)
            sCurDir = sCurDir & aDirs(i) & "\"
            If Dir(sCurDir, vbDirectory) = vbNullString Then
                MkDir sCurDir
            End If
        Next i
    End If
End Sub