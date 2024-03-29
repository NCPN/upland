Option Compare Database
Option Explicit

' ------------------------------------------
' -------         REVISIONS           ------
' ------------------------------------------
'  BLC - 8/10/2015 - Vestigial, unused subroutines FillCanopyGapsOld & UpdateCanopyGapsOld
'                    were removed. These routines, developed by R. DenBleyker in 11/2006
'                    caused the application to fail to compile.
'                    FillCanopyGaps() & UpdateCanopyGaps() handle their functionality.
'                    Option Explicit was added.
' ------------------------------------------

Sub ClearCanopyGaps(intLastField As Integer)
' Clear fields on fsub_Canopy_Gap (when navigating transects).  11/2006 Russ DenBleyker
' Northern Colorado Plateau Network
    On Error GoTo Err_Handler

    Dim intRecordCount As Integer
    Dim strFieldName As String

    If intLastField > 0 Then
      intRecordCount = 1
      Do Until intRecordCount > intLastField
        ' Clear form fields.
        strFieldName = "F" & intRecordCount  ' Set field name
        Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!(strFieldName) = Null
        strFieldName = "F" & (intRecordCount + 1) ' Set field name
        Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!(strFieldName) = Null
        intRecordCount = intRecordCount + 2
      Loop
    End If
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (ClearCanopyGaps)"
            Resume Exit_Sub
    End Select

End Sub

Sub ClearBasalGaps(intLastField As Integer)
' Clear fields on fsub_Basal_Gap (when navigating transects).  11/2006 Russ DenBleyker
' Northern Colorado Plateau Network
    On Error GoTo Err_Handler

    Dim intRecordCount As Integer
    Dim strFieldName As String

    If intLastField > 0 Then
      intRecordCount = 1
      Do Until intRecordCount > intLastField
        ' Clear form fields.
        strFieldName = "F" & intRecordCount  ' Set field name
        Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!(strFieldName) = Null
        strFieldName = "F" & (intRecordCount + 1) ' Set field name
        Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!(strFieldName) = Null
        intRecordCount = intRecordCount + 2
      Loop
    End If
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (ClearBasalGaps)"
            Resume Exit_Sub
    End Select

End Sub
Function FillBasalGaps(strTransectID As String) As Integer
' Fill controls in fsub_Basal_Gap from tbl_Basal_Gaps.  11/2006 Russ DenBleyker
' Northern Colorado Plateau Network
    On Error GoTo Err_Handler
    
    Dim stDocName As String
    Dim db As Database
    Dim Gaps As dao.Recordset
    Dim strSQL As String
    Dim intRecordCount As Integer
    Dim strFieldName As String

    strSQL = "Select * FROM tbl_Basal_Gaps WHERE Transect_ID = '" & strTransectID & "' ORDER BY Gap_ID"
    Set db = CurrentDb
  ' Get the records for this transect
    Set Gaps = db.OpenRecordset(strSQL)
    If Gaps.EOF Then
      Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!LastEnd = 0
      Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!LastStart = 0
      FillBasalGaps = 0  ' Indicate no records retreived
      Gaps.Close
      GoTo Exit_Procedure
    End If
    intRecordCount = 0
    Gaps.MoveFirst
    Do Until Gaps.EOF
    ' Fill form fields.
      intRecordCount = intRecordCount + 1
      strFieldName = "F" & intRecordCount  ' Set field name
      Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!(strFieldName) = Gaps!Gap_Start
      intRecordCount = intRecordCount + 1
      strFieldName = "F" & intRecordCount  ' Set field name
      Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!(strFieldName) = Gaps!Gap_End
      Gaps.MoveNext
    Loop
    ' Save last entries for future verification purposes
    Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!LastEnd = Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!(strFieldName)
    strFieldName = "F" & (intRecordCount - 1)  ' Set field name
    Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!LastStart = Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!(strFieldName)
    
    FillBasalGaps = intRecordCount  ' Indicate index of last record encountered
    Gaps.Close
    
Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (FillBasalGaps)"
            Resume Exit_Procedure
    End Select

End Function

' ---------------------------------
' FUNCTION:     UpdateBasalGaps
' Description:  Update tbl_Basal_Gaps from controls in fsub_Basal_Gap
' Assumptions:  -
' Parameters:   strTransectID - transect identifier (string)
'               strField - field to update (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Russ DenBleyker, November 2006 - Northern Colorado Plateau Network
' Adapted:      -
' Revisions:
'   RD  - 11/2006   - initial version
'   BLC - 8/11/2015 - defined intField1, intField2 (integer), updated error handler/documentation
' ---------------------------------
Function UpdateBasalGaps(strTransectID As String, strField As String) As Integer
    On Error GoTo Err_Handler
    
    Dim stDocName As String
    Dim db As Database
    Dim Gaps As dao.Recordset
    Dim NewGap As dao.Recordset
    Dim strSQL As String
    Dim strFieldName1 As String
    Dim strFieldName2 As String
    Dim intField As Integer  ' The integer portion of the field name will serve as Gap_ID.
    Dim intField1 As Integer
    Dim intField2 As Integer
    Dim result As Byte
    
    If IsNull(strField) Then
      UpdateBasalGaps = 0  ' Field name is null, disregard.
      GoTo Exit_Function
    End If
    If IsNull(Forms!frm_Data_Entry!frm_Canopy_Transect.Form!Visit_Date) Then
      MsgBox "Visit date required."
      UpdateBasalGaps = 0
      GoTo Exit_Function
    End If
    intField = Right(strField, Len(strField) - 1) ' Get the field number
    result = intField Mod 2
    If result = 0 Then   ' Set up the field numbers properly.
      intField2 = intField
      intField1 = intField - 1
    Else
      intField1 = intField
      intField2 = intField + 1
    End If
    strFieldName1 = "F" & intField1  ' Set first field name
    strFieldName2 = "F" & intField2  ' Set second field name
    If intField > Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!LastField Then  ' If field number higher than last one filled this is an add.
      If intField > Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!LastField + 2 Then  ' If field number higher than last one plus two, this is a screw-up.
        MsgBox "You cannot have gaps in the gap form!"
        UpdateBasalGaps = 1
        GoTo Exit_Function
      End If
      If result = 1 And IsNull(Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!(strFieldName2)) Then
        UpdateBasalGaps = 0  ' If add and 1st field of pair, then there is no error if second is null.
        GoTo Exit_Function   ' Give them a chance to enter the second one
      End If
      If result = 0 And IsNull(Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!(strFieldName1)) Then
        MsgBox "Start required."
        UpdateBasalGaps = 1  ' If add and 2nd field of pair, abort for missing start.
        GoTo Exit_Function
      End If
      If CInt(Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!(strFieldName1)) > 5000 Then
        MsgBox "Start value exceeds domain limit."
        UpdateBasalGaps = 1  ' Indicate no update.
        GoTo Exit_Function
      End If
      If CInt(Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!(strFieldName2)) > 5000 Then
        MsgBox "End value exceeds domain limit."
        UpdateBasalGaps = 1  ' Indicate no update.
        GoTo Exit_Function
      End If
      If (CInt(Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!(strFieldName1)) >= CInt(Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!LastEnd)) And (CInt(Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!LastEnd) > 0) Then
        MsgBox "Start must be less than previous end."
        UpdateBasalGaps = 1  ' Indicate no update.
        GoTo Exit_Function
      End If
      If CInt(Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!(strFieldName1)) <= CInt(Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!(strFieldName2)) Then
        MsgBox "Start must be greater than end."
        UpdateBasalGaps = 1  ' Indicate no update.
        GoTo Exit_Function
      End If
      If Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!(strFieldName1) - Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!(strFieldName2) < 20 Then
        MsgBox "You cannot have gaps < 20cm."
        UpdateBasalGaps = 1  ' Indicate no update.
        GoTo Exit_Function
      End If
      Set db = CurrentDb
      Set NewGap = db.OpenRecordset("tbl_Basal_Gaps")
        NewGap.AddNew
        NewGap!Gap_ID = intField1
        NewGap!Transect_ID = strTransectID
        NewGap!Gap_Start = Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!(strFieldName1)
        NewGap!Gap_End = Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!(strFieldName2)
        NewGap.Update
        NewGap.Close
        ' Save last field count and values for navigation and verification purposes.
        Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!LastStart = Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!(strFieldName1)
        Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!LastEnd = Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!(strFieldName2)
        Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!LastField = intField2  ' Update last filled field on form
        UpdateBasalGaps = 0  ' Indicate OK.
    Else   ' Update routines.
      strSQL = "Select * FROM tbl_Basal_Gaps WHERE Transect_ID = '" & strTransectID & "' AND Gap_ID = " & intField1
      Set db = CurrentDb
      Set Gaps = db.OpenRecordset(strSQL)
      Gaps.MoveFirst
      Gaps.Edit
      If result = 0 Then
        If Gaps!Gap_Start - Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!(strFieldName2) < 20 Then
          MsgBox "You cannot have gaps < 20cm."
          UpdateBasalGaps = 1  ' Indicate no update.
          Gaps.Close
          GoTo Exit_Function
        End If
        Gaps!Gap_End = Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!(strFieldName2)
      Else
        If Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!(strFieldName1) - Gaps!Gap_End < 20 Then
          MsgBox "You cannot have gaps < 20cm."
          UpdateBasalGaps = 1  ' Indicate no update.
          Gaps.Close
          GoTo Exit_Function
        End If
        Gaps!Gap_Start = Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Basal_Gap.Form!(strFieldName1)
      End If
      Gaps.Update
      Gaps.Close
      UpdateBasalGaps = 0  ' Indicate OK.
    End If   '  End if for add/edit decision.

Exit_Function:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - UpdateBasalGaps[mod_GapHandlers])"
    End Select
    Resume Exit_Function
End Function

Function FillCanopyGaps(strTransectID As String) As Integer
' Fill controls in fsub_Canopy_Gap from tbl_Canopy_Gaps.  3/2011 Russ DenBleyker
' Northern Colorado Plateau Network
    On Error GoTo Err_Handler
    
    Dim stDocName As String
    Dim db As Database
    Dim Gaps As dao.Recordset
    Dim strSQL As String
    Dim intRecordCount As Integer
    Dim strFieldName As String

    strSQL = "Select * FROM tbl_Canopy_Gaps WHERE Transect_ID = '" & strTransectID & "' ORDER BY Gap_ID"
    Set db = CurrentDb
  ' Get the records for this transect
    Set Gaps = db.OpenRecordset(strSQL)
    If Gaps.EOF Then
      Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!LastEnd = 0
      Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!LastStart = 0
      FillCanopyGaps = 0  ' Indicate no records retreived
      Gaps.Close
      GoTo Exit_Procedure
    End If
    intRecordCount = 0
    Gaps.MoveFirst
    Do Until Gaps.EOF
    ' Fill form fields.
      intRecordCount = intRecordCount + 1
      strFieldName = "F" & intRecordCount  ' Set field name
      Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!(strFieldName) = Gaps!Start
      intRecordCount = intRecordCount + 1
      strFieldName = "F" & intRecordCount  ' Set field name
      Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!(strFieldName) = Gaps!Gap_End
      Gaps.MoveNext
    Loop
    ' Save last entries for future verification purposes
    Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!LastEnd = Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!(strFieldName)
    strFieldName = "F" & (intRecordCount - 1)  ' Set field name
    Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!LastStart = Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!(strFieldName)
    
    FillCanopyGaps = intRecordCount  ' Indicate index of last record encountered
    Gaps.Close
    
Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (FillCanopyGaps)"
            Resume Exit_Procedure
    End Select

End Function

' ---------------------------------
' FUNCTION:     UpdateCanopyGaps
' Description:  Update tbl_Canopy_Gaps from controls in fsub_Canopy_Gap
' Assumptions:  -
' Parameters:   strTransectID - transect identifier (string)
'               strField - field to update (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Russ DenBleyker, March 2011 - Northern Colorado Plateau Network
' Adapted:      -
' Revisions:
'   RD  -  3/2011   - initial version
'   BLC - 8/11/2015 - defined intField1, intField2 (integer), updated error handler/documentation
' ---------------------------------
Function UpdateCanopyGaps(strTransectID As String, strField As String) As Integer
    On Error GoTo Err_Handler
    
    Dim stDocName As String
    Dim db As Database
    Dim Gaps As dao.Recordset
    Dim NewGap As dao.Recordset
    Dim strSQL As String
    Dim strFieldName1 As String
    Dim strFieldName2 As String
    Dim intField As Integer  ' The integer portion of the field name will serve as Gap_ID.
    Dim intField1 As Integer
    Dim intField2 As Integer
    Dim result As Byte
    
    If IsNull(strField) Then
      UpdateCanopyGaps = 0  ' Field name is null, disregard.
      GoTo Exit_Function
    End If
    If IsNull(Forms!frm_Data_Entry!frm_Canopy_Transect.Form!Visit_Date) Then
      MsgBox "Visit date required."
      UpdateCanopyGaps = 0
      GoTo Exit_Function
    End If
    intField = Right(strField, Len(strField) - 1) ' Get the field number
    result = intField Mod 2
    If result = 0 Then   ' Set up the field numbers properly.
      intField2 = intField
      intField1 = intField - 1
    Else
      intField1 = intField
      intField2 = intField + 1
    End If
    strFieldName1 = "F" & intField1  ' Set first field name
    strFieldName2 = "F" & intField2  ' Set second field name
    If intField > Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!LastField Then  ' If field number higher than last one filled this is an add.
      If intField > Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!LastField + 2 Then  ' If field number higher than last one plus two, this is a screw-up.
        MsgBox "You cannot have gaps in the gap form!"
        UpdateCanopyGaps = 1
        GoTo Exit_Function
      End If
      If result = 1 And IsNull(Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!(strFieldName2)) Then
        UpdateCanopyGaps = 0  ' If add and 1st field of pair, then there is no error if second is null.
        GoTo Exit_Function   ' Give them a chance to enter the second one
      End If
      If result = 0 And IsNull(Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!(strFieldName1)) Then
        MsgBox "Start required."
        UpdateCanopyGaps = 1  ' If add and 2nd field of pair, abort for missing start.
        GoTo Exit_Function
      End If
      If CInt(Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!(strFieldName1)) > 5000 Then
        MsgBox "Start value exceeds domain limit."
        UpdateCanopyGaps = 1  ' Indicate no update.
        GoTo Exit_Function
      End If
      If CInt(Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!(strFieldName2)) > 5000 Then
        MsgBox "End value exceeds domain limit."
        UpdateCanopyGaps = 1  ' Indicate no update.
        GoTo Exit_Function
      End If
      If (CInt(Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!(strFieldName1)) <= CInt(Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!LastEnd)) And (CInt(Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!LastEnd) > 0) Then
        MsgBox "Start must be greater than previous end."
        UpdateCanopyGaps = 1  ' Indicate no update.
        GoTo Exit_Function
      End If
      If CInt(Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!(strFieldName1)) >= CInt(Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!(strFieldName2)) Then
        MsgBox "Start must be less than end."
        UpdateCanopyGaps = 1  ' Indicate no update.
        GoTo Exit_Function
      End If
      If Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!(strFieldName2) - Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!(strFieldName1) < 20 Then
        MsgBox "You cannot have gaps < 20cm."
        UpdateCanopyGaps = 1  ' Indicate no update.
        GoTo Exit_Function
      End If
      Set db = CurrentDb
      Set NewGap = db.OpenRecordset("tbl_Canopy_Gaps")
        NewGap.AddNew
        NewGap!Gap_ID = intField1
        NewGap!Transect_ID = strTransectID
        NewGap!Start = Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!(strFieldName1)
        NewGap!Gap_End = Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!(strFieldName2)
        NewGap.Update
        NewGap.Close
        ' Save last field count and values for navigation and verification purposes.
        Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!LastStart = Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!(strFieldName1)
        Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!LastEnd = Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!(strFieldName2)
        Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!LastField = intField2  ' Update last filled field on form
        UpdateCanopyGaps = 0  ' Indicate OK.
    Else   ' Update routines.
      strSQL = "Select * FROM tbl_Canopy_Gaps WHERE Transect_ID = '" & strTransectID & "' AND Gap_ID = " & intField1
      Set db = CurrentDb
      Set Gaps = db.OpenRecordset(strSQL)
      Gaps.MoveFirst
      Gaps.Edit
      If result = 0 Then
        If Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!(strFieldName2) - Gaps!Start < 20 Then
          MsgBox "You cannot have gaps < 20cm."
          UpdateCanopyGaps = 1  ' Indicate no update.
          Gaps.Close
          GoTo Exit_Function
        End If
        Gaps!Gap_End = Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!(strFieldName2)
      Else
        If Gaps!Gap_End - Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!(strFieldName1) < 20 Then
          MsgBox "You cannot have gaps < 20cm."
          UpdateCanopyGaps = 1  ' Indicate no update.
          Gaps.Close
          GoTo Exit_Function
        End If
        Gaps!Start = Forms!frm_Data_Entry!frm_Canopy_Transect.Form!fsub_Canopy_Gap.Form!(strFieldName1)
      End If
      Gaps.Update
      Gaps.Close
      UpdateCanopyGaps = 0  ' Indicate OK.
    End If   '  End if for add/edit decision.

Exit_Function:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - UpdateBasalGaps[mod_GapHandlers])"
    End Select
    Resume Exit_Function
End Function