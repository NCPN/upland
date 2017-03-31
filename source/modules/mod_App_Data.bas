Option Compare Database

Option Explicit

' =================================
' MODULE:       mod_App_Data
' Level:        Application module
' Version:      1.26
' Description:  data functions & procedures specific to this application
'
' Source/date:  Bonnie Campbell, 2/9/2015
' Revisions:    BLC - 2/9/2015  - 1.00 - initial version
'               BLC - 2/18/2015 - 1.01 - included subforms in fillList
'               BLC - 5/1/2015  - 1.02 - integerated into Invasives Reporting tool
'               BLC - 5/22/2015 - 1.03 - added PopulateList()
'               BLC - 6/3/2015  - 1.04 - added IsUsedTargetArea()
'               BLC - 5/5/2016  - 1.05 - added GetRiverSegments(), GetProtocolVersion()
'                                        changed to Exit_Handler vs. Exit_Function
'               BLC - 6/28/2016 - 1.06 - added ToggleIsActive(), revised getParkState() to GetParkState()
'               BLC - 7/26/2016 - 1.07 - added SetRecord(), GetRecords()
'               BLC - 7/28/2016 - 1.08 - added UpsertRecord()
'               BLC - 7/30/2016 - 1.09 - added ToggleSensitive()
'               BLC - 8/8/2016  - 1.10 - updated UpsertRecord() for additional forms
'               BLC - 9/1/2016  - 1.11 - added UploadSurveyFile(), updated UpsertRecord()
'               BLC - 9/13/2016 - 1.12 - added FetchAddlData()
'               BLC - 9/21/2016 - 1.13 - updated SetRecord() i_login parameters
'               BLC - 9/22/2016 - 1.14 - added templates
'               BLC - 10/16/2016 - 1.15 - fixed PopulateCombobox() to properly set recordset
'               BLC - 10/19/2016 - 1.16 - renamed UploadSurveyFile() to UploadCSVFile() to genericize
'               BLC - 10/24/2016 - 1.17 - updated SetRecord(), ToggleIsActive()
'               BLC - 10/28/2016 - 1.18 - updated i_task, TempVars("ContactID") -> TempVars("AppUserID")
'               BLC - 1/9/2017   - 1.19 - revised UpsertRecord from ContactID to ID,
'                                         added GetRecords templates
'               BLC - 1/24/2017  - 1.20 - added IsNPS flag for SetRecord() contacts
'               BLC - 2/1/2017   - 1.21 - updated UpsertRecord() to handle form upserts
'                                         for forms w/o lists/msg & msg icons
'               BLC - 2/3/2017   - 1.22 - location adjustments for UpsertRecord() & SetRecord()
'               BLC - 2/7/2017   - 1.23 - added template - s_location_with_loctypeID_sensitivity
' --------------------------------------------------------------------
'               BLC, 3/22/2017          added updated version to Upland db
' --------------------------------------------------------------------
'               BLC, 3/22/2017  - 1.24 - removed big rivers only components
'                                        revised for uplands
'               BLC, 3/29/2017  - 1.25 - added FieldCheck, FieldOK, Dependencies for templates
'               BLC, 3/30/2017  - 1.26 - added non-parameterized query option for GetRecords()
' =================================

'' ---------------------------------
'' SUB:          fillList
'' Description:  Fill a list (or listbox like subform) from specific queries for datasheets, species or other items
'' Assumptions:  Either a listbox or subform control is being populated
'' Parameters:   frm - main form object
''               ctrl - either:
''                      lbx - main form listbox object (for filling a listbox control)
''                      sfrm - subform object (for populating a subform control)
'' Returns:      N/A
'' Throws:       none
'' References:   none
'' Source/date:
'' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
'' Revisions:
''   BLC, 2/6/2015  - initial version
''   BLC, 2/18/2015 - adapted to include subform as well as listbox controls
''   BLC, 5/1/2015  - integrated into Invasives Reporting tool
'' ---------------------------------
'Public Sub fillList(frm As Form, ctrlSource As Control, Optional ctrlDest As Control)
'
'On Error GoTo Err_Handler
'
'    Dim db As DAO.Database
'    Dim rs As DAO.Recordset
'    Dim strQuery As String, strSQL As String
'
'    'output to form or listbox control?
'
'    'determine data source
'    Select Case ctrlSource.name
'
'        Case "lbxDataSheets", "sfrmDatasheets" 'Datasheets
'            strQuery = "qry_Active_Datasheets"
'            strSQL = CurrentDb.QueryDefs(strQuery).sql
'
'        Case "lbxSpecies", "lbxTgtSpecies", "fsub_Species_Listbox" 'Species
'            strQuery = "qry_Plant_Species"
'            strSQL = CurrentDb.QueryDefs(strQuery).sql
'
'    End Select
'
'    'fetch data
'    Set db = CurrentDb
'    Set rs = db.OpenRecordset(strSQL)
'
'    'set TempVars
'    TempVars.Add "strSQL", strSQL
'
'    If Not ctrlDest Is Nothing Then
'        'populate list & headers
'        PopulateList ctrlSource, rs, ctrlDest
'    Else
'        'populate only ctrlSource headers
'        PopulateListHeaders ctrlSource, rs
'    End If
'
'Exit_Handler:
'    Exit Sub
'
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - fillList[mod_App_Data])"
'    End Select
'    Resume Exit_Handler
'End Sub

'' ---------------------------------
'' SUB:          PopulateList
'' Description:  Populate listbox and similar controls from recordset
'' Assumptions:  -
'' Parameters:   ctrlSource - source control (listbox/listview)
''               rs - recordset used to populate control (recordset object)
''               ctrlDest - destination control (listbox/listview)
'' Returns:      -
'' Throws:       none
'' References:   none
'' Source/date:
'' krish KM, Aug. 27, 2014
'' http://stackoverflow.com/questions/25526904/populate-listbox-using-ado-recordset
'' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
'' Revisions:
''   BLC - 2/6/2015 - initial version
''   BLC - 5/10/2015 - moved to mod_List from mod_Lists
''   BLC - 5/20/2015 - changed from tbxMasterCode to tbxLUCode
''   BLC - 5/22/2015 - moved to mod_App_Data from mod_List
'' ---------------------------------
'Public Sub PopulateList(ctrlSource As Control, rs As Recordset, ctrlDest As Control)
'
'On Error GoTo Err_Handler
'
'    Dim frm As Form
'    Dim rows As Integer, cols As Integer, i As Integer, j As Integer, matches As Integer, iZeroes As Integer
'    Dim strItem As String, strColHeads As String, aryColWidths() As String
'
'    Set frm = ctrlSource.Parent
'
'    rows = rs.RecordCount
'    cols = rs.Fields.Count
'
'    'address no records
'    If Nz(rows, 0) = 0 Then
'        MsgBox "Sorry, no records found..."
'        GoTo Exit_Handler
'    End If
'
'    'handle sfrm controls (acSubform = 112)
'    If ctrlSource.ControlType = acSubform Then
'        Set ctrlSource.Form.Recordset = rs
'
'        ctrlSource.Form.Controls("tbxCode").ControlSource = "Code"
'        ctrlSource.Form.Controls("tbxSpecies").ControlSource = "Species"
'        'ctrlSource.Form.Controls("tbxMasterCode").ControlSource = "Master_PLANT_Code"
'        ctrlSource.Form.Controls("tbxLUCode").ControlSource = "LUCode"
'        ctrlSource.Form.Controls("tbxTransectOnly").ControlSource = "Transect_Only"
'        ctrlSource.Form.Controls("tbxTgtAreaID").ControlSource = "Target_Area_ID"
'
'        'set the initial record count (MoveLast to get full count, MoveFirst to set display to first)
'        rs.MoveLast
'        ctrlSource.Parent.Form.Controls("lblSfrmSpeciesCount").Caption = rs.RecordCount & " species"
'        rs.MoveFirst
'
'        GoTo Exit_Handler
'    End If
'
'    'fetch column widths array
'    aryColWidths = Split(ctrlSource.ColumnWidths, ";")
'
'    'count number of 0 width elements
'    iZeroes = CountArrayValues(aryColWidths, "0")
'
'    'clear out existing values
'    ClearList ctrlSource
'
'    'populate column names (if desired)
'    If ctrlSource.ColumnHeads = True Then
'        PopulateListHeaders ctrlSource, rs
'
'        'populate second listbox headers if present
'        If ctrlDest.ColumnHeads = True Then
'            ClearList ctrlDest
'            PopulateListHeaders ctrlDest, rs
'        End If
'    End If
'
'    'populate data
'    Select Case ctrlSource.RowSourceType
'        Case "Table/Query"
'            Set ctrlSource.Recordset = rs
'        Case ""
'
'            'initialize
'            i = 0
'
'            Do Until rs.EOF
'
'                'initialize item
'                strItem = ""
'
'                'generate item
'                For j = 0 To cols - 1
'                    'check if column is displayed width > 0
'                    If CInt(aryColWidths(j)) > 0 Then
'
'                        strItem = strItem & rs.Fields(j).Value & ";"
'
'                        'determine how many separators there are (";") --> should equal # cols
'                        matches = (Len(strItem) - Len(Replace$(strItem, ";", ""))) / Len(";")
'
'                        'add item if not already in list --> # of ; should equal cols - 1
'                        'but # in list should only be # of non-zero columns --> cols - iZeroes
'                        If matches = cols - iZeroes Then
'                            ctrlSource.AddItem strItem
'                            'reset the string
'                            strItem = ""
'                        End If
'
'                    End If
'
'                Next
'
'                i = i + 1
'
'                rs.MoveNext
'            Loop
'        Case "Field List"
'    End Select
'
'Exit_Handler:
'    Exit Sub
'
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - PopulateList[mod_App_Data])"
'    End Select
'    Resume Exit_Handler
'End Sub

'' ---------------------------------
'' SUB:          AddListToTable
'' Description:  Populate table from listbox
'' Assumptions:  -
'' Parameters:   lbx - listbox control
'' Returns:      -
'' Throws:       none
'' References:   none
'' Source/date:  Bonnie Campbell, June 3, 2015 - for NCPN tools
'' Adapted:      -
'' Revisions:
''   BLC - 6/3/2015 - initial version
'' ---------------------------------
'Public Sub AddListToTable(lbx As ListBox)
'
'On Error GoTo Err_Handler
'
'Dim aryFields() As String
'Dim aryFieldTypes() As Variant
'Dim strCode As String, strSpecies As String, strLUCode As String
'Dim iRow As Integer, iTransectOnly As Integer, iTgtAreaID As Integer
'
'    iRow = lbx.ListCount - 1 'Forms("frm_Tgt_Species").Controls("lbxTgtSpecies").ListCount - 1
'
'    ReDim Preserve aryFields(0 To iRow)
'
'    'header row (iRow = 0)
'    aryFields(0) = "Code;Species;LUCode;Transect_Only;Target_Area_ID"   'iRow = 0
'    aryFieldTypes = Array(dbText, dbText, dbText, dbInteger, dbInteger)
'
'    'data rows (iRow > 0)
'    For iRow = 1 To lbx.ListCount - 1
'
'        ' ---------------------------------------------------
'        '  NOTE: listbox column MUST have a non-zero width to retrieve its value
'        ' ---------------------------------------------------
'         strCode = lbx.Column(0, iRow) 'column 0 = Master_PLANT_Code (Code)
'         strSpecies = lbx.Column(1, iRow) 'column 1 = Species name (Species)
'         strLUCode = lbx.Column(2, iRow) 'column 2 = LU_Code (LUCode)
'         iTransectOnly = Nz(lbx.Column(3, iRow), 0) 'column 3 = Transect_Only (TransectOnly)
'         iTgtAreaID = Nz(lbx.Column(4, iRow), 0) 'column 4 = Target_Area_ID (TgtAreaID)
'
'        aryFields(iRow) = strCode & ";" & strSpecies & ";" & strLUCode & ";" & iTransectOnly & ";" & iTgtAreaID
'
'    Next
'
'    'save the existing records to temp_Listbox_Recordset & replace any existing records
'    SetListRecordset lbx, True, aryFields, aryFieldTypes, "temp_Listbox_Recordset", True
'
'Exit_Handler:
'    Exit Sub
'
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - PopulateList[mod_App_Data])"
'    End Select
'    Resume Exit_Handler
'End Sub

' ---------------------------------
' FUNCTION:     GetParkState
' Description:  Retrieve the state associated with a park (via tlu_Parks)
' Assumptions:  Park state is properly identified in tlu_Parks
' Parameters:   parkCode - 4 character park designator
' Returns:      ParkState - 2 character state abbreviation
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 19, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/19/2015  - initial version
'   BLC - 6/28/2016  - revised to uppercase GetParkState vs getParkState
' ---------------------------------
Public Function GetParkState(ParkCode As String) As String

On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim State As String, strSQL As String
   
    'handle only appropriate park codes
    If Len(ParkCode) <> 4 Then
        GoTo Exit_Handler
    End If
    
    'generate SQL ==> NOTE: LIMIT 1; syntax not viable for Access, use SELECT TOP x instead
    strSQL = "SELECT TOP 1 ParkState FROM tlu_Parks WHERE ParkCode LIKE '" & ParkCode & "';"
            
    'fetch data
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL)
    
    'assume only 1 record returned
    If rs.RecordCount > 0 Then
        State = rs.Fields("ParkState").value
    End If
   
    'return value
    GetParkState = State
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetParkState[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     getListLastModifiedDate
' Description:  Retrieve the last modified date with a park (via tbl_Target_List)
' Assumptions:  -
' Parameters:   tgtYear - 4 digit year of list (integer)
'               parkCode - 4 character park designator (string)
' Returns:      date - last modified date (mmm-d-yyyy H:nn AMPM format) for the specified target list (string)
'                      if NULL (no last modified date) returns empty string
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, June 10, 2015 - for NCPN tools
' Revisions:
'   BLC - 6/10/2015  - initial version
' ---------------------------------
Public Function getListLastModifiedDate(TgtYear As Integer, ParkCode As String) As String

On Error GoTo Err_Handler
    
    Dim strCriteria As String

    'handle only appropriate park codes
    If Len(ParkCode) <> 4 Or TgtYear < 2000 Then
        GoTo Exit_Handler
    End If
    
    'set lookup criteria
    strCriteria = "Park_Code LIKE '" & ParkCode & "' AND CInt(Target_Year) = " & CInt(TgtYear)
    
    'Debug.Print strCriteria
        
    'lookup last modified date & return value
    getListLastModifiedDate = Nz(Format(DLookup("Last_Modified", "tbl_Target_List", strCriteria), "mmm-d-yyyy H:nn AMPM"), "")
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - getListLastModifiedDate[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     IsUsedTargetArea
' Description:  Determine if the target area is in use by a target list
' Parameters:   TgtAreaID - target area idenifier (integer)
' Returns:      boolean - true if target area is in use, false if not
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, June 3, 2015 - for NCPN tools
' Revisions:
'   BLC - 6/3/2015  - initial version
' ---------------------------------
Public Function IsUsedTargetArea(TgtAreaID As Integer) As Boolean

On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    
    'default
    IsUsedTargetArea = False
    
    'generate SQL ==> NOTE: LIMIT 1; syntax not viable for Access, use SELECT TOP x instead
    strSQL = "SELECT TOP 1 Target_Area_ID FROM tbl_Target_Species WHERE Target_Area_ID = " & TgtAreaID & ";"
            
    'fetch data
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL)
    
    'assume only 1 record returned
    If rs.RecordCount > 0 Then
        IsUsedTargetArea = True
    Else
        GoTo Exit_Handler
    End If
       
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsUsedTargetArea[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:     PopulateTree
' Description:  Populate the treeview control
' Parameters:   TreeType - treeview type (string)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, June 3, 2015 - for NCPN tools
' Revisions:
'   BLC - 6/3/2015  - initial version
' ---------------------------------
Public Sub PopulateTree(TreeType As String)

On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    
    Select Case TreeType
        Case "ParkSiteFeatureTransectPlot"
            strSQL = "SELECT * FROM qry_Park_Site_Feature_Transect_Plot"
        Case "Photo"
    End Select
                   
    'fetch data
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL)
    
    'assume only 1 record returned
    If rs.RecordCount > 0 Then
        
        
        
        
    Else
        GoTo Exit_Handler
    End If
       
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateTree[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          PopulateCombobox
' Description:  Populate priority/status comboboxes
' Parameters:   cbx - combobox control to populate (ComboBox)
'               BoxType - type of combobox, priority or status (string)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
'  https://msdn.microsoft.com/en-us/library/office/ff845773.aspx
' Adapted:      Bonnie Campbell, June 3, 2015 - for NCPN tools
' Revisions:
'   BLC - 6/3/2015  - initial version
'   BLC - 10/12/2016 - fixed to set combobox recordset
' ---------------------------------
Public Sub PopulateCombobox(cbx As ComboBox, BoxType As String)

On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    
    Select Case BoxType
        Case ""
        Case "priority"
            strSQL = "SELECT ID, Priority FROM Priority ORDER BY Sequence ASC;"
        Case "status"
            strSQL = "SELECT ID, Status FROM Status ORDER BY Sequence ASC;"
    End Select
 
     'fetch data
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL)
 
    'assume only 1 record returned
    If rs.RecordCount > 0 Then
        Set cbx.Recordset = rs
    Else
        GoTo Exit_Handler
    End If
       
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateCombobox[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     GetProtocolVersion
' Description:  Retrieve protocol version, effective & retire dates
' Assumptions:  Assumes only one version of the protocol is active at once
' Parameters:   blnAllVersions - indicator if all versions should be retrieved (boolean)
' Returns:      Protocol name, version, effective & retire dates, last modified date
' Note:         To retrieve values, data must be retrieved from the array:
'                   ary(0,0) = ProtocolName
'                   ary(1,0) = Version
'                   ary(2,0) = EffectiveDate
'                   ary(3,0) = RetireDate
'                   ary(4,0) = LastModified
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, May 5, 2016 - for NCPN tools
' Revisions:
'   BLC - 5/5/2016  - initial version
' ---------------------------------
Public Function GetProtocolVersion(Optional blnAllVersions As Boolean = False) As Variant
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String, strWhere As String
    Dim Count As Integer
    Dim metadata() As Variant
   
    'handle only appropriate park codes
    If blnAllVersions Then
        strWhere = ""
    Else
        strWhere = "WHERE RetireDate IS NULL"
    End If
    
    'generate SQL
'    strSQL = "SELECT ProtocolName, Version, EffectiveDate, RetireDate, LastModified FROM Protocol " _
'                & strWHERE & ";"
    strSQL = GetTemplate("s_protocol_info", "strWHERE" & PARAM_SEPARATOR & strWhere)
    
    'fetch data
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL)
        
    If rs.BOF And rs.EOF Then GoTo Exit_Handler
        
    With rs
        .MoveLast
        .MoveFirst
        Count = .RecordCount
    
        metadata = rs.GetRows(Count)
 
        .Close
    End With
    
    'return value
    GetProtocolVersion = metadata
    
Exit_Handler:
    Set rs = Nothing
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetProtocolVersion[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     GetSOPMetadata
' Description:  Retrieve SOP metadata (abbreviation code, #, version, effective date)
' Assumptions:  Assumes only one active/effective SOP # for a given area
' Parameters:   area - area covered by the SOP (string)
' Returns:      SOP metadata - Code, SOP #, Version, EffectiveDate
' Note:         To retrieve value, data must be retrieved from the array:
'                   ary(0,0) = SOP #
'               Assuming there is only one matching SOP for each area
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, May 5, 2016 - for NCPN tools
' Revisions:
'   BLC - 5/5/2016  - initial version
'   BLC - 5/11/2016 - revised to getting full SOP metadata vs. number only
' ---------------------------------
Public Function GetSOPMetadata(area As String) As Variant
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
        
    'generate SQL
    '---------------------------------------------------------------------
    ' NOTE: use * vs % for the LIKE wildcard
    '       if it is not used strSQL will work in a query directly,
    '       but will fail to return records via a VBA recordset
    '       So    "...LIKE '" & LCase(area) & "*';"   works
    '       But   "...LIKE '" & LCase(area) & "%';"   does not (except in direct Query SQL)
    '
    ' c.f.  Hans Up, May 17, 2011 & discussion
    '       http://stackoverflow.com/questions/6037290/use-of-like-works-in-ms-access-but-not-vba
    '---------------------------------------------------------------------
    strSQL = GetTemplate("s_sop_metadata", "area" & PARAM_SEPARATOR & LCase(area))
    
    'fetch data
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL, dbOpenDynaset)
        
    'return value
    Set GetSOPMetadata = rs
    
Exit_Handler:
    Set rs = Nothing
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetSOPNum[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     GetParkID
' Description:  Retrieve the ID associated with a park
' Assumptions:  -
' Parameters:   ParkCode - 4 character park designator (string)
' Returns:      ID - unique park identifier (long)
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, June 28, 2016 - for NCPN tools
' Revisions:
'   BLC - 6/28/2016  - initial version
'   BLC - 1/12/2017  - revised to use GetRecords() vs. GetTemplate()
' ---------------------------------
Public Function GetParkID(ParkCode As String) As Long
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim ID As Long
   
    'handle only appropriate park codes
    If Len(ParkCode) <> 4 Then
        GoTo Exit_Handler
    End If
    
    'generate SQL
'    strSQL = GetTemplate("s_park_id", "ParkCode" & PARAM_SEPARATOR & ParkCode)
            
    'fetch data
'    Set db = CurrentDb
    Set rs = GetRecords("s_park_id") 'db.OpenRecordset(strSQL)

    If rs.BOF And rs.EOF Then GoTo Exit_Handler

    rs.MoveLast
    rs.MoveFirst
    
    If Not (rs.BOF And rs.EOF) Then
        ID = rs.Fields("ID")
    End If
    
    rs.Close
    
    'return value
    GetParkID = ID
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetParkID[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Sub:          ToggleIsActive
' Description:  Toggle IsActive button click actions
' Assumptions:  -
' Parameters:   Context - form context for the action (string)
'               ID - id of record to toggle (long)
'               IsActive - state to change IsActiveFlag to (Byte), 0 - active, 1 - inactive
'                          optional for ModWentworth scale retire date
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 20, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/20/2016 - initial version
'   BLC - 6/28/2016 - shifted from ContactList form to mod_App_Data
'   BLC - 10/20/2016 - added ModWentworth retire date toggle
'   BLC - 10/24/2016 - revised to use SetRecord()
' ---------------------------------
Public Sub ToggleIsActive(Context As String, ID As Long, Optional IsActive As Byte)
On Error GoTo Err_Handler
    
'    Dim strSQL As String
'
'    Select Case Context
'        Case "Contact"
'            strSQL = GetTemplate("u_contact_isactive_flag", _
'                      "IsActiveFlag" & PARAM_SEPARATOR & IsActive & _
'                      "|ID" & PARAM_SEPARATOR & ID)
'        Case "Site"
'            strSQL = GetTemplate("u_site_isactive_flag", _
'                      "IsActiveFlag" & PARAM_SEPARATOR & IsActive & _
'                      "|ID" & PARAM_SEPARATOR & ID)
'        Case "ModWentworthScale"
'            strSQL = GetTemplate("u_mod_wentworth_retireyear", _
'                      "RetireDate" & PARAM_SEPARATOR & Date & "|ID" & _
'                      PARAM_SEPARATOR & ID)
'    End Select
'
'    DoCmd.SetWarnings False
'    DoCmd.RunSQL (strSQL)
'    DoCmd.SetWarnings True
    
    Dim Template As String
    
    Select Case Context
        Case "Contact"
            Template = "u_contact_isactive_flag"
        Case "Site"
            Template = "u_site_isactive_flag"
        Case "ModWentworthScale"
            Template = "u_mod_wentworth_retireyear"
            
    End Select
    
    Dim Params(0 To 3) As Variant
    
    Params(0) = Template
    Params(1) = ID
    Params(2) = IIf(InStr(Template, "wentworth") > 0, Year(Date), IsActive)
        
    SetRecord Template, Params
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ToggleIsActive[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          ToggleSensitive
' Description:  Toggle Sensitive button click actions
' Assumptions:  -
' Parameters:   Context - form context for the action (string)
'               ID - id of record to toggle (long)
'               Sensitive - state to change SensitiveFlag to (Byte), 0 - active, 1 - inactive
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, July 30, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/30/2016 - initial version
' ---------------------------------
Public Sub ToggleSensitive(Context As String, ID As Long, Sensitive As Byte)
On Error GoTo Err_Handler
    
    Dim Template As String
    
    Template = IIf(Sensitive = 1, "i_", "d_")
    
    Template = Template & "Sensitive" & Context
    
    If Right(Template, 1) <> "s" Then Template = Template & "s"
    
'    Select Case Context
'        Case "Locations"
''            strSQL = GetTemplate("u_location_sensitive_flag", _
''                      "SensitiveFlag" & PARAM_SEPARATOR & Sensitive & _
''                      "|ID" & PARAM_SEPARATOR & ID)
'            strToggle = strToggle & "Sensitive" & Context & "s"
'        Case "Species"
'            strSQL = GetTemplate("u_species_sensitive_flag", _
'                      "SensitiveFlag" & PARAM_SEPARATOR & Sensitive & _
'                      "|ID" & PARAM_SEPARATOR & ID)
'    End Select

'    DoCmd.SetWarnings False
'    DoCmd.RunSQL (strSQL)
'    DoCmd.SetWarnings True
    
    Dim Params(0 To 3) As Variant
    
    Params(0) = Template
    Params(1) = ID
    Params(2) = Sensitive
        
    SetRecord Template, Params
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ToggleSensitive[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          GetRecords
' Description:  Retrieve records based on template
' Assumptions:  -
' Parameters:   Template - SQL template name (string)
' Returns:      rs - data retrieved (recordset)
' Throws:       none
' References:
'   user1938742, October 17, 2014
'   http://stackoverflow.com/questions/26422970/run-query-with-parameters-and-display-in-listbox-ms-access-2013
' Source/date:  Bonnie Campbell, July 26, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/26/2016 - initial version
'   BLC - 9/22/2016 - added templates
'   BLC - 1/9/2017 - added templates
'   BLC - 2/7/2017 - added template - s_location_with_loctypeID_sensitivity
'   BLC - 3/28/2017 - added upland templates, removed big rivers templates
'   BLC - 3/30/2017 - added option for non-parameterized queries (Else)
' ---------------------------------
Public Function GetRecords(Template As String) As DAO.Recordset
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    
    With db
        Set qdf = .QueryDefs("usys_temp_qdf")
        
        With qdf
        
            'check if record exists in site
            .SQL = GetTemplate(Template)
        
            Select Case Template
                        
                Case "s_access_level"
                    '-- required parameters --
                    .Parameters("lvl") = TempVars("tempLvl")
                    
                    'clear the tempvar
                    TempVars.Remove "tempLvl"
                                                                                                               
                Case "s_get_parks"
                    '-- required parameters --
                                                    
'                Case "s_mod_wentworth_for_eventyr"
'                    '-- required parameters --
'                    'default event year to current year if not passed in
'                    .Parameters("eventyr") = Nz(TempVars("EventYear"), Year(Now))
                
                Case "s_park_id"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
           
                Case "s_template_num_records"
                    '-- required parameters --
                    
                
                Case "s_top_rooted_species_last_year_by_park"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
    
                    'revise TOP X --> 99 is replaced by # species to return (from datasheet defaults)
                    '             -->  8 is replaced by # blanks to return (")
                    .SQL = Replace(Replace(.SQL, 99, TempVars("TopSpecies")), 8, TempVars("TopBlanks"))
                    
                Case "qa_ndc_nodatacollected_fuels1000hr_transectA"
                Case "qa_ndc_nodatacollected_fuels1000hr_transectB"
                Case "qa_ndc_nodatacollected_fuels1000hr_transectC"
                Case "qa_ndc_nodatacollected_fuels1000hr_transectD"
                Case "qa_ndc_nodatacollected_saplings"
                Case "qa_ndc_nodatacollected_seedlings"
                Case "qa_ndc_nodatacollected_si_disturbance"
                Case "qa_ndc_nodatacollected_si_exotics"
                Case "qa_ndc_notrecorded_census"
                Case "qa_ndc_notrecorded_exoticfreq"
                Case "qa_ndc_notrecorded_fuels1000hr"
                Case "qa_ndc_notrecorded_fuels1000hr_transectA"
                Case "qa_ndc_notrecorded_fuels1000hr_transectB"
                Case "qa_ndc_notrecorded_fuels1000hr_transectC"
                Case "qa_ndc_notrecorded_fuels1000hr_transectD"
                Case "qa_ndc_notrecorded_saplings"
                Case "qa_ndc_notrecorded_seedlings"
                Case "qa_ndc_notrecorded_shrubs"
                Case "qa_ndc_notrecorded_si_disturbance"
                Case "qa_ndc_notrecorded_si_exotics"
                
                Case "qa_ndc_fuels1000hr_transects"
                
                Case "qa_ndc_fuels1000hr_transects_by_plot"
                    '-- required parameters --
                    .Parameters("pkid") = TempVars("ParkCode")
                    .Parameters("pid") = TempVars("plotID")
                Case "qa_ndc_nodata_census_by_plot"
                    '-- required parameters --
                    .Parameters("pkid") = TempVars("ParkCode")
                    .Parameters("pid") = TempVars("plotID")
                
                Case "qa_ndc_nodata_exoticfreq_by_plot"
                    '-- required parameters --
                    .Parameters("pkid") = TempVars("ParkCode")
                    .Parameters("pid") = TempVars("plotID")
                
                Case "qa_ndc_nodata_fuels1000hr_by_plot"
                    '-- required parameters --
                    .Parameters("pkid") = TempVars("ParkCode")
                    .Parameters("pid") = TempVars("plotID")

                Case "s_tsys_datasheet_defaults"
                    '-- required parameters --
'                    .Parameters("parkID") = TempVars("ParkID")
                
                Case Else
                    'handle other non-parameterized queries
            End Select
            
            Set rs = .OpenRecordset(dbOpenDynaset)
            
        End With
        
    End With
    
    Set GetRecords = rs
    
Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetRecords[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Function:     SetRecord
' Description:  Insert/update/delete record based on template
' Assumptions:  -
' Parameters:   template - SQL template name (string)
'               params - array of parameters for template (variant)
' Returns:      id - ID of record inserted, updated, deleted (long integer)
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, July 26, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/26/2016 - initial version
'   BLC - 9/21/2016 - updated i_login parameters
'   BLC - 10/24/2016 - added flag templates (contact, site, mod wentworth)
'   BLC - 10/28/2016 - updated TempVars("ContactID") -> TempVars("AppUserID"), updated i_task
'   BLC - 1/24/2017 - added IsNPS flag parameter for contacts
'   BLC - 3/24/2017 - set SkipRecordAction = False for uplands, removed unused big rivers cases,
'                     added uplands cases, delete cases
'   BLC - 3/29/2017 - added FieldOK, FieldCheck, Dependencies parameters for templates
' ---------------------------------
Public Function SetRecord(Template As String, Params As Variant) As Long
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim SkipRecordAction As Boolean
    Dim ID As Long
    
    'exit w/o values
    If Not IsArray(Params) Then GoTo Exit_Handler
    
    'default <-- upland does not have RecordAction table implemented so skip!
    SkipRecordAction = True 'False
            
    'default ID (if not set as param)
    ID = 0
    
    Set db = CurrentDb
    
    With db
        Set qdf = .QueryDefs("usys_temp_qdf")
        
        With qdf
        
            'check if record exists in site
            .SQL = GetTemplate(Template)
            
            '-------------------
            ' set SQL parameters --> .Parameters("") = params()
            '-------------------
            
            '-------------------------------------------------------------------------
            ' NOTE:
            '   param(0) --> reserved for record action RefTable (ReferenceType)
            '   last param(x) --> used as record ID for updates
            '-------------------------------------------------------------------------
            Select Case Template
            
        '-----------------------
        '  INSERTS
        '-----------------------
                Case "i_num_records"
                    '-- required parameters --
                    .Parameters("rid") = Params(1)  'record ID
                    .Parameters("num") = Params(2)  'number of records
                    .Parameters("fok") = Params(3)  'field ok? (QC pass/fail)
                    
                Case "i_template"
                    '-- required parameters --
                    .Parameters("tname") = Params(1)        'TemplateName
                    .Parameters("contxt") = Params(2)       'Context
                    '.Parameters("tmpl").Type = dbMemo       'set it to a memo field
                    'Limit template SQL to 255 characters to avoid
                    'error 3271 SetRecord mod_App_Data Invalid property value.
                    'templates > 255 characters must be edited directly in the table
                    .Parameters("tmpl") = Left(Params(3), 255) 'TemplateSQL
                    .Parameters("rmks") = Params(4)         'Remarks
                    .Parameters("effdate") = Params(5)      'EffectiveDate
                    .Parameters("cid") = Params(6)          'CreatedBy_ID (contactID)
                    .Parameters("prms") = Params(7)         'Params
                    .Parameters("syntx") = Params(8)        'Syntax
                    .Parameters("vers") = Params(9)         'Version
                    .Parameters("sflag") = Params(10)       'IsSupported
                    .Parameters("lmid") = TempVars("AppUserID") 'lastmodifiedID
                    .Parameters("fqc") = Params(11)         'FieldCheck
                    .Parameters("fok") = Params(12)         'FieldOK
                    .Parameters("dep") = Params(13)         'Dependencies
                
        '-----------------------
        '  UPDATES
        '-----------------------
                Case "u_num_records"
                    '-- required parameters --
                    .Parameters("rid") = Params(1)
                    .Parameters("num") = Params(2)
                    .Parameters("fok") = Params(3)
                    
                Case "u_template"
                    '-- required parameters --
                    .Parameters("id") = Params(1)
                
        '-----------------------
        '  DELETES
        '-----------------------
                Case "d_num_records_all"
                    '-- required parameters --
                
                Case "d_num_records"
                    '-- required parameters --
                    .Parameters("rid") = Params(1)
            
            End Select
            
            .Execute dbFailOnError
                
    ' -------------------
    '  Record Action
    ' -------------------
            'handle unrecorded actions & those which don't generate an ID
            If SkipRecordAction Then GoTo Exit_Handler
            
            If ID = 0 Then
                'retrieve identity
                ID = db.OpenRecordset("SELECT @@IDENTITY;")(0)
            End If
            
            'set record action
            .SQL = GetTemplate("i_record_action")
                                            
            '-- required parameters --
            .Parameters("RefTable") = Params(0)
            .Parameters("RefID") = ID
            .Parameters("ID") = TempVars("AppUserID") 'TempVars("ContactID")
            .Parameters("Activity") = "DE"
            .Parameters("ActionDate") = CDate(Format(Now(), "YYYY-mm-dd hh:nn:ss AMPM"))
                                
            .Execute dbFailOnError
            
            'cleanup
            .Close
        
        End With

        SetRecord = ID
    End With
                
Exit_Handler:
    'cleanup
    Set qdf = Nothing
    Set db = Nothing

    Exit Function
Err_Handler:
    Select Case Err.Number

      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetRecord[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Function:     UpsertRecord
' Description:  Handle insert/update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
'   gecko_1, February 10, 2005
'   http://www.access-programmers.co.uk/forums/showthread.php?t=81221
'   Khinsu, August 19, 2013
'   http://stackoverflow.com/questions/18317059/how-to-test-if-item-exists-in-recordset
'   HansUp, April 4, 2013
'   http://stackoverflow.com/questions/15823687/findfirst-vba-access2010-unbound-form-runtime-error
' Source/date:  Bonnie Campbell, July 28, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/28/2016 - initial version
'   BLC - 9/1/2016  - added vegwalk, photo
'   BLC - 10/4/2016 - added template, adjusted for form w/o list
'   BLC - 10/14/2016 - updated to accommodate non-users for contacts
'   BLC - 1/9/2017 - revised retrieve ID from ContactID to ID, revised i_event to use TempVar("SiteID")
'   BLC - 2/1/2017 - handle form upserts for forms w/o lists/msg & msg icons
'   BLC - 2/3/2017 - location adjustments
'   BLC - 3/27/2017 - removed big rivers cases, replaced w/ upland cases
' ---------------------------------
Public Sub UpsertRecord(ByRef frm As Form)
On Error GoTo Err_Handler
    
' ----------------------------------------------------------------------------------
'    1) Click to edit
'       a) populates form fields
'       b) tbxID is set
'
'       c) change values --> i) compare against existing values
'                           ii) no existing values match ==> update
'                           iii) existing values match ==> message no change
'
'   2) Enter new values
'       a) enables save button
'       b) click save -->   i) compare against existing values
'                           ii) no existing values match ==> insert
'                           iii) existing values match ==> message no change
' ----------------------------------------------------------------------------------
    
    Dim DoAction As String, strCriteria As String, strTable As String
    Dim NoList As Boolean
    Dim obj As Object
    
    'use generic object to handle multiple obj types
    With obj
    
        'default
        NoList = False
        strTable = frm.Name
    
        Select Case frm.Name
            
            Case "Template"
                'Dim tpl As New Template
                Dim tpl As Template
                
                With tpl
                    .IsSupported = 1
                    .Context = ""
                    .EffectiveDate = Date
                    .Remarks = ""
                    .TemplateName = ""
                    .Version = ""
                    .TemplateSQL = ""
                    .Syntax = ""
    
                End With
                
                'set the generic object --> Template
                Set obj = tpl
                
                'cleanup
                Set tpl = Nothing
                           
            Case "TemplateAdd"
                'Dim tpl As New Template
                
                With tpl
                    .TemplateName = frm!tbxTemplate
                    .Context = .TemplateName
                    .IsSupported = 1 '.IsSupported default = 1 (i.e. yes)
                    .Version = frm!tbxVersion
                    .Syntax = frm!cbxSyntax
                    .TemplateSQL = frm!tbxTemplateSQL
                    .EffectiveDate = frm!tbxEffectiveDate
                    '.Params handled when .TemplateSQL set
                    '.Params = GetParamsFromSQL(.TemplateSQL)
                    .Remarks = frm!tbxRemarks
                    .ContactID = TempVars("AppUserID")
                    
                    'set the generic object --> Transducer
                    Set obj = tpl
                    
                    'cleanup
                    Set tpl = Nothing
                End With
                
                'inserts only, no ID?
                NoList = True
                
            

            Case Else
                GoTo Exit_Handler
        End Select
                
        'set insert/update based on whether its an edit or new entry
        DoAction = IIf(frm!tbxID.value > 0, "u", "i")
        
        If NoList Then
                    
            'form doesn't contain list subform or message/icon fields
            'so cut to the chase -> do nothing here
            
        Else
        
            'check if the record already exists by checking event list form records
            'event list form pulls active records for park, river segment
            Dim rs As DAO.Recordset
            
            Set rs = frm!list.Form.RecordsetClone
            rs.FindFirst strCriteria
            
            If rs.NoMatch Then
                ' --- INSERT ---
                frm!lblMsg.ForeColor = lngLime
                frm!lblMsgIcon.ForeColor = lngLime
                frm!lblMsgIcon.Caption = StringFromCodepoint(uDoubleTriangleBlkR)
                frm!lblMsg.Caption = IIf(DoAction = "i", "Inserting new record...", "Updating record...")
            Else
                ' --- UPDATE ---
                'record already exists & ID > 0
                
                'retrieve ID
                If frm!tbxID.value = rs("ID") Then 'rs("Contact.ID") Then
                    'IDs are equivalent, just change the data
                    frm!lblMsg.ForeColor = lngLime
                    frm!lblMsgIcon.ForeColor = lngLime
                    frm!lblMsgIcon.Caption = StringFromCodepoint(uDoubleTriangleBlkR)
                    frm!lblMsg.Caption = "Updating record..."
                Else
                    'prevent duplicate record entries
                    frm!lblMsg.ForeColor = lngYellow
                    frm!lblMsgIcon.ForeColor = lngYellow
                    frm!lblMsgIcon.Caption = StringFromCodepoint(uDoubleTriangleBlkR)
                    frm!lblMsg.Caption = "Oops, record already exists."
                    GoTo Exit_Handler
                End If
                
            End If
        End If
        
        'T/F refers to whether the record is an update (T) or insert (F)
        obj.SaveToDb IIf(DoAction = "i", False, True)
        
        'add the action record --> DONE via SaveToDb (thru SetRecord)
        
        'set the tbxID.value ==> tbxID is a bound control, can't set it this way
        'tbxID = .ID
        'frm!tbxID.Value = obj.ID
        'frm.Controls("tbxID").Value = obj.ID
    End With
    
    'clear values & refresh display
    frm.ReadyForSave 'Application defined error? --> ensure ReadyForSave is Public Sub
    'Forms!frm.ReadyForSave
    
    'handle situations where Access is saving same record
    
    'save record changes from form first to avoid "Write Conflict" errors
    'where form & SQL are attempting to save record
    'frm.Dirty = False
    
'    If frm.Dirty Then
    If frm.Dirty And Not NoList Then
        Debug.Print "UpsertRecord " & frm.Name & " DIRTY"
        'frm.Dirty = False
        
        frm!lblMsg.ForeColor = lngYellow
        frm!lblMsgIcon.ForeColor = lngYellow
        frm!lblMsgIcon.Caption = StringFromCodepoint(uDoubleTriangleBlkR)
        frm!lblMsg.Caption = "** DIRTY **" 'UNSAVED CHANGES! **"
        
    Else
        Debug.Print "UpsertRecord " & frm.Name & " CLEAN"
    End If
        
' CHECK IF POPULATING FORM IS THE ISSUE...
'    PopulateForm frm, frm!tbxID.Value
    
'    'refresh list
'    frm!list.Requery
    
    frm.Requery
    
    'handle list forms - update messages, icon & refresh
    If Not NoList Then
        'clear messages & icon
        frm!lblMsgIcon.Caption = ""
        frm!lblMsg.Caption = ""
        
        'refresh list
        frm!list.Requery
    End If
    
    'exit
    GoTo Exit_Handler
    
Form_Without_List:
    DoAction = "i"
    Resume Next

Exit_Handler:
    'cleanup
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - UpsertRecord[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Sub

'' ---------------------------------
'' Sub:          SetObserverRecorder
'' Description:  Sets data observer & recorder
'' Assumptions:  -
'' Parameters:   obj - object to set observer/recorder on (object)
''               tbl - name of table being modified (string)
'' Returns:      -
'' Throws:       none
'' References:   -
'' Source/date:  Bonnie Campbell, August 9, 2016 - for NCPN tools
'' Adapted:      -
'' Revisions:
''   BLC - 8/9/2016 - initial version
'' ---------------------------------
'Public Sub SetObserverRecorder(obj As Object, tbl As String)
'On Error GoTo Err_Handler
'
'    'handle record actions
'    Dim act As New RecordAction
'    With act
'
'    'Recorder
'        .RefAction = "R"
'        .ContactID = obj.RecorderID
'        .RefID = obj.ID
'        .RefTable = tbl
'        .SaveToDb
'
'    'Observer
'        .RefAction = "O"
'        .ContactID = obj.ObserverID
'        .RefID = obj.ID
'        .RefTable = tbl
'        .SaveToDb
'
'    End With
'
'Exit_Handler:
'    Exit Sub
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - SetObserverRecorder[mod_App_Data])"
'    End Select
'    Resume Exit_Handler
'End Sub

'' ---------------------------------
'' Sub:          UploadCSVFile
'' Description:  Uploads data into database from CSV file
'' Assumptions:  -
'' Parameters:   strFilename - name of file being uploaded (string)
'' Returns:      -
'' Throws:       none
'' References:   -
'' Source/date:  Bonnie Campbell, September 1, 2016 - for NCPN tools
'' Adapted:      -
'' Revisions:
''   BLC - 9/1/2016 - initial version
''   BLC - 10/19/2016 - renamed to UploadCSVFile from UploadSurveyFile to genericize
'' ---------------------------------
'Public Sub UploadCSVFile(strFilename As String)
'On Error GoTo Err_Handler
'
'    'import to table
'    ImportCSV strFilename, "usys_temp_csv", True, True
'
'Exit_Handler:
'    Exit Sub
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - UploadCSVFile[mod_App_Data])"
'    End Select
'    Resume Exit_Handler
'End Sub

' ---------------------------------
' Function:          FetchAddlData
' Description:  Retrieves additional data field(s)
' Assumptions:
'               fields are delimited w/ a pipe (|)
' Parameters:   tbl - name of table to retrieve from (string)
'               field(s) - name of field to retrieve (string)
'               id - record to retrieve's ID (long)
' Returns:      field value(s) for record (DAO.Recordset)
' Throws:       none
' References:
'   Steven Thomas, November 28, 2011
'   https://blogs.office.com/2011/11/28/display-real-time-information-with-the-controltip-property/
' Source/date:  Bonnie Campbell, September 13, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/13/2016 - initial version
' ---------------------------------
Public Function FetchAddlData(tbl As String, Fields As String, ID As Long) As DAO.Recordset
On Error GoTo Err_Handler
    
    'values are required --> exit if not
    If Len(tbl) = 0 Or Len(Fields) = 0 Or Not (ID > 0) Then GoTo Exit_Handler
    
    'begin retrieval
    Dim field As String
    Dim strFields As String
    Dim strSQL As String
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    
    Set db = CurrentDb
    
    With db
        Set qdf = .QueryDefs("usys_temp_qdf")
        
        With qdf
            
            'check for multiple fields
            If InStr(Fields, "|") > 0 Then
                Dim aryFlds() As String
                Dim i As Integer
                
                aryFlds = Split(Fields, "|")
                
                For i = 0 To UBound(aryFlds)
                    strFields = aryFlds(i) & ","
                Next
                
                'remove extra comma
                strFields = IIf(Right(strFields, 1) = ",", RTrim(strFields), strFields)
            
            Else
                
                strFields = Fields
            End If
            
            'base
            strSQL = "SELECT " & strFields & " FROM " & tbl & " WHERE ID = " & ID & ";"
            
            'update the query SQL
            .SQL = strSQL
            
            Dim rs As DAO.Recordset

            Set rs = .OpenRecordset
                        
            'send results
            Set FetchAddlData = rs
            
            'cleanup
            Set rs = Nothing
            Set qdf = Nothing
            Set db = Nothing

        End With
    End With
    

Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FetchAddlData[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Sub:          GetHierarchyLevel
' Description:  Determine the hierarchy level set
' Assumptions:  -
' Parameters:   -
' Returns:      lvl - maximum level set in the application (string)
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, January 1, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/11/2017 - initial version
' ---------------------------------
Public Function GetHierarchyLevel() As String
On Error GoTo Err_Handler
    
    Dim lvl As String
    
    'default
    lvl = ""
    
    If Not TempVars("ParkCode") Is Nothing Then
        lvl = "park"
        If Not TempVars("River") Is Nothing Then
            lvl = "river"
            If Not TempVars("SiteCode") Is Nothing Then
                lvl = "site"
                If Not TempVars("Feature") Is Nothing Then
                    lvl = "feature"
                End If
            End If
        End If
    End If

    GetHierarchyLevel = lvl

Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetHierarchyLevel[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Sub:          RunPlotCheck
' Description:  Run plot check queries
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 27, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/27/2017 - initial version
'   BLC - 3/29/2017 - adjusted to accommodate FieldOK (pass/fail/unknown) values
'   BLC - 3/30/2017 - handle dependencies (queries dependent on queries)
' ---------------------------------
Public Function RunPlotCheck()
On Error GoTo Err_Handler

'    Dim Template As String
    Dim rs As DAO.Recordset, rs2 As DAO.Recordset
    Dim strTemplate As String, strDeps As String, strFieldOK As String, _
        strOperator As String, strField As String, CompareTo As String
    Dim iTemplate As Integer, i As Integer, iOK As Integer, isOK As Integer
    Dim blnFieldCheck As Boolean

    'clear num records
    ClearTable "NumRecords"
    
    'initialize AppTemplates if not populated
    If g_AppTemplates Is Nothing Then GetTemplates
        
    'fetch queries
'    Template = "s_template_num_records"
    
'    Set rs = GetRecords(Template) '--> can't run first since some queries are dependent
'
'    'iterate through records
'    If Not (rs.EOF And rs.BOF) Then
'        rs.MoveFirst
'        Do Until rs.EOF
'
'            'run query & retrieve record #s
'            Set rs2 = GetRecords(rs("TemplateName"))
'
'            'handle dependencies first
'            'Dependencies = comma separated list of queries template is dependent on
'            If Len(rs2("Dependencies")) > 0 Then _
'                HandleDependentQueries rs2("Dependencies"), "run"
'
'            'add values to numrecords
'            Dim Params(0 To 3) As Variant
'
'            Params(0) = "i_num_records"
'            Params(1) = rs("ID")
'            Params(2) = rs2.RecordCount
'            Params(3) = IIf(rs("FieldOK"), 1, -1)
'
'            SetRecord "i_num_records", Params
'
'            Debug.Print Params(1) & " " & rs("TemplateName") & " " & Params(2)
'
'            rs.MoveNext
'        Loop
'    End If

    'use g_AppTemplates scripting dictionary vs. recordset to avoid missing dependencies
    'iterate through queries
    For i = 0 To g_AppTemplates.Count - 2
    
        With g_AppTemplates.Items()(i)
            strTemplate = .Item("TemplateName")
            
            Debug.Print strTemplate
            
            If Len(.Item("FieldOK")) > 0 And .Item("FieldCheck") Then _
                SetPlotCheckResult strTemplate, "insert"
'            iTemplate = .Item("ID")
'            strDeps = .Item("Dependencies")
'            strFieldOK = .Item("FieldOK")
'            blnFieldCheck = .Item("FieldCheck")
        End With
        
'        'include only templates w/ FieldCheck = 1
'        If blnFieldCheck Then
'            'handle dependencies first
'            'Dependencies = comma separated list of queries template is dependent on
'            If Len(strDeps) > 0 Then _
'                HandleDependentQueries strDeps, "run"
'
'            'run query & retrieve record #s
'            Set rs = GetRecords(strTemplate)
'
'            'default
'            isOK = 0
'
'            'add values to numrecords
'            Dim Params(0 To 3) As Variant
'
'            Params(0) = "i_num_records"
'            Params(1) = iTemplate
'            Params(2) = rs.RecordCount
'
'            If Len(strFieldOK) > 0 Then
'                'assess if field check is fulfilled
'
'                'determine comparitor
'                iOK = CInt(Right(strFieldOK, 1))
'
'                'fetch the operator
'                strOperator = Left(Right(strFieldOK, Len(strFieldOK) - InStr(strFieldOK, "]")), 1)
'
'                'fetch the field/item to check
'                strField = Replace(Left(strFieldOK, InStr(strFieldOK, "]") - 1), "[", "")
'
'                Select Case strField
'                    Case "NumRecords"
'                        CompareTo = rs.RecordCount
'                    Case Else
'                        CompareTo = strField
'                End Select
'
'                Select Case strOperator
'                    Case "="
'                        isOK = IIf(CompareTo = iOK, 1, 0)
'                    Case "<"
'                        isOK = IIf(CompareTo < iOK, 1, 0)
'                    Case ">"
'                        isOK = IIf(CompareTo > iOK, 1, 0)
'                End Select
'
'            End If
'
'            Params(3) = isOK
'
'            SetRecord "i_num_records", Params
'
'            Debug.Print Params(1) & " " & strTemplate & " " & Params(2)
'        End If
    Next
    
Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RunPlotCheck[mod_App_Data form])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Sub:          SetPlotCheckResult
' Description:  Run plot check queries
' Assumptions:  -
' Parameters:   strTemplate - template name (string)
'               action - insert or update (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 30, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/30/2017 - initial version
'   BLC - 3/29/2017 - adjusted to accommodate FieldOK (pass/fail/unknown) values
'   BLC - 3/30/2017 - handle dependencies (queries dependent on queries)
' ---------------------------------
Public Function SetPlotCheckResult(strTemplate As String, action As String)
On Error GoTo Err_Handler

    Dim rs As DAO.Recordset, rs2 As DAO.Recordset
    Dim strDeps As String, strFieldOK As String, _
        strOperator As String, strField As String, CompareTo As String
    Dim iTemplate As Long
    Dim i As Integer, iOK As Integer
    Dim blnFieldCheck As Boolean, isOK As Boolean

    'clear num records
'    ClearTable "NumRecords"
    
    'initialize AppTemplates if not populated
    If g_AppTemplates Is Nothing Then GetTemplates
        
    'use g_AppTemplates scripting dictionary vs. recordset to avoid missing dependencies
    'iterate through queries
'    For i = 0 To g_AppTemplates.Count - 1
    
'        With g_AppTemplates.Items()(i)
        With g_AppTemplates(strTemplate)
 '           strTemplate = .Item("TemplateName")
            iTemplate = .Item("ID")
            strDeps = .Item("Dependencies")
            strFieldOK = .Item("FieldOK")
            blnFieldCheck = .Item("FieldCheck")
        End With
        
        'include only templates w/ FieldCheck = 1
'        If blnFieldCheck Then
            'handle dependencies first
            'Dependencies = comma separated list of queries template is dependent on
            If Len(strDeps) > 0 Then _
                HandleDependentQueries strDeps, "run"
                            
            'run query & retrieve record #s
            Set rs = GetRecords(strTemplate)
                
            'default
            isOK = 0
                
            'add values to numrecords
            Dim Params(0 To 3) As Variant
            
            Params(0) = LCase(Left(action, 1)) & "_num_records"
            Params(1) = iTemplate
            Params(2) = rs.RecordCount
            
            If Len(strFieldOK) > 0 Then
                'assess if field check is fulfilled
                
                'determine comparitor
                iOK = CInt(Right(strFieldOK, 1))
                
                'fetch the operator
                strOperator = Left(Right(strFieldOK, Len(strFieldOK) - InStr(strFieldOK, "]")), 1)
                
                'fetch the field/item to check
                strField = Replace(Left(strFieldOK, InStr(strFieldOK, "]") - 1), "[", "")
                
                Select Case strField
                    Case "NumRecords"
                        CompareTo = rs.RecordCount
                    Case Else
                        CompareTo = strField
                End Select
            
                Select Case strOperator
                    Case "="
                        isOK = IIf(CompareTo = iOK, 1, 0)
                    Case "<"
                        isOK = IIf(CompareTo < iOK, 1, 0)
                    Case ">"
                        isOK = IIf(CompareTo > iOK, 1, 0)
                End Select
            
            End If
            
            Params(3) = isOK
            
            'clear original value
            DeleteRecord "NumRecords", iTemplate, False
            
            SetRecord "i_num_records", Params
            
            Debug.Print Params(1) & " " & strTemplate & " " & Params(2)
 '       End If
    'Next
    
Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetPlotCheckResult[mod_App_Data form])"
    End Select
    Resume Exit_Handler
End Function


Public Function test()

    'HandleDependentQueries "68,60,69,70,71,72,73,74,75", "run"
    'HandleDependentQueries "68,60,69,70,71,72,73,74,75", "remove"
    RemoveTemplateQueries
 
    'Set g_AppTemplates = Nothing
    'GetTemplates
 
    'RunPlotCheck
    
    'GetTemplateIDs
 
End Function