Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =124
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    DatasheetFontHeight =9
    ItemSuffix =13
    Left =1092
    Top =300
    Right =8292
    Bottom =6048
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x1385341e7574e340
    End
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin Section
            Height =5760
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =525
                    Left =4155
                    Top =960
                    Width =900
                    ColumnInfo ="\"\";\"\";\"10\";\"8\""
                    Name ="Park_Code"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qry_Parks.Unit_Code FROM qry_Parks; "
                    ColumnWidths ="525"
                    AfterUpdate ="[Event Procedure]"
                    OnChange ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1935
                            Top =960
                            Width =2100
                            Height =245
                            FontWeight =700
                            Name ="Select a park if desired_Label"
                            Caption ="Select a park if desired"
                            EventProcPrefix ="Select_a_park_if_desired_Label"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3135
                    Top =2940
                    Width =1334
                    Height =300
                    TabIndex =3
                    Name ="Button_Close"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =720
                    Left =4155
                    Top =1380
                    Width =900
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"3\";\"2\""
                    Name ="Visit_Date"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qry_Event_Date.Visit_Year FROM qry_Event_Date; "
                    ColumnWidths ="720"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1935
                            Top =1380
                            Width =2100
                            Height =245
                            FontWeight =700
                            Name ="Select a date if desired_Label"
                            Caption ="Select a year if desired"
                            EventProcPrefix ="Select_a_date_if_desired_Label"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1980
                    Top =300
                    Width =3060
                    Height =420
                    FontSize =14
                    FontWeight =700
                    Name ="Label6"
                    Caption ="Species Report"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3135
                    Top =2400
                    Width =1320
                    Height =345
                    TabIndex =4
                    Name ="Button_rpt_by_Park"
                    Caption ="Report by Park"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin ComboBox
                    Enabled = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4155
                    Top =1800
                    Width =720
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"3\";\"2\""
                    Name ="Plot"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Plot_ID FROM tbl_Locations WHERE Unit_Code = 'CARE' ORDER BY Plot_ID"
                    ColumnWidths ="420"
                    OnGotFocus ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1935
                            Top =1800
                            Width =2100
                            Height =245
                            FontWeight =700
                            Name ="Plot_Select_Label"
                            Caption ="Select a plot if desired"
                        End
                    End
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =5040
                    Top =1800
                    Width =1740
                    Height =420
                    ForeColor =16711680
                    Name ="lblPlotHint"
                    Caption ="Park selection required to select a plot."
                    LayoutCachedLeft =5040
                    LayoutCachedTop =1800
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =2220
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' MODULE:       Form_frm_Species_Report_Select
' Level:        Form module
' Version:      1.04
' Description:  data functions & procedures specific to species report by park
'
' Source/date:  Bonnie Campbell, 2/2/2016
' Revisions:    RDB - unknown  - 1.00 - initial version
'               BLC - 2/2/2016 - 1.01 - added documentation, no data collected integration
'               BLC - 3/7/2016 - 1.02 - fixed strSQL replacement issue in Button_rpt_by_Park_Click
'               BLC - 3/8/2016 - 1.03 - fixed old data refresh issue (Button_rpt_by_Park_Click),
'                                       added documentation & error handling to other subroutines
'               BLC - 3/9/2016 - 1.04 - added primary key & index to improve report performance (Button_rpt_by_Park_Click)
'               BLC - 3/16/2016 - 1.05 - added enabling plot dropdown when park is chosen (Park_Code_AfterUpdate, Park_Code_Change)
' =================================

' ---------------------------------
' SUB:          Plot_GotFocus
' Description:  Contains actions occuring on plot focus
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Russ DenBleyker, unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   RD - ?          - initial version
'   BLC - 3/8/2016  - added documentation & error handling
'   BLC - 3/16/2016 - deprecated but not removed - the park should now always be selected since plot is disabled until
'                     park is chosen, as a result the park code "You must select a park first." should never be called
' ---------------------------------
Private Sub Plot_GotFocus()
On Error GoTo Err_Handler

  If IsNull(Me!Park_Code) Then
    MsgBox "You must select a park first."
    Me!Park_Code.SetFocus
  End If
  
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Plot_GotFocus[Form_frm_Species_Report_Select])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Park_Code_Change
' Description:  Contains actions occuring after changing park code
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, March 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/16/2016 - initial version
' ---------------------------------
Private Sub Park_Code_Change()
On Error GoTo Err_Handler

  'plot dropdown should be disabled (default)
  Me!Plot.Enabled = False
  Me.Refresh
  
  If Not IsNull(Me!Park_Code) Then
    Me!Plot.RowSource = "SELECT Plot_ID FROM tbl_Locations WHERE Unit_Code = '" & Me!Park_Code & "' ORDER BY Plot_ID"
    Me!Plot.Requery
    'enable plot dropdown
    Me!Plot.Enabled = True
  End If
  
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Park_Code_AfterUpdate[Form_frm_Species_Report_Select])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Park_Code_AfterUpdate
' Description:  Contains actions occuring after updating park code
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Russ DenBleyker, unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   RD - ?          - initial version
'   BLC - 3/8/2016  - added documentation & error handling
'   BLC - 3/16/2016 - enabled plot dropdown when park code is selected (otherwise it's disabled)
' ---------------------------------
Private Sub Park_Code_AfterUpdate()
On Error GoTo Err_Handler

  If Not IsNull(Me!Park_Code) Then
    Me!Plot.RowSource = "SELECT Plot_ID FROM tbl_Locations WHERE Unit_Code = '" & Me!Park_Code & "' ORDER BY Plot_ID"
    Me!Plot.Requery
    'enable plot dropdown
    Me!Plot.Enabled = True
  End If
  
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Park_Code_AfterUpdate[Form_frm_Species_Report_Select])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Button_rpt_by_Park_Click
' Description:  Button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Juan Soto, October 31, 2011
'   https://accessexperts.com/blog/2011/10/31/more-alter-table-sql-statement-help/
'   Aiken, December 3, 2014
'   http://stackoverflow.com/questions/19369132/declare-and-initialize-string-array-in-vba
' Source/date:  Russ DenBleyker, unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   RD - ?          - initial version
'   BLC - 8/21/2015 - added update for underlying qry's table (temp_Sp_Rpt_by_Park_Complete)
'   BLC - 9/1/2015  - updated comment reference since underlying query qry_Sp_Rpt_by_Park_Rollup was
'                     replaced by the table temp_Sp_Rpt_by_Park_Rollup, added temporary revision
'                     of qry_Sp_Rpt_by_Park_Complete_Create_Table for filtered data (added appropriate
'                     WHERE clause) & reverted to original qdf SQL after running (for next time)
'   BLC - 3/7/2016  - fixed error where strSQL was not being replaced w/ strSQLNew causing report
'                     to ignore change in WHERE clause/prior data & not regenerate the proper report data set
'   BLC - 3/8/2016  - fixed issue where old data wasn't being refreshed in temp_Sp_Rpt_by_Park_Rollup
'                     by running qry_Sp_Rpt_by_Park_Rollup_Create_Table to refresh data from current back-end
'   BLC - 3/9/2016  - added multi-column indices to speed _Rollup query & report generation
'                     indexes performed slightly better than adding multicolum/autogenerated ID primary keys to
'                     temp_Sp_Rpt_by_Park_Complete & temp_Sp_Rpt_by_Park_Rollup tables
'                     report generation time dropped from ~5min -> ~48sec for 18212 results across all parks/years,
'                     added additional query/table statusbar messaging
'                     n.b. status bar messages are superceded by app Run Query, etc. (fix later)
'   BLC - 3/16/2016 - fix issue resulting in parameter error #3474 data type mismatch in criteria expression
'                     due to where clause in existing qry_Sp_Rpt_by_Park_Complete_Create_Table which was:
'                     WHERE Len(SpeciesYears) > Len(Replace(SpeciesYears, CStr(2014), ''))
'                     remove it and leave existing ORDER BY clause
'                     added report filter display via OpenArgs
' ---------------------------------
Private Sub Button_rpt_by_Park_Click()
On Error GoTo Err_Handler

    'set statusbar notice
    SysCmd acSysCmdSetStatus, "Running report ..."
    
    Screen.MousePointer = 11 'Hour Glass

    Dim strFilter As String, strWhere As String, strParkWhere As String, strPlotWhere As String, strYrWhere As String, strSpeciesYear As String
    Dim stDocName As String

    'defaults
    strFilter = ""
    strWhere = ""
    strParkWhere = ""
    strPlotWhere = ""
    strYrWhere = ""

    stDocName = "rpt_Species_by_Park"
    
    ' Set where condition if needed
    If (IsNull(Me!Park_Code) + IsNull(Me!Visit_Date) + IsNull(Me!Plot)) > -3 Then
      
      'park
      If Not IsNull(Me!Park_Code) Then
        strParkWhere = "Unit_Code = '" & Me!Park_Code & "'"
        strFilter = Me!Park_Code
      End If
      
      'plot --> NOTE: assumes UI will not allow plot selection w/o park
      If Not IsNull(Me!Plot) Then
        strPlotWhere = "Plot_ID = " & Me!Plot
        strFilter = strFilter & "- plot #" & Me!Plot
      End If
      
      'year
      If Not IsNull(Me!Visit_Date) Then
'        strYrWhere = "Len(SpeciesYear) > Len(Replace(SpeciesYear, CStr(" & Me!Visit_Date & "), ''))"
        '(qry_Sp_Rpt_All.Utah_Species+"-"+CStr(qry_Sp_Rpt_All.Year)) AS SpeciesYear
        strSpeciesYear = "(qry_Sp_Rpt_All.Utah_Species+' - '+CStr(qry_Sp_Rpt_All.Year))"
        strYrWhere = "Len(" & strSpeciesYear & ") > Len(Replace(" & strSpeciesYear & ", CStr(" & Me!Visit_Date & "), ''))"
        
        'set filter display
        Select Case Len(strFilter)
            Case 0 'year only
                strFilter = CStr(Me!Visit_Date)
            Case 4 'park only
                strFilter = strFilter & "-" & CStr(Me!Visit_Date)
            Case Is > 4 'park & plot
                strFilter = Replace(strFilter, "-", "-" & CStr(Me!Visit_Date) & " ")
        End Select
      
      Else
        'clear extra "-" for park & plot filter
        strFilter = Replace(strFilter, "-", "")
      End If
      
      'prepare where using string array & PrepareWhereClause
      Dim ary() As String
      ary = Split(strParkWhere & ";" & strPlotWhere & ";" & strYrWhere, ";")
      strWhere = PrepareWhereClause(ary)
      
'      If Not IsNull(Me!Park_Code) Then
'        strWhere = "Unit_Code = '" & Me!Park_Code & "'"
'        If Not IsNull(Me!Plot) Then
'          strWhere = strWhere & " And Plot_ID = " & Me!Plot
'        End If
'        If Not IsNull(Me!Visit_Date) Then
'          'strWhere = strWhere & " AND Visit_Year = " & Me!Visit_Date
'          'strWhere = strWhere & " AND " & Me!Visit_Date & " IN (replace(SpeciesYears, '|', ','))"
'          'strWhere = strWhere & " AND " & Me!Visit_Date & " LIKE SpeciesYears"
'          strWhere = strWhere & " AND Len(SpeciesYear) > Len(Replace(SpeciesYear, CStr(" & Me!Visit_Date & "), ''))"
'        End If
'      Else
'        'strWhere = "Visit_Year = " & Me!Visit_Date
'        'strWhere = Me!Visit_Date & " IN (replace(SpeciesYears, '|', ','))"
'        'WHERE Len(SpeciesYears) > Len(Replace(SpeciesYears, CStr(2014), ''));
'        strWhere = "Len(SpeciesYear) > Len(Replace(SpeciesYear, CStr(" & Me!Visit_Date & "), ''))"
'      End If
    End If
    
    'retrieve querydef
    Dim qdf As QueryDef
    Dim strSQL As String
    
    Set qdf = CurrentDb.QueryDefs("qry_Sp_Rpt_by_Park_Complete_Create_Table")
    strSQL = qdf.SQL

'SELECT DISTINCT
'qry_Sp_Rpt_All.Unit_Code,
'qry_Sp_Rpt_All.Year,
'qry_Sp_Rpt_All.Plot_ID,
'qry_Sp_Rpt_All.Master_Family,
'qry_Sp_Rpt_All.Utah_Species,
'(qry_Sp_Rpt_All.Utah_Species+"-"+CStr(qry_Sp_Rpt_All.Year)) AS SpeciesYear,
'(qry_Sp_Rpt_All.Unit_Code+"-"+CStr(qry_Sp_Rpt_All.Plot_ID)+"-"+CStr(qry_Sp_Rpt_All.Utah_Species)) AS ParkPlotSpecies,
'(qry_Sp_Rpt_All.Unit_Code+"-"+CStr(qry_Sp_Rpt_All.Utah_Species)) AS ParkSpecies,
'(qry_Sp_Rpt_All.Unit_Code+"-"+CStr(qry_Sp_Rpt_All.Plot_ID)) AS ParkPlot INTO temp_Sp_Rpt_by_Park_Complete
'FROM qry_Sp_Rpt_All
'WHERE Len(SpeciesYears) > Len(Replace(SpeciesYears, CStr(2014), ''))
'ORDER BY qry_Sp_Rpt_All.Unit_Code, qry_Sp_Rpt_All.Plot_ID, qry_Sp_Rpt_All.Master_Family, qry_Sp_Rpt_All.Utah_Species;

    'update the SQL if parameters exist
    If Len(strWhere) > 0 Then
        Dim iOrderBy As Integer
        Dim strSQLNew As String
        
        'replace ORDER with WHERE clause + ORDER
        strSQLNew = Replace(strSQL, "ORDER", " WHERE " & strWhere & " ORDER")
        qdf.SQL = strSQLNew 'was strSQL
    End If
    
    'update underlying table (temp_Sp_Rpt_by_Park_Complete is used in report's underlying table temp_Sp_Rpt_by_Park_Rollup)
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "qry_Sp_Rpt_by_Park_Complete_Create_Table", acViewNormal
    
    'update status bar
    SysCmd acSysCmdSetStatus, "Generating complete results..."
    'DoEvents
    'Application.Echo False, "Generating complete results..."
    'Application.Echo True, ""
    
    'add an index to improve report performance
    Dim strIdxSQL As String
    
    strIdxSQL = "CREATE INDEX idxParkPlotSpeciesYear ON temp_Sp_Rpt_by_Park_Complete (ParkPlotSpecies, Year)"
    CurrentDb.Execute strIdxSQL
    
    DoCmd.SetWarnings True
    
    'reset qdf SQL
    qdf.SQL = strSQL
    
    'update underlying table (temp_Sp_Rpt_by_Park_Rollup)
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "qry_Sp_Rpt_by_Park_Rollup_Create_Table", acViewNormal
    
    'update status bar
    SysCmd acSysCmdSetStatus, "Generating rollup..."
    'DoEvents
    'Application.Echo False, "Generating rollup..."
    'Application.Echo True, ""
    
    'add an index to improve report performance
    strIdxSQL = "CREATE INDEX idxParkPlotSpeciesYears ON temp_Sp_Rpt_by_Park_Rollup (ParkPlotSpecies, SpeciesYears)"
    CurrentDb.Execute strIdxSQL
    
    DoCmd.SetWarnings True
    
    'update status bar
    SysCmd acSysCmdSetStatus, "Preparing report..."
    'DoEvents
    'Application.Echo False, "Preparing report..."
    'Application.Echo True, ""
    
    'translate SQL Where for rollup --> SpeciesYear = SpeciesYears, ,qry_Sp_Rpt_All.Year = SpeciesYears, qry_Sp_Rpt_All.Utah_species = "Utah.species"
    Dim aryText() As String
    aryText = Split("SpeciesYear|SpeciesYears||qry_Sp_Rpt_All.Year|SpeciesYears||qry_Sp_Rpt_All.Utah_species|Utah_species", "||")
    strWhere = ReplaceMulti(strWhere, aryText)
    'strWhere = Replace(strWhere, Replace(strSpeciesYear, "SpeciesYear", "SpeciesYears"), "SpeciesYears")
    
    'open report --> strWhere = WHERE clause filter, strFilter = display for filter if present
    DoCmd.OpenReport stDocName, acViewPreview, , strWhere, acWindowNormal, strFilter
    
    SysCmd acSysCmdSetStatus, "Report complete."
    
    Screen.MousePointer = 1 'Standard Cursor
    'clear status bar
    SysCmd acSysCmdSetStatus, " "
    
Exit_Handler:
    Set qdf = Nothing
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Button_rpt_by_Park_Click[Form_frm_Species_Report_Select])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Button_Close_Click
' Description:  Contains actions occuring when close button is clicked
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Russ DenBleyker, unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   RD - ?          - initial version
'   BLC - 3/8/2016  - added documentation & error handling
' ---------------------------------
Private Sub Button_Close_Click()
On Error GoTo Err_Handler

    DoCmd.Close

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Button_Close_Click[Form_frm_Species_Report_Select])"
    End Select
    Resume Exit_Handler
End Sub
