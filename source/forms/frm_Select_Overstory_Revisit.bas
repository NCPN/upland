Version =21
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5040
    DatasheetFontHeight =9
    ItemSuffix =11
    Left =3795
    Top =-12315
    Right =9300
    Bottom =-8085
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x9d2210c6b41ee340
    End
    Caption ="Select for Plot Revisit Data Sheet"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
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
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin Section
            Height =3600
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =10
                    ListWidth =3600
                    Left =2280
                    Top =1080
                    Width =840
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"8\""
                    Name ="cbxPark"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT tbl_Locations.Unit_Code, tlu_Parks.ParkName FROM tlu_Parks INNER"
                        " JOIN tbl_Locations ON tlu_Parks.ParkCode = tbl_Locations.Unit_Code ORDER BY tbl"
                        "_Locations.Unit_Code;"
                    ColumnWidths ="720;2880"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =960
                            Top =1080
                            Width =1260
                            Height =245
                            FontWeight =700
                            Name ="lblPark"
                            Caption ="Select a Park"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =720
                    Top =540
                    Width =3615
                    Height =360
                    FontSize =12
                    FontWeight =700
                    Name ="lblTitle"
                    Caption ="Overstory Revisit Data Sheet"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =600
                    Top =2580
                    Width =1395
                    Height =300
                    TabIndex =3
                    Name ="btnClose"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"

                    HoverColor =65280
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =540
                    Left =2280
                    Top =1560
                    Width =840
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"3\";\"2\""
                    Name ="cbxPlotID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Plot_ID FROM tbl_locations WHERE [Unit_Code] = 'ARCH' ORDER BY Plot_ID"
                    ColumnWidths ="540"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1575
                            Top =1560
                            Width =645
                            Height =245
                            FontWeight =700
                            Name ="lblPlotID"
                            Caption ="Plot ID"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2880
                    Top =2580
                    Height =300
                    TabIndex =4
                    Name ="btnReport"
                    Caption ="Preview Report"
                    OnClick ="[Event Procedure]"

                    HoverColor =65280
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =510
                    Left =2280
                    Top =2040
                    Width =840
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"3\";\"2\""
                    Name ="cbxVisitYear"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT Year([Start_Date]) AS Visit_Year FROM tbl_Locations LEFT JOIN tb"
                        "l_Events ON tbl_Locations.Location_ID = tbl_Events.Location_ID WHERE (((Year([St"
                        "art_Date])) Is Not Null)) AND [Unit_Code] = 'ARCH' AND [Plot_ID] = 18 ORDER BY Y"
                        "ear([Start_Date]);"
                    ColumnWidths ="510"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1260
                            Top =2040
                            Width =960
                            Height =245
                            FontWeight =700
                            Name ="lblVisitYear"
                            Caption ="Visit Year"
                        End
                    End
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
' MODULE:       Form_frm_Select_Overstory_Revisit
' Level:        Form module
' Version:      1.02
' Description:  data functions & procedures specific to oak exotic frequency monitoring
'
' Source/date:  Russ DenBleyker, unknown
' Adapted by:   Bonnie Campbell, 3/8/2016
' Revisions:    RDB - unknown  - 1.00 - initial version
'               BLC - 3/8/2016 - 1.01 - added documentation
'               BLC - 2/2/2018 - 1.02 - revise to enable dropdowns based on park > plot > visit year
'                                       rename controls for consistency (lbl, btn, cbx)
'                                       updated & cleanup code
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Private m_Title As String
Private m_Directions As String
Private m_CallingForm As String

'---------------------
' Event Declarations
'---------------------
Public Event InvalidTitle(Value As String)
Public Event InvalidDirections(Value As String)
Public Event InvalidCallingForm(Value As String)

'---------------------
' Properties
'---------------------
Public Property Let Title(Value As String)
    If Len(Value) > 0 Then
        m_Title = Value

        'set the form title & caption
        Me.lblTitle.Caption = m_Title
        Me.Caption = m_Title
    Else
        RaiseEvent InvalidTitle(Value)
    End If
End Property

Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Let Directions(Value As String)
    If Len(Value) > 0 Then
        m_Directions = Value

        'set the form directions
        'Me.lblDirections.Caption = m_Directions
    Else
        RaiseEvent InvalidDirections(Value)
    End If
End Property

Public Property Get Directions() As String
    Directions = m_Directions
End Property

Public Property Let CallingForm(Value As String)
        m_CallingForm = Value
End Property

Public Property Get CallingForm() As String
    CallingForm = m_CallingForm
End Property

'---------------------
' Methods
'---------------------

' ---------------------------------
' Sub:          Form_Open
' Description:  form opening actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, February 2, 2018 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 2/2/2018 - initial version
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    'default
    Me.CallingForm = "frm_Data_Entry"
        
    If Len(Nz(Me.OpenArgs, "")) > 0 Then Me.CallingForm = Me.OpenArgs

    'minimize calling form
    ToggleForm Me.CallingForm, -1
    
    'disable PlotID & Visit Year until park/plot ID set
    Me.cbxPlotID.Enabled = False
    Me.cbxVisitYear.Enabled = False

    'hover
    btnReport.HoverColor = lngGreen
    btnClose.HoverColor = lngGreen

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Form_Open[frm_Select_Overstory_Revisit form])"
    End Select
    Resume Exit_Handler

End Sub

' ---------------------------------
' Sub:          Form_Load
' Description:  form loading actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, February 2, 2018 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 2/2/2018 - initial version
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler


Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[frm_Select_Overstory_Revisit form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Current
' Description:  form loading actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, February 2, 2018 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 2/2/2018 - initial version
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler

    'enable report button only when values are set
    If Len(cbxPark) > 0 And cbxPlotID > 0 And cbxVisitYear > 0 Then
        btnReport.Enabled = True
    Else
        btnReport.Enabled = False
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[frm_Select_Overstory_Revisit form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxPark_AfterUpdate
' Description:  Handles park code after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Russ DenBleyker, unknown
' Adapted:      -
' Revisions:
'   RDB, unknown  - initial version
'   BLC, 3/08/2016 - added documentation
'   BLC, 2/02/2018 - revised control name cbxPark vs. Park_Code, enable plot ID when park is set
'   AZ,  3/23/2022 - changed query for plot drop down, but it needs to be in the On Change event
' ---------------------------------
Private Sub cbxPark_AfterUpdate()
On Error GoTo Err_Handler

    Me!cbxPlotID = Null
    If Not IsNull(Me!cbxPark) Then
      
      Me!cbxPlotID.RowSource = "SELECT Plot FROM tbl_Revisit_List " _
            & "WHERE [PARK] = '" & Me!cbxPark & "' ORDER BY Plot"
      Me!cbxPlotID.Requery
      
    Else
      MsgBox "You must select a park!"
    End If
    
    'enable/disable
    ToggleControls
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxPark_AfterUpdate[Form_frm_Select_Overstory_Revisit])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxPlotID_AfterUpdate
' Description:  Handles plot ID after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, February 2, 2018 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC, 2/2/2018 - initial version
' ---------------------------------
Private Sub cbxPlotID_AfterUpdate()
On Error GoTo Err_Handler

    Me!cbxVisitYear = Null
    If Not IsNull(Me!cbxPark) And Not IsNull(Me!cbxPlotID) Then
    
      Me!cbxVisitYear.RowSource = "SELECT DISTINCT Year([Start_Date]) AS Visit_Year " _
            & "FROM tbl_Locations " _
            & "LEFT JOIN tbl_Events ON tbl_Locations.Location_ID = tbl_Events.Location_ID " _
            & "WHERE (((Year([Start_Date])) Is Not Null)) " _
            & "AND [Unit_Code] = '" & Me!cbxPark & "' " _
            & "AND [Plot_ID] = " & Me.cbxPlotID & " " _
            & "ORDER BY Year([Start_Date]);"
    
'      Me!cbxVisitYear.RowSource = "SELECT Visit_Year FROM qry_sel_Visit_Year " _
'            & "WHERE [Unit_Code] = '" & Me!cbxPark & "' " _
'            & "AND [Plot_ID] = " & Me.cbxPlotID & " ORDER BY Visit_Year"
      Me!cbxVisitYear.Requery
      
    Else
      MsgBox "You must select a park!"
    End If

    'enable/disable
    ToggleControls
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxPlotID_AfterUpdate[Form_frm_Select_Overstory_Revisit])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxVisitYear_AfterUpdate
' Description:  Handles plot ID after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, February 2, 2018 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC, 2/2/2018 - initial version
' ---------------------------------
Private Sub cbxVisitYear_AfterUpdate()
On Error GoTo Err_Handler

    'enable/disable Report button
    If cbxVisitYear > 0 Then
        btnReport.Enabled = True
    Else
        btnReport.Enabled = False
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxVisitYear_AfterUpdate[Form_frm_Select_Overstory_Revisit])"
    End Select
    Resume Exit_Handler
End Sub
' ---------------------------------
' SUB:          btnReport_Click
' Description:  Handles btn report click actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Russ DenBleyker, unknown
' Adapted:      -
' Revisions:
'   RDB, unknown  - initial version
'   BLC, 3/8/2016 - added documentation
'   BLC, 2/2/2018 - revise control name btnReport vs. ButtonReport
' ---------------------------------
Private Sub btnReport_Click()
On Error GoTo Err_Handler

    Dim stDocName As String
    Dim stWhereCondition As String
    If IsNull(Me!cbxPark) Or IsNull(Me!cbxPlotID) Or IsNull(Me!cbxVisitYear) Then
      MsgBox "Park Code, Plot Number, and Visit Year are all required."
      Exit Sub
    End If
    stWhereCondition = "[Unit_Code] = '" & Me!cbxPark & "' AND [Plot_Id] = " & Me!cbxPlotID & "AND [Visit_Year] = '" & Me!cbxVisitYear & "'"
    stDocName = "rpt_OT_Census"
    DoCmd.OpenReport stDocName, acViewPreview, , stWhereCondition
    DoCmd.Close acForm, "frm_Select_Overstory_Revisit"

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnReport_Click[Form_frm_Select_Overstory_Revisit])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnClose_Click
' Description:  Handles close btn click actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Russ DenBleyker, unknown
' Adapted:      -
' Revisions:
'   RDB, unknown  - initial version
'   BLC, 3/8/2016 - added documentation
'   BLC, 2/2/2018 - revised control name btnClose vs. ButtonClose
' ---------------------------------
Private Sub btnClose_Click()
On Error GoTo Err_Handler

    DoCmd.Close

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnClose_Click[Form_frm_Select_Overstory_Revisit])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          ToggleControls
' Description:  Enables/disables controls based on values
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, February 2, 2018 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC, 2/2/2018 - initial version
' ---------------------------------
Private Sub ToggleControls()
On Error GoTo Err_Handler

    'default = disable
    cbxPlotID.Enabled = False
    cbxVisitYear.Enabled = False
    
    'enable when park / park & plot ID are set
    If Len(cbxPark) > 0 Then cbxPlotID.Enabled = True
    If cbxPlotID > 0 Then cbxVisitYear.Enabled = True

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ToggleControls[Form_frm_Select_Overstory_Revisit])"
    End Select
    Resume Exit_Handler
End Sub
