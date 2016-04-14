Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    TabularFamily =124
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6660
    DatasheetFontHeight =9
    ItemSuffix =15
    Left =1335
    Top =5400
    Right =7245
    Bottom =9045
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x99c2faf85388e340
    End
    RecordSource ="qry_Fuels_1000"
    Caption ="fsub_Fuels_1000"
    BeforeInsert ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =840
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =180
                    Top =540
                    Width =840
                    Height =240
                    FontWeight =700
                    Name ="Transect_Label"
                    Caption ="Transect"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1260
                    Top =540
                    Width =1335
                    Height =240
                    FontWeight =700
                    Name ="Diameter_Label"
                    Caption ="Diameter (in)"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2880
                    Top =540
                    Width =1440
                    Height =240
                    FontWeight =700
                    Name ="Decay_Class_Label"
                    Caption ="DecayClass"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =1440
                    Top =60
                    Width =2100
                    Height =360
                    FontSize =14
                    FontWeight =700
                    Name ="Label12"
                    Caption ="1000-hr fuels"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =85
                    Left =3600
                    Top =120
                    Width =840
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label13"
                    Caption ="(> 3 in)"
                    FontName ="Tahoma"
                End
            End
        End
        Begin Section
            Height =420
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =420
                    Height =255
                    ColumnWidth =2310
                    Name ="Fuels_1000_ID"
                    ControlSource ="Fuels_1000_ID"
                    StatusBarText ="Unique record identifier - primary key"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =600
                    Top =60
                    Width =420
                    Height =255
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Event_ID"
                    ControlSource ="Event_ID"
                    StatusBarText ="Foreign key to tbl_Events"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1515
                    Top =60
                    Width =840
                    Height =255
                    ColumnWidth =2310
                    TabIndex =3
                    Name ="Diameter"
                    ControlSource ="Diameter"
                    StatusBarText ="Diameter to nearest .5 inch"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =1125
                    Left =2880
                    Top =60
                    TabIndex =4
                    Name ="Decay_Class"
                    ControlSource ="Decay_Class"
                    RowSourceType ="Value List"
                    RowSource ="\"sound\";\"rotten\";\"Not recorded\""
                    ColumnWidths ="1125"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =4680
                    Top =60
                    Width =705
                    Height =300
                    TabIndex =5
                    ForeColor =255
                    Name ="ButtonDelete"
                    Caption ="Delete"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =180
                    Top =60
                    Width =840
                    ColumnWidth =225
                    TabIndex =2
                    Name ="Transect"
                    ControlSource ="Transect"
                    RowSourceType ="Value List"
                    RowSource ="A;B;C;D"
                    StatusBarText ="Transect associated with 1000-hr fuel measurements."
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
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
' MODULE:       Form_fsub_Fuels_1000
' Level:        Form module
' Version:      1.03
' Description:  data functions & procedures specific to fuels monitoring
'
' Source/date:  Bonnie Campbell, 2/2/2016
' Revisions:    RDB - unknown  - 1.00 - initial version
'               BLC - 2/2/2016 - 1.01 - added documentation, checkbox for no species found
'               BLC - 3/18/2016 - 1.02 - added handling for 1000hr fuels A-D checkboxes for no fuels found
'               BLC - 3/23/2016 - 1.03 - revised Delete_Click to add 1000hr fuels A-D & main 1000hr
'                                        NoDataCollected records when last record deleted
' =================================

' ---------------------------------
' SUB:          Form_Load
' Description:  Handles form loading actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 2, 2016 - for NCPN tools
' Revisions:
'   BLC, 2/2/2016  - initial version
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler


Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[Form_fsub_Fuels_1000])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_BeforeInsert
' Description:  Handles form pre-insert actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Russ DenBleyker, unknown
' Adapted:      Bonnie Campbell, February 11, 2016 - for NCPN tools
' Revisions:
'   RDB, unknown    - initial version
'   BLC, 2/11/2016  - added no data collected info updates
' ---------------------------------
Private Sub Form_BeforeInsert(Cancel As Integer)
On Error GoTo Err_Handler

    If IsNull(Me!Event_ID) Then
      MsgBox "You must enter event information first."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
      GoTo Exit_Handler
    End If
    
    ' Create the GUID primary key value
    If IsNull(Me!Fuels_1000_ID) Then
        If GetDataType("tbl_Fuels_1000", "Fuels_1000_ID") = dbText Then
            Me.Fuels_1000_ID = fxnGUIDGen
        End If
    End If

    '-----------------------------------
    ' update the NoDataCollected info
    '-----------------------------------
    Dim NoData As Scripting.Dictionary
    
    'remove the no data collected record
    Set NoData = SetNoDataCollected(Me.Parent.Form.Controls("Event_ID"), "E", "Fuel-1000hr", 0)
        
    'update checkbox/rectangle
    Me.Parent.Form.Controls("cbxNo1000hr") = 0
    Me.Parent.Form.Controls("cbxNo1000hr").Enabled = False
    Me.Parent.Form.Controls("rctNo1000hr").Visible = False

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeInsert[Form_fsub_Fuels_1000])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_AfterUpdate
' Description:  Handles form post-update actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, March 18, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC, 3/18/2016  - initial version
' ---------------------------------
Private Sub Form_AfterUpdate()
On Error GoTo Err_Handler

    'handle individual transect no fuel data collected
    Check1000hrFuels

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeInsert[Form_fsub_Fuels_1000])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Transect_AfterUpdate
' Description:  Handles form post-update actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, March 18, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC, 3/18/2016    - initial version
' ---------------------------------
Private Sub Transect_AfterUpdate()
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeInsert[Form_fsub_Fuels_1000])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub Transect_BeforeUpdate(Cancel As Integer)
  Dim Veg_Type As Variant
  
  If Not IsNull(Me!Transect) And Me!Transect = "D" Then
    Veg_Type = DLookup("[Vegetation_Type]", "tbl_Locations", "[Location_ID] = '" & Me.Parent!Location_ID & "'")
    If Not IsNull(Veg_Type) And (Veg_Type = "oak scrub") Then
      MsgBox "Value is not within domain limits"
      DoCmd.CancelEvent
      SendKeys "{ESC}"
    End If
  End If
  
End Sub

Private Sub Transect_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub


Private Sub Decay_Class_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Diameter_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

' ---------------------------------
' SUB:          ButtonDelete_Click
' Description:  Handles delete button actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Russ DenBleyker, unknown
' Adapted:      Bonnie Campbell, February 11, 2016 - for NCPN tools
' Revisions:
'   RDB, unknown  - initial version
'   BLC, 2/11/2016 - added error handling, documentation, refresh checkbox/no data collected
'   BLC, 3/23/2016 - revised to add NoDataCollected records for A-D as well as main 1000hr
'                    when last record deleted
' ---------------------------------
Private Sub ButtonDelete_Click()
On Error GoTo Err_Handler

  Dim intReply As Integer
  
  intReply = MsgBox("Are you sure you want to delete this record?", vbYesNo, "Delete Record")
    If intReply = vbYes Then
      DoCmd.SetWarnings False
      DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
      DoCmd.DoMenuItem acFormBar, acEditMenu, 6, , acMenuVer70
      DoCmd.SetWarnings True
      Me.Requery
    End If

    '-----------------------------------
    ' update the NoDataCollected info IF no records now exist
    '-----------------------------------
    
    Check1000hrFuels
    
    If Me.RecordsetClone.RecordCount = 0 Then
    
        Dim NoData As Scripting.Dictionary
        
        With Me.Parent.Form
            
            'remove the no data collected record
            Set NoData = SetNoDataCollected(.Controls("Event_ID"), "E", "Fuel-1000hr", 1)
        
            'update checkbox/rectangle
            .Controls("cbxNo1000hr") = 1
            .Controls("cbxNo1000hr").Enabled = True
            .Controls("rctNo1000hr").Visible = True
            
            'update A-D
            SetNoDataCollected .Controls("Event_ID"), "E", "Fuel-1000hr-A", 1
            .Controls("cbxNo1000hrA") = 1
            .Controls("cbxNo1000hrA").Enabled = True
            .Controls("rctNo1000hrA").Visible = True
            SetNoDataCollected .Controls("Event_ID"), "E", "Fuel-1000hr-B", 1
            .Controls("cbxNo1000hrB") = 1
            .Controls("cbxNo1000hrB").Enabled = True
            .Controls("rctNo1000hrB").Visible = True
            SetNoDataCollected .Controls("Event_ID"), "E", "Fuel-1000hr-C", 1
            .Controls("cbxNo1000hrC") = 1
            .Controls("cbxNo1000hrC").Enabled = True
            .Controls("rctNo1000hrC").Visible = True
            SetNoDataCollected .Controls("Event_ID"), "E", "Fuel-1000hr-D", 1
            .Controls("cbxNo1000hrD") = 1
            .Controls("cbxNo1000hrD").Enabled = True
            .Controls("rctNo1000hrD").Visible = True
        
        End With
        
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ButtonDelete_Click[Form_fsub_Fuels_1000])"
    End Select
    Resume Exit_Handler
End Sub
