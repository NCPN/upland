Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =124
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =12600
    DatasheetFontHeight =9
    ItemSuffix =62
    Left =108
    Right =12720
    Bottom =9816
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xd4e1e7326d12e340
    End
    RecordSource ="tbl_Site_Impact"
    Caption ="frm_Canopy_Transect"
    BeforeInsert ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnKeyDown ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =255
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
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            CanGrow = NotDefault
            Height =9360
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8880
                    Top =60
                    Width =630
                    Height =180
                    ColumnWidth =2310
                    Name ="Impact_ID"
                    ControlSource ="Impact_ID"
                    StatusBarText ="Unique record identifier - primary key"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9720
                    Top =60
                    Width =630
                    Height =180
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Event_ID"
                    ControlSource ="Event_ID"
                    StatusBarText ="M. Link to tbl_Locations (Loc_ID)"

                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =4200
                    Top =180
                    Width =960
                    ColumnWidth =1035
                    TabIndex =3
                    Name ="Visit_Date"
                    ControlSource ="Visit_Date"
                    Format ="Short Date"
                    StatusBarText ="Date of visit."
                    InputMask ="99/99/0000;0;_"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3300
                            Top =180
                            Width =900
                            Height =240
                            FontWeight =700
                            Name ="Visit_Date_Label"
                            Caption ="Visit Date"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =1965
                    Left =1080
                    Top =180
                    Width =1500
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Recorder"
                    ControlSource ="Recorder"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Contacts.Contact_ID, tlu_Contacts.Last_Name, tlu_Contacts.First_Name "
                        "FROM tlu_Contacts WHERE (((tlu_Contacts.Active)=1)); "
                    ColumnWidths ="0;975;990"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =180
                            Top =180
                            Width =900
                            Height =245
                            FontWeight =700
                            Name ="Recorder_Label"
                            Caption ="Recorder"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4080
                    Top =5460
                    Width =2220
                    TabIndex =7
                    Name ="ButtonSiteSketch"
                    Caption ="Disturbance sketch/photo"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4080
                    LayoutCachedTop =5460
                    LayoutCachedWidth =6300
                    LayoutCachedHeight =5820
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin Subform
                    OverlapFlags =87
                    Left =840
                    Top =1440
                    Width =8517
                    Height =3837
                    TabIndex =6
                    Name ="Disturbance Details"
                    SourceObject ="Form.fsub_Impact_Details"
                    LinkChildFields ="Impact_ID"
                    LinkMasterFields ="Impact_ID"
                    EventProcPrefix ="Disturbance_Details"

                    LayoutCachedLeft =840
                    LayoutCachedTop =1440
                    LayoutCachedWidth =9357
                    LayoutCachedHeight =5277
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =840
                            Top =1200
                            Width =1920
                            Height =240
                            FontWeight =700
                            Name ="Disturbance Details Label"
                            Caption ="Disturbance Details"
                            EventProcPrefix ="Disturbance_Details_Label"
                            LayoutCachedLeft =840
                            LayoutCachedTop =1200
                            LayoutCachedWidth =2760
                            LayoutCachedHeight =1440
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2700
                    Top =660
                    Width =420
                    TabIndex =4
                    Name ="Percent_Top_Kill"
                    ControlSource ="Percent_Top_Kill"
                    StatusBarText ="Percentage of Gambel oak cover in plot affected by top kill."
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =660
                            Width =2460
                            Height =240
                            FontWeight =700
                            Name ="Label58"
                            Caption ="Percent Gambel Oak top kill"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5700
                    Top =660
                    Width =420
                    TabIndex =5
                    Name ="Percent_Dead"
                    ControlSource ="Percent_Dead"
                    StatusBarText ="Percentage of Gambel oak cover in plot that is completely dead."
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3360
                            Top =660
                            Width =2295
                            Height =240
                            FontWeight =700
                            Name ="Label59"
                            Caption ="Percent Gambel Oak Dead"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =1080
                    Top =6120
                    Width =8934
                    Height =3000
                    TabIndex =8
                    Name ="fsub_Dist_Exotic"
                    SourceObject ="Form.fsub_Dist_Exotic"
                    LinkChildFields ="Impact_ID"
                    LinkMasterFields ="Impact_ID"

                    LayoutCachedLeft =1080
                    LayoutCachedTop =6120
                    LayoutCachedWidth =10014
                    LayoutCachedHeight =9120
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Left =6780
                    Top =900
                    Width =2460
                    Height =480
                    BackColor =6750207
                    Name ="rctNoDisturbance"
                    OnClick ="[Event Procedure]"
                    LayoutCachedLeft =6780
                    LayoutCachedTop =900
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =1380
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =6900
                    Top =1050
                    Width =300
                    TabIndex =9
                    Name ="cbxNoDisturbance"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="No disturbances found"

                    LayoutCachedLeft =6900
                    LayoutCachedTop =1050
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =1290
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontFamily =0
                            Left =7130
                            Top =1020
                            Width =2016
                            Height =228
                            FontWeight =600
                            Name ="lblNoDistubance"
                            Caption ="No Disturbances Found"
                            ControlTipText ="No disturbances found"
                            LayoutCachedLeft =7130
                            LayoutCachedTop =1020
                            LayoutCachedWidth =9146
                            LayoutCachedHeight =1248
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Left =7860
                    Top =5580
                    Width =2100
                    Height =480
                    BackColor =6750207
                    Name ="rctNoSpecies"
                    OnClick ="[Event Procedure]"
                    LayoutCachedLeft =7860
                    LayoutCachedTop =5580
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =6060
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =7980
                    Top =5730
                    Width =300
                    TabIndex =10
                    Name ="cbxNoSpecies"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="No exotic species found"

                    LayoutCachedLeft =7980
                    LayoutCachedTop =5730
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =5970
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontFamily =0
                            Left =8210
                            Top =5700
                            Width =1650
                            Height =240
                            FontWeight =600
                            Name ="lblNoSpecies"
                            Caption ="No Species Found"
                            ControlTipText ="No exotic species found"
                            LayoutCachedLeft =8210
                            LayoutCachedTop =5700
                            LayoutCachedWidth =9860
                            LayoutCachedHeight =5940
                        End
                    End
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
' MODULE:       Form_frm_Site_Impact
' Level:        Form module
' Version:      1.02
' Description:  data functions & procedures specific to site impact monitoring
'
' Source/date:  Bonnie Campbell, 2/2/2016
' Revisions:    RDB - unknown  - 1.00 - initial version
'               BLC - 2/2/2016 - 1.01 - added documentation, checkbox for no species found
'               BLC - 2/9/2016 - 1.02 - added NoDataCollected collection changes
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
'   BLC, 2/10/2016 - add setting of public collection for NoDataCollected checkbox values
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler

    Dim dNoData As Scripting.Dictionary

    'fetch no data info & set checkboxes
    Set dNoData = GetNoDataCollected(Me.Event_ID)
    
    With dNoData
        Me.cbxNoDisturbance.Value = dNoData.item("SiteImpact-Disturbance")
        Me.cbxNoSpecies.Value = dNoData.item("SiteImpact-Exotic")
    End With
    
    'set the yellow rectangles visible if no records
    Me.rctNoDisturbance.Visible = (Me.Disturbance_Details.Form.RecordsetClone.RecordCount = 0)
    Me.rctNoSpecies.Visible = (Me.fsub_Dist_Exotic.Form.RecordsetClone.RecordCount = 0)
    
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[Form_frm_Site_Impact])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxNoSpecies_Click
' Description:  Handles checkbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 2, 2016 - for NCPN tools
' Revisions:
'   BLC, 2/2/2016  - initial version
' ----------------------------------
Private Sub cbxNoSpecies_Click()
On Error GoTo Err_Handler

    'set dictionary & db value (abs is used to drive 1 = true, 0 = false since -1 is true in access/vba)
    SetNoDataCollected Me.Event_ID, "SiteImpact-Exotic", Abs(Me.cbxNoSpecies.Value)

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxNoSpecies_Click[Form_frm_Site_Impact])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxNoDisturbance_Click
' Description:  Handles checkbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 9, 2016 - for NCPN tools
' Revisions:
'   BLC, 2/9/2016  - initial version
' ----------------------------------
Private Sub cbxNoDisturbance_Click()
On Error GoTo Err_Handler

    'set dictionary & db value (abs is used to drive 1 = true, 0 = false since -1 is true in access/vba)
    SetNoDataCollected Me.Event_ID, "SiteImpact-Disturbance", Abs(Me.cbxNoDisturbance.Value)

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxNoDisturbance_Click[Form_frm_Site_Impact])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          rctNoSpecies_Click
' Description:  Handles rectangular box click actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 9, 2016 - for NCPN tools
' Revisions:
'   BLC, 2/9/2016  - initial version
' ----------------------------------
Private Sub rctNoSpecies_Click()
On Error GoTo Err_Handler

    cbxNoSpecies_Click

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - rctNoSpecies_Click[Form_frm_Site_Impact])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          rctNoDisturbance_Click
' Description:  Handles rectangular box click actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 9, 2016 - for NCPN tools
' Revisions:
'   BLC, 2/9/2016  - initial version
' ----------------------------------
Private Sub rctNoDisturbance_Click()
On Error GoTo Err_Handler

    cbxNoDisturbance_Click

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - rctNoDisturbance_Click[Form_frm_Site_Impact])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    On Error GoTo Err_Handler
    If IsNull(Me!Event_ID) Then
      MsgBox "You must enter event information first."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
      GoTo Exit_Procedure
    End If

    ' Default to Events Start Date if visit date is null
    If IsNull(Me.Parent!Start_Date) Then
      MsgBox "Missing site visit date."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
      GoTo Exit_Procedure
    ElseIf IsNull(Me!Visit_Date) Then
      Me!Visit_Date = Me.Parent!Start_Date
    End If
    
    ' Create the GUID primary key value
    If IsNull(Me!Impact_ID) Then
        If GetDataType("tbl_Site_Impact", "Impact_ID") = dbText Then
            Me.Impact_ID = fxnGUIDGen
        End If
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

Private Sub ButtonSiteSketch_Click()
On Error GoTo Err_ButtonSiteSketch_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Site_Sketch"
    
    stLinkCriteria = "[Impact_ID]=" & "'" & Me![Impact_ID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_ButtonSiteSketch_Click:
    Exit Sub

Err_ButtonSiteSketch_Click:
    MsgBox Err.Description
    Resume Exit_ButtonSiteSketch_Click
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Percent_Dead_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Percent_Top_Kill_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Recorder_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Visit_Date_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub
