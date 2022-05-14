Version =21
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    TabularFamily =124
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14160
    DatasheetFontHeight =9
    ItemSuffix =35
    Left =405
    Top =2205
    Right =14355
    Bottom =5700
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x12768bfd3188e340
    End
    RecordSource ="qry_OT_Census"
    Caption ="frm_OT_Census"
    BeforeInsert ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
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
            Height =1140
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =180
                    Top =840
                    Width =660
                    Height =240
                    FontWeight =700
                    Name ="Quad_Label"
                    Caption ="Quad"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1080
                    Top =840
                    Width =600
                    Height =240
                    FontWeight =700
                    Name ="Tag_No_Label"
                    Caption ="Tag #"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1860
                    Top =840
                    Width =720
                    Height =240
                    FontWeight =700
                    Name ="Species_Label"
                    Caption ="Species"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =4320
                    Top =660
                    Width =840
                    Height =420
                    FontWeight =700
                    Name ="DBH_Label"
                    Caption ="Diameter (cm)"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =6360
                    Top =660
                    Width =1020
                    Height =420
                    FontWeight =700
                    Name ="Crown_Health_Label"
                    Caption ="Crown Health"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =7680
                    Top =660
                    Width =960
                    Height =420
                    FontWeight =700
                    Name ="Crown_Class_Label"
                    Caption ="Crown Class"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =8820
                    Top =840
                    Width =1560
                    Height =240
                    FontWeight =700
                    Name ="Notes_Label"
                    Caption ="Notes/Conditions"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =3000
                    Top =60
                    Width =3600
                    Height =420
                    FontSize =14
                    FontWeight =700
                    Name ="Label18"
                    Caption ="Overstory Census"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10440
                    Top =300
                    Height =300
                    Name ="ButtonMaster"
                    Caption ="Master Lookup"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =10440
                    LayoutCachedTop =300
                    LayoutCachedWidth =11880
                    LayoutCachedHeight =600
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10440
                    Top =660
                    Height =300
                    TabIndex =1
                    Name ="ButtonUnknown"
                    Caption ="Unknown Species"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =10440
                    LayoutCachedTop =660
                    LayoutCachedWidth =11880
                    LayoutCachedHeight =960
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5205
                    Top =840
                    Width =975
                    Height =240
                    FontWeight =700
                    Name ="lblDBHDRC"
                    Caption ="DBH/DRC"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =5205
                    LayoutCachedTop =840
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =1080
                End
            End
        End
        Begin Section
            Height =540
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
                    Name ="Census_ID"
                    ControlSource ="Census_ID"
                    StatusBarText ="Unique record identifier - primary key"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextAlign =2
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
                    FontName ="Tahoma"

                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =240
                    Top =60
                    Width =600
                    Height =255
                    ColumnWidth =600
                    TabIndex =2
                    Name ="Quad"
                    ControlSource ="Quad"
                    StatusBarText ="Quadrat number"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1080
                    Top =60
                    Width =600
                    Height =255
                    ColumnWidth =600
                    TabIndex =3
                    Name ="Tag_No"
                    ControlSource ="Tag_No"
                    StatusBarText ="Tag number"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4500
                    Top =60
                    Width =480
                    Height =255
                    ColumnWidth =2310
                    TabIndex =5
                    Name ="DBH"
                    ControlSource ="DBH"
                    StatusBarText ="Diameter at breast height in centimeters"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8820
                    Top =60
                    Width =2340
                    Height =450
                    ColumnWidth =3000
                    TabIndex =9
                    Name ="Notes"
                    ControlSource ="Notes"
                    StatusBarText ="Notes about any significant damage to a living tree"
                    FontName ="Tahoma"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =6480
                    Left =1860
                    Top =60
                    Width =2304
                    TabIndex =4
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="Species"
                    ControlSource ="Species"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT   q.Master_PLANT_Code,  q.LU_Code,  q.Utah_Species, q.Lifeform  "
                        " FROM qryU_Top_Canopy  q WHERE (((q.Utah_Species) Is Not Null)  AND ((q.[Lifefor"
                        "m])='Tree')) OR (q.[LU_Code] = 'JUNIPERUS')  OR (q.[LU_Code] = 'QUEGAM')    ORDE"
                        "R BY q.LU_Code      UNION    (SELECT  DISTINCT u.Unknown_Code,  u.Unknown_Code, "
                        "    u.Plant_Type+ \" - \" + u.Plant_Description,  u.Plant_Type AS Lifeform    FR"
                        "OM tbl_Unknown_Species u WHERE u.Plant_Type  IN ('Tree','Other') OR u.Plant_Type"
                        " IS NULL   ORDER BY u.Unknown_Code);"
                    ColumnWidths ="0;2160;4320"
                    OnGotFocus ="[Event Procedure]"
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2016
                    Left =6360
                    Top =60
                    Width =1080
                    TabIndex =7
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"3\";\"2\""
                    Name ="Crown_Health"
                    ControlSource ="Crown_Health"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Crown_Health_Class.Crown_Health_Class, tlu_Crown_Health_Class.Class_D"
                        "escription FROM tlu_Crown_Health_Class; "
                    ColumnWidths ="288;1728"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =1950
                    Left =7620
                    Top =60
                    Width =1020
                    TabIndex =8
                    ColumnInfo ="\"\";\"\";\"10\";\"50\""
                    Name ="Crown_Class"
                    ControlSource ="Crown_Class"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Crown_Class.Crown_Class FROM tlu_Crown_Class; "
                    ColumnWidths ="1950"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =11340
                    Top =120
                    Width =705
                    Height =300
                    TabIndex =10
                    ForeColor =255
                    Name ="ButtonDelete"
                    Caption ="Delete"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =495
                    Left =5340
                    Top =60
                    Width =780
                    TabIndex =6
                    Name ="DType"
                    ControlSource ="DType"
                    RowSourceType ="Value List"
                    RowSource ="\"dbh\";\"DRC\""
                    ColumnWidths ="495"

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
' MODULE:       Form_fsub_OT_Census
' Level:        Form module
' Version:      1.03
' Description:  data functions & procedures specific to overstory census monitoring
'
' Source/date:  Bonnie Campbell, 2/2/2016
' Revisions:    RDB - unknown  - 1.00 - initial version
'               BLC - 2/2/2016 - 1.01 - added documentation, set checkbox for no species found
'               BLC - 3/8/2016 - 1.02 - added Species_GotFocus() to refresh species lists to include
'                                       new unknowns
'               BLC - 4/13/2016 - 1.03 - added refresh for underlying subforms for conditional formatting
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
'   JRB, 6/x/2006  - initial version
'   RDB, unknown   - ?
'   BLC, 2/2/2016  - added documentation
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler

    Dim Veg_Type As Variant
        
    Veg_Type = DLookup("[Vegetation_Type]", "tbl_Locations", "[Location_ID] = '" & Me.Parent!Location_ID & "'")
    If Not IsNull(Veg_Type) And Veg_Type = "oak scrub" Then
      Me!Crown_Class.Visible = False
      Me!Crown_Class_Label.Visible = False
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[Form_fsub_OT_Census])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Species_GotFocus
' Description:  Handles species actions when control has focus
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, March 8, 2016 - for NCPN tools
' Adapted:
' Revisions:
'   BLC, 3/8/2016  - initial version
' ---------------------------------
Private Sub Species_GotFocus()
On Error GoTo Err_Handler

    'update the data to ensure new unknowns are added
    Me.ActiveControl.Requery

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Species_GotFocus[Form_fsub_OT_Census])"
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
' Source/date:
' Adapted:      Bonnie Campbell, February 2, 2016 - for NCPN tools
' Revisions:
'   BLC, 2/2/2016  - initial version
'   BLC, 2/2/2016  - initial version
' ---------------------------------
Private Sub Form_BeforeInsert(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Make sure there is an events record
    If IsNull(Me.Parent!Start_Date) Then
      MsgBox "Missing site visit date."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
      Exit Sub
    End If

    ' Create the GUID primary key value
    If IsNull(Me!Census_ID) Then
        If GetDataType("tbl_OT_Census", "Census_ID") = dbText Then
            Me.Census_ID = fxnGUIDGen
        End If
    End If

    '-----------------------------------
    ' update the NoDataCollected info
    '-----------------------------------
    Dim noData As Scripting.Dictionary
    
    'remove the no data collected record
    Set noData = SetNoDataCollected(Me.Parent.Form.Controls("Event_ID"), "E", "OverstoryTree-Census", 0)
        
    'update checkbox/rectangle
    Me.Parent.Form.Controls("cbxNoCensus") = 0
    Me.Parent.Form.Controls("cbxNoCensus").Enabled = False
    Me.Parent.Form.Controls("rctNoCensus").Visible = False

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeInsert[Form_fsub_OT_Tree_Saplings])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub ButtonMaster_Click()
On Error GoTo Err_ButtonMaster_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Master_Species"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_ButtonMaster_Click:
    Exit Sub

Err_ButtonMaster_Click:
    MsgBox Err.Description
    Resume Exit_ButtonMaster_Click
    
End Sub

Private Sub ButtonUnknown_Click()
On Error GoTo Err_ButtonUnknown_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_List_Unknown"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , acDialog
    Me.Refresh

Exit_ButtonUnknown_Click:
    Exit Sub

Err_ButtonUnknown_Click:
    MsgBox Err.Description
    Resume Exit_ButtonUnknown_Click
    
End Sub

Private Sub Crown_Class_AfterUpdate()
  If Not IsNull(Me!Crown_Class) Then
    If (Me!Crown_Class <> "none") And (Me!Crown_Health > 4) Then
      MsgBox "Crown Class must be 'none' on dead trees."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
    End If
  End If
End Sub

Private Sub Crown_Health_AfterUpdate()
  Dim Veg_Type As Variant
  Dim DBH_Limit As Integer
  
  If Not IsNull(Me!Crown_Health) Then
    If IsNull(Me!DBH) And (Me!Crown_Health <> 6) Then
      MsgBox "DBH cannot be null unless Crown Health is dead, fallen."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
      Exit Sub
    End If
    DBH_Limit = 15  ' Set default DBH limit
    Veg_Type = DLookup("[Vegetation_Type]", "tbl_Locations", "[Location_ID] = '" & Me.Parent!Location_ID & "'")
    If Not IsNull(Veg_Type) And (Veg_Type = "oak scrub") Then
      If (Me!Crown_Health <> 0) And (Me!Crown_Health <> 5) And (Me!Crown_Health <> 6) Then
        MsgBox "Crown Health must be 0, 5, or 6 on oak plots"
        DoCmd.CancelEvent
        SendKeys "{ESC}"
        Exit Sub
      End If
      If Not IsNull(Me!Species) And Me!Species = "QUGA" Then
        DBH_Limit = 10
      End If
    End If
    If Not IsNull(Veg_Type) And (Veg_Type = "woodland") Then
      If (Me!Crown_Health <> 0) And (Me!Crown_Health <> 5) And (Me!Crown_Health <> 6) Then
        MsgBox "Crown Health must be 0, 5, or 6 on woodland plots"
        DoCmd.CancelEvent
        SendKeys "{ESC}"
        Exit Sub
      End If
    End If
    If Me!Crown_Health > 4 Then
      Me!Crown_Class = "None"  ' Force crown class to none on dead trees
      Exit Sub  ' Smaller DBH is ok on dead trees
    ElseIf Me!DBH <= DBH_Limit Then
      MsgBox "DBH value is not within domain limits"
      DoCmd.CancelEvent
      SendKeys "{ESC}"
    End If
  End If
End Sub

Private Sub DBH_BeforeUpdate(Cancel As Integer)
  Dim Veg_Type As Variant
  Dim DBH_Limit As Integer
  
  If Not IsNull(Me!DBH) And Not IsNull(Me!Crown_Health) Then
    If Me!Crown_Health = 5 Then  ' Smaller DBH is ok on standing dead trees
      Exit Sub
    End If
    DBH_Limit = 15  ' Set default DBH limit
    Veg_Type = DLookup("[Vegetation_Type]", "tbl_Locations", "[Location_ID] = '" & Me.Parent!Location_ID & "'")
    If Not IsNull(Veg_Type) And (Veg_Type = "oak scrub") Then
      If Not IsNull(Me!Species) And Me!Species = "QUGA" Then
        DBH_Limit = 10
      End If
    End If
    If Me!DBH <= DBH_Limit Then
      MsgBox "Value is not within domain limits"
      DoCmd.CancelEvent
      SendKeys "{ESC}"
    End If
  ElseIf Me!Crown_Health <> 6 Then
    MsgBox "DBH cannot be null unless Crown Health is dead,fallen."
    DoCmd.CancelEvent
    SendKeys "{ESC}"
  End If
  
End Sub

Private Sub DBH_AfterUpdate()

  If Not IsNull(Me!DBH) And IsNull(Me!DType) Then
    Me!DType = "dbh"   ' Default type indicator to dbh
  ElseIf IsNull(Me!DBH) Then
    Me!DType = Null    ' If they null out dbh, then type gets nulled
  End If

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
'   BLC, 4/13/2016 - added requery of related subform to clear/set conditional formatting on change
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
    If Me.RecordsetClone.RecordCount = 0 Then
    
        Dim noData As Scripting.Dictionary
        
        'remove the no data collected record
        Set noData = SetNoDataCollected(Me.Parent.Form.Controls("Event_ID"), "E", "OverstoryTree-Census", 1)
    
        'update checkbox/rectangle
        Me.Parent.Form.Controls("cbxNoCensus") = 1
        Me.Parent.Form.Controls("cbxNoCensus").Enabled = True
        Me.Parent.Form.Controls("rctNoCensus").Visible = True
        
        'refresh the subform to clear conditional formatting
        Me.Requery
        
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ButtonDelete_Click[Form_fsub_LP_Belt_Shrub])"
    End Select
    Resume Exit_Handler
End Sub
