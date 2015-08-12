﻿Version =20
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
    Width =12240
    DatasheetFontHeight =9
    ItemSuffix =35
    Left =255
    Top =75
    Right =13920
    Bottom =7890
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
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
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
                    Caption ="DBH/DRC (cm)"
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
                    Left =10680
                    Top =60
                    Height =300
                    Name ="ButtonMaster"
                    Caption ="Master Lookup"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10680
                    Top =420
                    Height =300
                    TabIndex =1
                    Name ="ButtonUnknown"
                    Caption ="Unknown Species"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5220
                    Top =840
                    Width =960
                    Height =240
                    FontWeight =700
                    Name ="Label34"
                    Caption ="Indicator"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
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
                    RowSource ="SELECT qryU_Top_Canopy.Master_PLANT_Code, qryU_Top_Canopy.LU_Code, qryU_Top_Cano"
                        "py.Utah_Species FROM qryU_Top_Canopy WHERE (((qryU_Top_Canopy.Utah_Species) Is N"
                        "ot Null)) ORDER BY qryU_Top_Canopy.LU_Code; "
                    ColumnWidths ="0;2160;4320"

                End
                Begin ComboBox
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

Private Sub DBH_AfterUpdate()

  If Not IsNull(Me!DBH) And IsNull(Me!DType) Then
    Me!DType = "dbh"   ' Default type indicator to dbh
  ElseIf IsNull(Me!DBH) Then
    Me!DType = Null    ' If they null out dbh, then type gets nulled
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

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

Private Sub Form_Load()
    Dim Veg_Type As Variant
    
    Veg_Type = DLookup("[Vegetation_Type]", "tbl_Locations", "[Location_ID] = '" & Me.Parent!Location_ID & "'")
    If Not IsNull(Veg_Type) And Veg_Type = "oak scrub" Then
      Me!Crown_Class.visible = False
      Me!Crown_Class_Label.visible = False
    End If
End Sub
Private Sub ButtonDelete_Click()
On Error GoTo Err_ButtonDelete_Click

  Dim intReply As Integer
  
  intReply = MsgBox("Are you sure you want to delete this record?", vbYesNo, "Delete Record")
    If intReply = vbYes Then
      DoCmd.SetWarnings False
      DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
      DoCmd.DoMenuItem acFormBar, acEditMenu, 6, , acMenuVer70
      DoCmd.SetWarnings True
      Me.Requery
    End If

Exit_ButtonDelete_Click:
    Exit Sub

Err_ButtonDelete_Click:
    MsgBox Err.Description
    Resume Exit_ButtonDelete_Click
    
End Sub
