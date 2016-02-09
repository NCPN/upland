Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =124
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =15360
    DatasheetFontHeight =9
    ItemSuffix =67
    Left =855
    Top =2610
    Right =16110
    Bottom =10980
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x2518a6c77056e340
    End
    RecordSource ="qry_LP_Belt_Transect"
    Caption ="frm_Canopy_Transect"
    BeforeInsert ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
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
            Height =8640
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =120
                    Top =60
                    Width =630
                    Height =180
                    ColumnWidth =2310
                    FontSize =10
                    Name ="Transect_ID"
                    ControlSource ="Transect_ID"
                    StatusBarText ="Unique record identifier - primary key"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =840
                    Top =60
                    Width =630
                    Height =180
                    ColumnWidth =2310
                    FontSize =10
                    TabIndex =1
                    Name ="Event_ID"
                    ControlSource ="Event_ID"
                    StatusBarText ="M. Link to tbl_Locations (Loc_ID)"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1080
                    Top =60
                    Width =360
                    ColumnWidth =465
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    ForeColor =255
                    Name ="Transect"
                    ControlSource ="Transect"
                    StatusBarText ="Transect number - 1, 2, or 3"

                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =60
                            Top =60
                            Width =1020
                            Height =240
                            FontSize =10
                            FontWeight =700
                            Name ="Transect_Label"
                            Caption ="Transect"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =3360
                    Top =60
                    Width =960
                    ColumnWidth =1035
                    TabIndex =3
                    Name ="Visit_Date"
                    ControlSource ="Visit_Date"
                    Format ="Short Date"
                    StatusBarText ="Date of visit."
                    InputMask ="99/99/0000;0;_"
                    OnKeyDown ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2460
                            Top =60
                            Width =900
                            Height =240
                            FontWeight =700
                            Name ="Visit_Date_Label"
                            Caption ="Visit Date"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1620
                    Top =60
                    Width =306
                    Height =306
                    TabIndex =6
                    Name ="ButtonPrevious"
                    Caption ="Command14"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddadadadad1dadadaadadadad11adadaddadadad111dadada ,
                        0xadadad1111adadaddadad11111dadadaadadad1111adadaddadadad111dadada ,
                        0xadadadad11adadaddadadadad1dadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    OnKeyDown ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Previous Record"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1980
                    Top =60
                    Width =306
                    Height =306
                    TabIndex =7
                    Name ="ButtonNext"
                    Caption ="Command15"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddadada1adadadadaadadad11adadadaddadada111adadada ,
                        0xadadad1111adadaddadada11111adadaadadad1111adadaddadada111adadada ,
                        0xadadad11adadadaddadada1adadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    OnKeyDown ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Next Record"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =1650
                    Left =5340
                    Top =60
                    Width =1620
                    TabIndex =4
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Observer"
                    ControlSource ="Observer"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Contacts.Contact_ID, tlu_Contacts.Last_Name, tlu_Contacts.First_Name "
                        "FROM tlu_Contacts WHERE (((tlu_Contacts.Active)=1)); "
                    ColumnWidths ="0;810;840"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4500
                            Top =60
                            Width =840
                            Height =245
                            FontWeight =700
                            Name ="Observer_Label"
                            Caption ="Observer"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =1545
                    Left =7980
                    Top =60
                    Width =1620
                    TabIndex =5
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Recorder"
                    ControlSource ="Recorder"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Contacts.Contact_ID, tlu_Contacts.Last_Name, tlu_Contacts.First_Name "
                        "FROM tlu_Contacts WHERE (((tlu_Contacts.Active)=1)); "
                    ColumnWidths ="0;750;795"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =7140
                            Top =60
                            Width =840
                            Height =245
                            FontWeight =700
                            Name ="Recorder_Label"
                            Caption ="Recorder"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9780
                    Top =60
                    Width =1260
                    Height =300
                    TabIndex =8
                    Name ="ButtonMaster"
                    Caption ="Master Species"
                    OnClick ="[Event Procedure]"
                    OnKeyDown ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =11160
                    Top =60
                    Height =300
                    TabIndex =9
                    Name ="ButtonUnknown"
                    Caption ="Unknown Species"
                    OnClick ="[Event Procedure]"
                    OnKeyDown ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Tab
                    OverlapFlags =85
                    Left =45
                    Top =540
                    Width =15075
                    Height =7980
                    TabIndex =10
                    Name ="TabCtl49"

                    LayoutCachedLeft =45
                    LayoutCachedTop =540
                    LayoutCachedWidth =15120
                    LayoutCachedHeight =8520
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =180
                            Top =945
                            Width =14805
                            Height =7440
                            Name ="pgBeltShrub"
                            Caption ="Density"
                            LayoutCachedLeft =180
                            LayoutCachedTop =945
                            LayoutCachedWidth =14985
                            LayoutCachedHeight =8385
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =215
                                    Left =300
                                    Top =960
                                    Width =12450
                                    Height =3840
                                    Name ="fsub_LP_Belt_Shrub"
                                    SourceObject ="Form.fsub_LP_Belt_Shrub"
                                    LinkChildFields ="Transect_ID"
                                    LinkMasterFields ="Transect_ID"

                                    LayoutCachedLeft =300
                                    LayoutCachedTop =960
                                    LayoutCachedWidth =12750
                                    LayoutCachedHeight =4800
                                End
                                Begin Subform
                                    OverlapFlags =215
                                    Left =240
                                    Top =5160
                                    Width =6354
                                    Height =2685
                                    TabIndex =1
                                    Name ="fsub_LP_Seedling"
                                    SourceObject ="Form.fsub_LP_Seedling"
                                    LinkChildFields ="Transect_ID"
                                    LinkMasterFields ="Transect_ID"

                                    LayoutCachedLeft =240
                                    LayoutCachedTop =5160
                                    LayoutCachedWidth =6594
                                    LayoutCachedHeight =7845
                                End
                                Begin Subform
                                    Visible = NotDefault
                                    OverlapFlags =215
                                    Left =6840
                                    Top =5160
                                    Width =5934
                                    Height =2685
                                    TabIndex =2
                                    Name ="fsub_LP_Exotic"
                                    SourceObject ="Form.fsub_LP_Exotic"
                                    LinkChildFields ="Transect_ID"
                                    LinkMasterFields ="Transect_ID"

                                    LayoutCachedLeft =6840
                                    LayoutCachedTop =5160
                                    LayoutCachedWidth =12774
                                    LayoutCachedHeight =7845
                                End
                            End
                        End
                        Begin Page
                            Visible = NotDefault
                            OverlapFlags =215
                            Left =180
                            Top =945
                            Width =14805
                            Height =7440
                            Name ="pgDensiometer"
                            Caption ="Spherical Densiometer"
                            LayoutCachedLeft =180
                            LayoutCachedTop =945
                            LayoutCachedWidth =14985
                            LayoutCachedHeight =8385
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =1440
                                    Top =1440
                                    Width =6330
                                    Height =2880
                                    Name ="fsub_LP_Densiometer"
                                    SourceObject ="Form.fsub_LP_Densiometer"
                                    LinkChildFields ="Transect_ID"
                                    LinkMasterFields ="Transect_ID"

                                    LayoutCachedLeft =1440
                                    LayoutCachedTop =1440
                                    LayoutCachedWidth =7770
                                    LayoutCachedHeight =4320
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            TextAlign =2
                                            Left =2520
                                            Top =1140
                                            Width =4140
                                            Height =300
                                            FontSize =10
                                            FontWeight =700
                                            Name ="fsub_LP_Densiometer Label"
                                            Caption ="Spherical Densiometer Readings"
                                            FontName ="Tahoma"
                                            EventProcPrefix ="fsub_LP_Densiometer_Label"
                                            LayoutCachedLeft =2520
                                            LayoutCachedTop =1140
                                            LayoutCachedWidth =6660
                                            LayoutCachedHeight =1440
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =180
                            Top =945
                            Width =14805
                            Height =7440
                            Name ="PgAdd"
                            Caption ="Exotic Frequency"
                            LayoutCachedLeft =180
                            LayoutCachedTop =945
                            LayoutCachedWidth =14985
                            LayoutCachedHeight =8385
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    Left =180
                                    Top =1440
                                    Width =4620
                                    Height =4530
                                    Name ="fsub_LP_Add_Species"
                                    SourceObject ="Form.fsub_LP_Add_Species"
                                    LinkChildFields ="Transect_ID"
                                    LinkMasterFields ="Transect_ID"

                                    LayoutCachedLeft =180
                                    LayoutCachedTop =1440
                                    LayoutCachedWidth =4800
                                    LayoutCachedHeight =5970
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            TextAlign =1
                                            Left =180
                                            Top =1140
                                            Width =3060
                                            Height =300
                                            FontSize =10
                                            FontWeight =700
                                            Name ="fsub_LP_Add_Species Label"
                                            Caption ="Species in 1-m Belt"
                                            FontName ="Tahoma"
                                            EventProcPrefix ="fsub_LP_Add_Species_Label"
                                            LayoutCachedLeft =180
                                            LayoutCachedTop =1140
                                            LayoutCachedWidth =3240
                                            LayoutCachedHeight =1440
                                        End
                                    End
                                End
                                Begin Subform
                                    OverlapFlags =255
                                    Left =4980
                                    Top =1440
                                    Width =7770
                                    Height =4560
                                    TabIndex =1
                                    Name ="fsub_LP_Exotic_Frequency"
                                    SourceObject ="Form.fsub_LP_Exotic_Frequency"
                                    LinkChildFields ="Transect_ID"
                                    LinkMasterFields ="Transect_ID"

                                    LayoutCachedLeft =4980
                                    LayoutCachedTop =1440
                                    LayoutCachedWidth =12750
                                    LayoutCachedHeight =6000
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =4980
                                            Top =1140
                                            Width =3720
                                            Height =300
                                            FontSize =10
                                            FontWeight =700
                                            Name ="fsub_LP_Exotic_Frequency Label"
                                            Caption ="Exotic Frequency - 1m x 1m quadrats"
                                            FontName ="Tahoma"
                                            EventProcPrefix ="fsub_LP_Exotic_Frequency_Label"
                                            LayoutCachedLeft =4980
                                            LayoutCachedTop =1140
                                            LayoutCachedWidth =8700
                                            LayoutCachedHeight =1440
                                        End
                                    End
                                End
                                Begin Subform
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    Left =4980
                                    Top =1440
                                    Width =9954
                                    Height =4560
                                    TabIndex =2
                                    Name ="fsub_LP_Exotic_Freq_Oak"
                                    SourceObject ="Form.fsub_LP_Exotic_Freq_Oak"
                                    LinkChildFields ="Transect_ID"
                                    LinkMasterFields ="Transect_ID"

                                    LayoutCachedLeft =4980
                                    LayoutCachedTop =1440
                                    LayoutCachedWidth =14934
                                    LayoutCachedHeight =6000
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =4980
                                            Top =1140
                                            Width =3720
                                            Height =300
                                            FontSize =10
                                            FontWeight =700
                                            Name ="fsub_LP_Exotic_Freq_Oak Label"
                                            Caption ="Exotic Frequency - 1m x 1m quadrats"
                                            FontName ="Tahoma"
                                            EventProcPrefix ="fsub_LP_Exotic_Freq_Oak_Label"
                                            LayoutCachedLeft =4980
                                            LayoutCachedTop =1140
                                            LayoutCachedWidth =8700
                                            LayoutCachedHeight =1440
                                        End
                                    End
                                End
                            End
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
' MODULE:       frm_LP_Belt_Transect
' Level:        Form module
' Version:      1.01
' Description:  data functions & procedures specific to LP belt transects
'
' Source/date:  John R. Boetsch, June 2006
' Adapted:      Bonnie Campbell, 2/2/2016
' Revisions:    RDB - unknown  - 1.00 - initial version
'               BLC - 2/2/2016 - 1.01 - added documentation, enabled seedlings & saplings for
'                                       oak scrub plots
' =================================

' ---------------------------------
' SUB:          Form_Open
' Description:  Handles form opening actions
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
'   BLC, 2/2/2016  - added documentation, enabled seedlings for oak scrub plots
'                    (Density tab, pgBeltShrub)
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler

  Dim Veg_Type As Variant
    Veg_Type = DLookup("[Vegetation_Type]", "tbl_Locations", "[Location_ID] = '" & Me.Parent!Location_ID & "'")
    If Not IsNull(Veg_Type) And (Veg_Type = "woodland" Or Veg_Type = "grassland/shrubland") Then
      Me!pgDensiometer.Visible = False
    End If
'    Additional species tab visible for all plots 2/15/2011 RD
'    If Not IsNull(Veg_Type) And (Veg_Type <> "forest") Then
'      Me!PgAdd.Visible = False
'    End If

'    No species richness form unless CEBR or TICA plot 1  3/9/2012 RD
    If Me.Parent!Unit_Code = "CEBR" Then
      Me!fsub_LP_Add_Species.Visible = True
    ElseIf (Me.Parent!Unit_Code = "TICA") And (Me.Parent!Plot_ID = 1) Then
      Me!fsub_LP_Add_Species.Visible = True
    Else
      Me!fsub_LP_Add_Species.Visible = False
    End If
    
    'Set up correct exotic species frequency form
    'oak scrub plots
    If Not IsNull(Veg_Type) And Veg_Type = "oak scrub" Then
        Me!fsub_LP_Exotic_Frequency.Form.Visible = False
        Me!fsub_LP_Exotic_Freq_Oak.Form.Visible = True
        'Me!fsub_LP_Add_Species.SetFocus  ' Set focus to richness tab so we can hide belt-shrub tab
        Me!Visit_Date.SetFocus
      
        '------------------------------------------------
        'enabled seedlings for oak plots - 2/2/2016 - BLC
        'but not shrubs
        '------------------------------------------------
        'Me!pgBeltShrub.visible = False
        Me!fsub_LP_Belt_Shrub.Form.Visible = False
        '------------------------------------------------
        Me!pgDensiometer.Visible = False
    Else
      Me!fsub_LP_Exotic_Frequency.Form.Visible = True
      Me!fsub_LP_Exotic_Freq_Oak.Form.Visible = False
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[Form_frm_LP_Belt_Transect])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    On Error GoTo Err_Handler

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
    If IsNull(Me!Transect_ID) Then
        If GetDataType("tbl_LP_Belt_Transect", "Transect_ID") = dbText Then
            Me.Transect_ID = fxnGUIDGen
 '           Forms!frm_Data_Entry!frm_LP_Transect.Form!fsub_Lower_Canopy.Form!Transect_ID = Me!Transect_ID
        End If
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

' ---------------------------------
' SUB:          ButtonPrevious_Click
' Description:  Handles previous button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 3, 2016 - for NCPN tools
' Revisions:
'   JRB, 6/x/2006  - initial version
'   RDB, unknown   - ?
'   BLC, 2/3/2016  - added documentation, revised to use transect number
'                    overlay vs. messagebox
' ---------------------------------
Private Sub ButtonPrevious_Click()
On Error GoTo Err_Handler
  Dim intTransect As Byte
  
  ' Disabled 3/20/09 on demand of ecologists
  ' If IsNull(Me!Recorder) And IsNull(Me!Observer) Then
  '    DoCmd.CancelEvent
  '    SendKeys "{ESC}"
  '  End If
  If Me!Transect = 1 Then
    MsgBox "Already on first transect"
  Else
    intTransect = Me!Transect
    DoCmd.GoToRecord , , acPrevious
'    DoCmd.GoToRecord , , 2
    Me!Transect = intTransect - 1
    
    '---------------------------
    'display overlay - 2/3/2016 - BLC
    '---------------------------
    'MsgBox "You are on transect " & Me!Transect & ".", 0, "Transect Verify"
    DoCmd.OpenForm "frm_Transect_Overlay", OpenArgs:=Me!Transect
    '---------------------------
  End If
  
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ButtonPrevious_Click[Form_frm_LP_Belt_Transect])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          ButtonNext_Click
' Description:  Handles next button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 3, 2016 - for NCPN tools
' Revisions:
'   JRB, 6/x/2006  - initial version
'   RDB, unknown   - ?
'   BLC, 2/3/2016  - added documentation, revised to use transect number
'                    overlay vs. messagebox
' ---------------------------------
Private Sub ButtonNext_Click()
On Error GoTo Err_Handler

' Disabled 3/20/09 on demand of ecologists.
'  If IsNull(Me!Recorder) And IsNull(Me!Observer) Then
'    MsgBox "You must record data in this transect before moving to the next."
'    GoTo Exit_ButtonNext_Click
'  End If
  Dim intTransect As Byte
    If IsNull(Me!Transect) Then
      Me!Transect = 1
    End If
  If Me!Transect = 3 Then
    MsgBox "Three transects maximum!"
  Else
    intTransect = Me!Transect
    DoCmd.GoToRecord , , acNext
    Me!Transect = intTransect + 1
    
    '---------------------------
    'display overlay - 2/3/2016 - BLC
    '---------------------------
    'MsgBox "You are on transect " & Me!Transect & ".", 0, "Transect Verify"
    DoCmd.OpenForm "frm_Transect_Overlay", OpenArgs:=Me!Transect
    '---------------------------
  End If


Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ButtonNext_Click[Form_frm_LP_Belt_Transect])"
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

    stDocName = "frm_List_Unknown"
    DoCmd.OpenForm stDocName, , , , , acDialog
    Me!fsub_LP_Belt_Shrub.Form!Species.Requery
    Me!fsub_LP_Seedling.Form!Species.Requery
'    Me!fsub_LP_Exotic.Form!Species.Requery   Page hidden 3/21/2011 RD
    Me!fsub_LP_Add_Species.Form!Species.Requery
    Me!fsub_LP_Exotic_Freq_Oak.Form!Species.Requery
    Me!fsub_LP_Exotic_Frequency.Form!Species.Requery

Exit_ButtonUnknown_Click:
    Exit Sub

Err_ButtonUnknown_Click:
    MsgBox Err.Description
    Resume Exit_ButtonUnknown_Click
    
End Sub

Private Sub ButtonMaster_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub ButtonNext_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub ButtonPrevious_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub ButtonUnknown_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Observer_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub Visit_Date_GotFocus()
    If IsNull(Me!Visit_Date) Then    ' Set default visit date
      Me!Visit_Date = Me.Parent!Start_Date
      Me.Refresh   ' Force save of transect record
    End If
End Sub

Private Sub Visit_Date_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub
