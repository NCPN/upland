Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    FilterOn = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =14568
    DatasheetFontHeight =10
    ItemSuffix =197
    Left =2895
    Top =3000
    Right =17460
    Bottom =15360
    DatasheetGridlinesColor =12632256
    Filter ="[Location_ID]='{512272A4-2013-42F2-886E-C6EE417DF50D}' AND [Event_ID]='201508101"
        "55312-533424019.813538'"
    RecSrcDt = Begin
        0xaf0254819108e340
    End
    RecordSource ="qfrm_DataEntry"
    Caption =" Data Entry Form - Filter by sampling event - Filter by sampling event"
    OnCurrent ="[Event Procedure]"
    BeforeInsert ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
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
    AllowLayoutView =0
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
        Begin Section
            CanGrow = NotDefault
            Height =14475
            BackColor =12574431
            Name ="Detail"
            Begin
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =6180
                    Top =660
                    Width =1080
                    Height =479
                    FontSize =9
                    FontWeight =700
                    TabIndex =12
                    Name ="cmdClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Close the data entry form"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    ColumnCount =5
                    ListWidth =7488
                    Left =7980
                    Top =900
                    Width =768
                    ColumnWidth =1440
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"3\";\"2\""
                    Name ="cboLocation_ID"
                    ControlSource ="Location_ID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Locations.Location_ID, tbl_Locations.Plot_ID, tbl_Locations.Unit_Code"
                        ", tbl_Locations.E_Coord, tbl_Locations.N_Coord FROM tbl_Locations ORDER BY tbl_L"
                        "ocations.Unit_Code, tbl_Locations.Plot_ID, tbl_Locations.E_Coord, tbl_Locations."
                        "N_Coord; "
                    ColumnWidths ="0;3456;1152;1440;1440"
                    StatusBarText ="Unique identifier for each sample location"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =8760
                            Top =840
                            Width =780
                            Height =255
                            FontWeight =700
                            Name ="labLocation_ID"
                            Caption ="Site"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1560
                    Top =900
                    Width =1080
                    TabIndex =2
                    Name ="txtStart_Date"
                    ControlSource ="Start_Date"
                    Format ="Short Date"
                    StatusBarText ="M. Starting date for the event (Start_Date)"
                    AfterUpdate ="[Event Procedure]"
                    InputMask ="99/99/0000;0;_"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =120
                            Top =900
                            Width =960
                            Height =240
                            FontWeight =700
                            Name ="Label55"
                            Caption ="Start Date"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1560
                    Top =480
                    Width =3480
                    TabIndex =4
                    Name ="txtXY"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =120
                            Top =480
                            Width =1365
                            Height =240
                            FontWeight =700
                            Name ="Label58"
                            Caption ="UTM E/N"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2580
                    Top =120
                    Width =840
                    TabIndex =5
                    Name ="txtUnit_Code"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =1560
                            Top =120
                            Width =990
                            Height =240
                            FontWeight =700
                            Name ="Label60"
                            Caption ="Park"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =60
                    Top =60
                    Width =5940
                    Name ="Box62"
                End
                Begin Tab
                    MultiRow = NotDefault
                    OverlapFlags =85
                    Left =45
                    Top =1260
                    Width =14523
                    Height =13215
                    TabIndex =6
                    Name ="pgTabs"
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =45
                    LayoutCachedTop =1260
                    LayoutCachedWidth =14568
                    LayoutCachedHeight =14475
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =180
                            Top =1665
                            Width =14250
                            Height =12675
                            Name ="pgPhotos"
                            Caption ="Photos"
                            LayoutCachedLeft =180
                            LayoutCachedTop =1665
                            LayoutCachedWidth =14430
                            LayoutCachedHeight =14340
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =360
                                    Top =1680
                                    Width =13680
                                    Height =6960
                                    Name ="frm_sub_Photo_Entry"
                                    SourceObject ="Form.frm_sub_Photo_Entry"
                                    LinkChildFields ="Event_ID"
                                    LinkMasterFields ="Event_ID"

                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =180
                            Top =1665
                            Width =14250
                            Height =12675
                            Name ="Point Intercept"
                            EventProcPrefix ="Point_Intercept"
                            Caption ="Point Intercept"
                            LayoutCachedLeft =180
                            LayoutCachedTop =1665
                            LayoutCachedWidth =14430
                            LayoutCachedHeight =14340
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =180
                                    Top =1800
                                    Width =14100
                                    Height =9915
                                    Name ="frm_LP_Transect"
                                    SourceObject ="Form.frm_LP_Transect"
                                    LinkChildFields ="Event_ID"
                                    LinkMasterFields ="Event_ID"

                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =180
                            Top =1665
                            Width =14250
                            Height =12675
                            Name ="pgBeltShrub"
                            Caption ="1-m Belt"
                            LayoutCachedLeft =180
                            LayoutCachedTop =1665
                            LayoutCachedWidth =14430
                            LayoutCachedHeight =14340
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =720
                                    Top =1680
                                    Width =12990
                                    Height =8655
                                    Name ="frm_LP_Belt_Transect"
                                    SourceObject ="Form.frm_LP_Belt_Transect"
                                    LinkChildFields ="Event_ID"
                                    LinkMasterFields ="Event_ID"

                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =180
                            Top =1665
                            Width =14250
                            Height =12675
                            Name ="pgGaps"
                            Caption ="Gap Intercepts"
                            LayoutCachedLeft =180
                            LayoutCachedTop =1665
                            LayoutCachedWidth =14430
                            LayoutCachedHeight =14340
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    SpecialEffect =0
                                    Left =180
                                    Top =1680
                                    Width =13770
                                    Height =10575
                                    Name ="frm_Canopy_Transect"
                                    SourceObject ="Form.frm_Canopy_Transect"
                                    LinkChildFields ="Event_ID"
                                    LinkMasterFields ="Event_ID"

                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =180
                            Top =1665
                            Width =14250
                            Height =12675
                            Name ="pgSS"
                            Caption ="Soil Stability"
                            LayoutCachedLeft =180
                            LayoutCachedTop =1665
                            LayoutCachedWidth =14430
                            LayoutCachedHeight =14340
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =180
                                    Top =1740
                                    Width =13860
                                    Height =7215
                                    Name ="fsub_Soil_Stability"
                                    SourceObject ="Form.fsub_Soil_Stability"
                                    LinkChildFields ="Event_ID"
                                    LinkMasterFields ="Event_ID"

                                End
                            End
                        End
                        Begin Page
                            Visible = NotDefault
                            OverlapFlags =215
                            Left =180
                            Top =1665
                            Width =14250
                            Height =12675
                            Name ="pgSLIntercept"
                            Caption ="SL Intercept"
                            LayoutCachedLeft =180
                            LayoutCachedTop =1665
                            LayoutCachedWidth =14430
                            LayoutCachedHeight =14340
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =1440
                                    Top =1800
                                    Width =11550
                                    Height =8655
                                    Name ="frm_SL_Transect"
                                    SourceObject ="Form.frm_SL_Transect"
                                    LinkChildFields ="Event_ID"
                                    LinkMasterFields ="Event_ID"

                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =180
                            Top =1665
                            Width =14250
                            Height =12675
                            Name ="pgOT"
                            Caption ="Overstory Trees"
                            LayoutCachedLeft =180
                            LayoutCachedTop =1665
                            LayoutCachedWidth =14430
                            LayoutCachedHeight =14340
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =1440
                                    Top =2040
                                    Width =10650
                                    Height =3780
                                    Name ="fsub_OT_Tree_Saplings"
                                    SourceObject ="Form.fsub_OT_Tree_Saplings"
                                    LinkChildFields ="Event_ID"
                                    LinkMasterFields ="Event_ID"

                                End
                                Begin Subform
                                    OverlapFlags =247
                                    Left =720
                                    Top =6300
                                    Width =12810
                                    Height =3780
                                    TabIndex =1
                                    Name ="fsub_OT_Census"
                                    SourceObject ="Form.fsub_OT_Census"
                                    LinkChildFields ="Event_ID"
                                    LinkMasterFields ="Event_ID"

                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    Left =3780
                                    Top =1680
                                    Width =1020
                                    TabIndex =2
                                    Name ="Sapling_Date"
                                    ControlSource ="Sapling_Date"
                                    Format ="Short Date"
                                    FontName ="Tahoma"
                                    InputMask ="99/99/0000;0;_"

                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =3240
                                            Top =1680
                                            Width =540
                                            Height =240
                                            FontWeight =700
                                            Name ="Label176"
                                            Caption ="Date"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    ListWidth =2880
                                    Left =6000
                                    Top =1680
                                    TabIndex =3
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                                    Name ="Sapling_Observer"
                                    ControlSource ="Sapling_Observer"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tlu_Contacts.Contact_ID, tlu_Contacts.Last_Name, tlu_Contacts.First_Name "
                                        "FROM tlu_Contacts WHERE (((tlu_Contacts.Active)=1)); "
                                    ColumnWidths ="0;1440;1440"
                                    FontName ="Tahoma"

                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =5100
                                            Top =1680
                                            Width =840
                                            Height =245
                                            FontWeight =700
                                            Name ="Label178"
                                            Caption ="Observer"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    ListWidth =1980
                                    Left =8640
                                    Top =1680
                                    TabIndex =4
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                                    Name ="Sapling_Recorder"
                                    ControlSource ="Sapling_Recorder"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tlu_Contacts.Contact_ID, tlu_Contacts.Last_Name, tlu_Contacts.First_Name "
                                        "FROM tlu_Contacts WHERE (((tlu_Contacts.Active)=1)); "
                                    ColumnWidths ="0;990;990"
                                    FontName ="Tahoma"

                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =7740
                                            Top =1680
                                            Width =840
                                            Height =245
                                            FontWeight =700
                                            Name ="Label182"
                                            Caption ="Recorder"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    Left =3900
                                    Top =5940
                                    Width =1019
                                    TabIndex =5
                                    Name ="Census_Date"
                                    ControlSource ="Census_Date"
                                    Format ="Short Date"
                                    FontName ="Tahoma"
                                    InputMask ="99/99/0000;0;_"

                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =3360
                                            Top =5940
                                            Width =540
                                            Height =240
                                            FontWeight =700
                                            Name ="Label190"
                                            Caption ="Date"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    ListWidth =2880
                                    Left =6180
                                    Top =5940
                                    TabIndex =6
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                                    Name ="Census_Observer"
                                    ControlSource ="Census_Observer"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tlu_Contacts.Contact_ID, tlu_Contacts.Last_Name, tlu_Contacts.First_Name "
                                        "FROM tlu_Contacts WHERE (((tlu_Contacts.Active)=1)); "
                                    ColumnWidths ="0;1440;1440"
                                    FontName ="Tahoma"

                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =5220
                                            Top =5940
                                            Width =885
                                            Height =245
                                            FontWeight =700
                                            Name ="Label192"
                                            Caption ="Observer"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    ListWidth =2880
                                    Left =8820
                                    Top =5940
                                    TabIndex =7
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                                    Name ="Census_Recorder"
                                    ControlSource ="Census_Recorder"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tlu_Contacts.Contact_ID, tlu_Contacts.Last_Name, tlu_Contacts.First_Name "
                                        "FROM tlu_Contacts WHERE (((tlu_Contacts.Active)=1)); "
                                    ColumnWidths ="0;1440;1440"
                                    FontName ="Tahoma"

                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =7920
                                            Top =5940
                                            Width =839
                                            Height =245
                                            FontWeight =700
                                            Name ="Label194"
                                            Caption ="Recorder"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =180
                            Top =1665
                            Width =14250
                            Height =12675
                            Name ="pgFuels"
                            Caption ="Fuels"
                            LayoutCachedLeft =180
                            LayoutCachedTop =1665
                            LayoutCachedWidth =14430
                            LayoutCachedHeight =14340
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin TextBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =4320
                                    Top =1740
                                    Width =960
                                    TabIndex =2
                                    Name ="FuelsDate"
                                    ControlSource ="Fuels_Date"
                                    Format ="Short Date"
                                    FontName ="Tahoma"
                                    InputMask ="99/99/0000;0;_"

                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =2700
                                            Top =1740
                                            Width =1620
                                            Height =240
                                            FontWeight =700
                                            Name ="Label123"
                                            Caption ="Observation Date"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    ListWidth =1980
                                    Left =6420
                                    Top =1740
                                    Width =1380
                                    TabIndex =3
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                                    Name ="FuelsObserver"
                                    ControlSource ="Fuels_Observer"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tlu_Contacts.Contact_ID, tlu_Contacts.Last_Name, tlu_Contacts.First_Name "
                                        "FROM tlu_Contacts WHERE (((tlu_Contacts.Active)=1)); "
                                    ColumnWidths ="0;990;990"
                                    FontName ="Tahoma"

                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =5580
                                            Top =1740
                                            Width =840
                                            Height =245
                                            FontWeight =700
                                            Name ="Label125"
                                            Caption ="Observer"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    ListWidth =1980
                                    Left =8940
                                    Top =1740
                                    Width =1380
                                    TabIndex =4
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                                    Name ="FuelsRecorder"
                                    ControlSource ="Fuels_Recorder"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tlu_Contacts.Contact_ID, tlu_Contacts.Last_Name, tlu_Contacts.First_Name "
                                        "FROM tlu_Contacts WHERE (((tlu_Contacts.Active)=1)); "
                                    ColumnWidths ="0;990;990"
                                    FontName ="Tahoma"

                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =8100
                                            Top =1740
                                            Width =840
                                            Height =245
                                            FontWeight =700
                                            Name ="Recorder_Label"
                                            Caption ="Recorder"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    Left =360
                                    Top =2160
                                    Width =13140
                                    Height =4680
                                    Name ="fsub_Fuels_LD"
                                    SourceObject ="Form.fsub_Fuels_LD"
                                    LinkChildFields ="Event_ID"
                                    LinkMasterFields ="Event_ID"

                                End
                                Begin Subform
                                    OverlapFlags =247
                                    Left =4020
                                    Top =7260
                                    Width =5940
                                    Height =3960
                                    TabIndex =1
                                    Name ="fsub_Fuels_1000"
                                    SourceObject ="Form.fsub_Fuels_1000"
                                    LinkChildFields ="Event_ID"
                                    LinkMasterFields ="Event_ID"

                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =180
                            Top =1665
                            Width =14250
                            Height =12675
                            Name ="pgImpact"
                            Caption ="Site Impact"
                            LayoutCachedLeft =180
                            LayoutCachedTop =1665
                            LayoutCachedWidth =14430
                            LayoutCachedHeight =14340
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =180
                                    Top =1680
                                    Width =12630
                                    Height =9840
                                    Name ="frm_Site_Impact"
                                    SourceObject ="Form.frm_Site_Impact"
                                    LinkChildFields ="Event_ID"
                                    LinkMasterFields ="Event_ID"

                                End
                            End
                        End
                        Begin Page
                            Visible = NotDefault
                            OverlapFlags =247
                            Left =180
                            Top =1665
                            Width =14250
                            Height =12675
                            Name ="pgCoords_and_loc_details"
                            Caption ="Presence Cover Density"
                            LayoutCachedLeft =180
                            LayoutCachedTop =1665
                            LayoutCachedWidth =14430
                            LayoutCachedHeight =14340
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =180
                                    Top =1665
                                    Width =13710
                                    Height =12015
                                    Name ="frm_Quadrat_Transect"
                                    SourceObject ="Form.frm_Quadrat_Transect"
                                    LinkChildFields ="Event_ID"
                                    LinkMasterFields ="Event_ID"

                                End
                            End
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =60
                    Top =780
                    Width =5940
                    Height =480
                    Name ="Box65"
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =7560
                    Top =540
                    Width =300
                    TabIndex =7
                    Name ="txtLocation_ID"
                    ControlSource ="Location_ID"
                    StatusBarText ="M. Link to tbl_Locations (Loc_ID)"

                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =7560
                    Top =900
                    TabIndex =8
                    Name ="Site_Selection"
                    ControlSource ="Site_Selection"
                    StatusBarText ="Site accepted or rejected"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =8100
                    Top =540
                    Width =540
                    TabIndex =9
                    Name ="version_key_number"
                    ControlSource ="version_key_number"
                    StatusBarText ="Master protocol version key"

                End
                Begin Label
                    OverlapFlags =247
                    Left =120
                    Top =120
                    Width =480
                    Height =240
                    FontWeight =700
                    Name ="Label91"
                    Caption ="Site"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    BackStyle =0
                    IMESentenceMode =3
                    Left =660
                    Top =120
                    Width =480
                    TabIndex =10
                    Name ="SiteDisplay"

                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =8880
                    Top =180
                    Width =4800
                    Height =899
                    TabIndex =1
                    Name ="Comments"
                    ControlSource ="Comments"
                    StatusBarText ="Plot revisit comments."

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7440
                            Top =180
                            Width =1380
                            Height =240
                            FontWeight =700
                            Name ="Label94"
                            Caption ="Visit Comments:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6180
                    Top =120
                    Width =1080
                    Height =480
                    TabIndex =11
                    Name ="ButtonCoord"
                    Caption ="Change Plot Coordinates"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =1650
                    Left =3780
                    Top =900
                    Width =2100
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Observer"
                    ControlSource ="Observer"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Contacts.Contact_ID, tlu_Contacts.Last_Name, tlu_Contacts.First_Name "
                        "FROM tlu_Contacts WHERE (((tlu_Contacts.Active)=1)); "
                    ColumnWidths ="0;810;839"

                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =2940
                            Top =900
                            Width =840
                            Height =245
                            FontWeight =700
                            Name ="Observer_Label"
                            Caption ="Observer"
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =247
                    Left =7620
                    Top =660
                    Width =1020
                    Height =480
                    TabIndex =13
                    Name ="ButtonComments"
                    Caption ="Add/Edit Comments"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =5340
                    Top =180
                    Width =360
                    TabIndex =14
                    Name ="Unit_Code"
                    ControlSource ="Unit_Code"
                    StatusBarText ="Park Code."

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
' FORM NAME:    frm_Data_Entry
' Description:  Primary field data entry form
' Data source:  tbl_Locations
' Data access:  edit; allow additions off except for new records
' Pages:        none
' Functions:    none
' References:   fxnSwitchboardIsOpen, fxnGUIDGen
' Source/date:  John R. Boetsch, June 2006
' Revisions:    <name, date, desc - add lines as you go>
' =================================

Private Sub cboLocation_ID_AfterUpdate()
' Update_Loc_Info
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
  If IsNull(Me!Start_Date) Then
        ' ask user if (s)he wants to enter data or cancel and close form
        If MsgBox("Visit date is missing - do you want to enter the missing data?", vbYesNo, "Date missing") = vbNo Then
            Me.Undo
        End If
  End If
End Sub

Private Sub Form_Close()
If IsLoaded("frm_Data_Gateway") Then
    Forms("frm_Data_Gateway").Requery
End If
End Sub

Private Sub Form_Current()

Update_Loc_Info
If Not IsNull(Me!txtUnit_Code) Then
'  MsgBox DLookup("[ParkState]", "tlu_Parks", "[ParkCode] = '" & Me!txtUnit_Code & "'")
  Me!frm_Quadrat_Transect.Form!fsub_Quadrat.Form!fsub_Quadrat_Shrubs.Form!State_Code = DLookup("[ParkState]", "tlu_Parks", "[ParkCode] = '" & Me!txtUnit_Code & "'")
  Me!frm_Quadrat_Transect.Form!fsub_Quadrat.Form!fsub_Species.Form!State_Code = DLookup("[ParkState]", "tlu_Parks", "[ParkCode] = '" & Me!txtUnit_Code & "'")
  Me!SiteDisplay = cboLocation_ID.Column(1)  ' Display the site number in heading
End If
End Sub

Private Sub Form_Load()
    Dim Veg_Type As Variant

  ' Display the proper tabs
    Veg_Type = DLookup("[Vegetation_Type]", "tbl_Locations", "[Location_ID] = '" & Me!Location_ID & "'")
    If Not IsNull(Veg_Type) And (Veg_Type = "forest" Or Veg_Type = "oak scrub") Then
      Me!pgSS.visible = False
      Me!pgGaps.visible = False
    End If
    If Not IsNull(Veg_Type) And (Veg_Type = "grassland/shrubland" Or Veg_Type = "oak scrub") Then
      Me!pgFuels.visible = False
    End If
    If Not IsNull(Veg_Type) And Veg_Type = "oak scrub" Then
  '    Me!pgBeltShrub.Visible = False  1m belt tab visible for oak plots - 2/15/2011 - RD
      Me!fsub_OT_Tree_Saplings.Form.visible = False
      Me!Sapling_Date.visible = False
      Me!Sapling_Observer.visible = False
      Me!Sapling_Recorder.visible = False
    Else
      Me!pgSLIntercept.visible = False
    End If
    If Not IsNull(Veg_Type) And Veg_Type = "woodland" Then
      Me!pgGaps.visible = False
      Me!fsub_OT_Census.Form!Crown_Class.visible = False
      Me!fsub_OT_Census.Form!Crown_Class_Label.visible = False
    End If
    If Not IsNull(Me!txtStart_date) Then
      If IsNull(Me!Fuels_Date) Then
        Me!Fuels_Date = Me!txtStart_date
      End If
      If IsNull(Me!Census_Date) Then
        Me!Census_Date = Me!txtStart_date
      End If
      If IsNull(Me!Sapling_Date) Then
        Me!Sapling_Date = Me!txtStart_date
      End If
    End If
    ' Modified to hide fuels form for TICA 1 [HMT, 3/13/2015]
    ' TICA 1 is a special case of a forest plot that does not have fuels data collected.
    If (Me!Unit_Code = "TICA") And (Me!Plot_ID = 1) Then
      Me!pgFuels.visible = False
    End If
End Sub

Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Err_Handler

    Dim strCaptionSuffix As String

    ' Set the opening parameters depending on the arguments passed from the previous form
    If Me.OpenArgs = "New record" Or Me.OpenArgs = "Filter by location" Then
        strCaptionSuffix = " - " & Me.OpenArgs
    ElseIf Me.OpenArgs = "New event" Then
        strCaptionSuffix = " - " & Me.OpenArgs
    ElseIf Me.OpenArgs <> "" Then
        strCaptionSuffix = " - " & "Filter by sampling event"
    End If
    Me.Caption = Me.Caption & strCaptionSuffix
    Me!txtStart_date.SetFocus
Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)

        Dim db As DAO.Database
        Dim Versions As DAO.Recordset
        Dim strSQL As String
        
    On Error GoTo Err_Handler
    
    ' Set master version number on event record
    Set db = CurrentDb
    strSQL = "SELECT [version_key_number] FROM [tbl_master_version] ORDER BY [version_key_number] DESC"
    Set Versions = db.OpenRecordset(strSQL)
    Versions.MoveFirst
    Me![version_key_number] = Versions![version_key_number]
    Versions.Close

    ' Create the GUID primary key value
    If IsNull(Me!Event_ID) Then
        If GetDataType("tbl_Events", "Event_ID") = dbText Then
            Me.Event_ID = fxnGUIDGen
        End If
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cmdClose_Click()
    On Error GoTo Err_Handler

    DoCmd.RunCommand acCmdSaveRecord
    DoCmd.Close , , acSaveNo

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub


Public Sub Update_Loc_Info()
' Description:  Updates associated location information when Location_ID is updated
' References:   GetCriteriaString
' Source/date:  Simon Kingston, Sept. 2006
' Revisions:    <name, date, desc - add lines as you go>

Dim strXY As Variant
Dim strCriteria As String

If IsNull(Me!txtLocation_ID) Then
    Me!txtXY = Null
    Me!txtUnit_Code = Null
Else
    strCriteria = GetCriteriaString("Location_ID=", "tbl_Locations", "Location_ID", Me.name, "txtLocation_ID")
    
    strXY = "E: " & Nz(DLookup("E_Coord", "tbl_Locations", strCriteria), "")
    strXY = strXY & "  N: " & Nz(DLookup("N_Coord", "tbl_Locations", strCriteria), "")
    Me!txtXY = strXY
    Me!txtUnit_Code = DLookup("Unit_Code", "tbl_Locations", strCriteria)
End If
End Sub

Private Sub pgTabs_Change()
  Select Case Me.pgTabs.Value  ' Display a silly message so the field crews know where they are
    Case 3
      If IsNull(Me!frm_Canopy_Transect.Form!Transect) Then
        MsgBox "You are on transect 1.", 0, "Transect Verify"
      Else
        MsgBox "You are on transect " & Me!frm_Canopy_Transect.Form!Transect & ".", 0, "Transect Verify"
      End If
    Case 1
      If IsNull(Me!frm_LP_Transect.Form!Transect) Then
        MsgBox "You are on transect 1.", 0, "Transect Verify"
      Else
        MsgBox "You are on transect " & Me!frm_LP_Transect.Form!Transect & ".", 0, "Transect Verify"
      End If
    Case 2
      If IsNull(Me!frm_LP_Belt_Transect.Form!Transect) Then
        MsgBox "You are on transect 1.", 0, "Transect Verify"
      Else
        MsgBox "You are on transect " & Me!frm_LP_Belt_Transect.Form!Transect & ".", 0, "Transect Verify"
      End If
    Case 5
      If IsNull(Me!frm_SL_Transect.Form!Transect) Then
        MsgBox "You are on transect 1.", 0, "Transect Verify"
      Else
        MsgBox "You are on transect " & Me!frm_SL_Transect.Form!Transect & ".", 0, "Transect Verify"
      End If
  End Select
End Sub

Private Sub txtStart_Date_AfterUpdate()
        Dim db As DAO.Database
        Dim Events As DAO.Recordset
        Dim strSQL As String
        
    On Error GoTo Err_Handler
    
    ' Check for duplicate date
    strSQL = "SELECT Event_ID FROM tbl_Events WHERE [Location_ID] = '" & Me!cboLocation_ID & "' AND [Start_Date] = #" & Me!Start_Date & "#"
'    MsgBox strSQL
    Set db = CurrentDb
    Set Events = db.OpenRecordset(strSQL)
    If Not Events.EOF Then
      MsgBox " Duplicate visit date - update cancelled."
      Me.Undo
      Events.Close
      DoCmd.Close
      GoTo Exit_Procedure
    End If
    Events.Close
Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox Err.Description
    Resume Exit_Procedure
End Sub
Private Sub ButtonCoord_Click()
On Error GoTo Err_ButtonCoord_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Location_Modify"
    
    stLinkCriteria = "[Location_ID]=" & "'" & Me![txtLocation_ID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , acDialog
    Update_Loc_Info
Exit_ButtonCoord_Click:
    Exit Sub

Err_ButtonCoord_Click:
    MsgBox Err.Description
    Resume Exit_ButtonCoord_Click
    
End Sub
Private Sub ButtonComments_Click()
On Error GoTo Err_ButtonComments_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim RevisitComments As DAO.Recordset
    Dim db As DAO.Database
    Dim strSQL As String
    
    If IsNull(Me!SiteDisplay) Or IsNull(Me!txtUnit_Code) Or IsNull(Me!txtStart_date) Or IsNull(Me!txtStart_date) Then  ' Got to have key fields
      Exit Sub
    End If
    stDocName = "frm_Revisit_Comments"
    strSQL = "SELECT * FROM tbl_Revisit_Comments Where [Unit_Code] = '" & Me!txtUnit_Code & "' AND [Plot_ID] = " & Me!SiteDisplay & " AND [VisitDate]=" & "#" & Me![txtStart_date] & "#"
'    strSQL = "SELECT * FROM tbl_Revisit_Comments Where [Unit_Code] = '" & Me!txtUnit_Code & "' AND [Plot_ID] = " & Me!SiteDisplay & " ORDER BY [VisitDate] DESC"
    Set db = CurrentDb
    Set RevisitComments = db.OpenRecordset(strSQL)
    If RevisitComments.EOF Then
      DoCmd.OpenForm stDocName, , , , , , "New"
    Else
      RevisitComments.MoveFirst
      stLinkCriteria = "[Unit_Code] = '" & Me!txtUnit_Code & "' AND [Plot_ID] = " & Me!SiteDisplay & " AND [VisitDate]=" & "#" & Me![txtStart_date] & "#"
'      stLinkCriteria = "[Unit_Code] = '" & Me!txtUnit_Code & "' AND [Plot_ID] = " & Me!SiteDisplay & " AND [VisitDate]=" & "#" & RevisitComments!VisitDate & "#"
      DoCmd.OpenForm stDocName, , , stLinkCriteria
    End If
    RevisitComments.Close
    Set RevisitComments = Nothing
Exit_ButtonComments_Click:
    Exit Sub

Err_ButtonComments_Click:
    MsgBox Err.Description
    Resume Exit_ButtonComments_Click
    
End Sub
