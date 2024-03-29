﻿Version =21
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
    ViewsAllowed =1
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =16296
    DatasheetFontHeight =10
    ItemSuffix =216
    Left =4470
    Top =2880
    Right =21015
    Bottom =13905
    DatasheetGridlinesColor =12632256
    Filter ="[Location_ID]='{E35D7F2C-A99C-41FE-ACEC-A1DAD79E24AC}' AND [Event_ID]='201708171"
        "65907-938545167.446136'"
    RecSrcDt = Begin
        0x171e359b4cb5e440
    End
    RecordSource ="qfrm_DataEntry"
    Caption =" Data Entry Form"
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
                    Name ="btnClose"
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
                    Top =1212
                    Width =16035
                    Height =13263
                    TabIndex =6
                    Name ="pgTabs"
                    OnChange ="[Event Procedure]"

                    LayoutCachedTop =1212
                    LayoutCachedWidth =16035
                    LayoutCachedHeight =14475
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =135
                            Top =1620
                            Width =15765
                            Height =12720
                            Name ="pgPhotos"
                            Caption ="Photos"
                            LayoutCachedLeft =135
                            LayoutCachedTop =1620
                            LayoutCachedWidth =15900
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

                                    LayoutCachedLeft =360
                                    LayoutCachedTop =1680
                                    LayoutCachedWidth =14040
                                    LayoutCachedHeight =8640
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =135
                            Top =1620
                            Width =15765
                            Height =12720
                            Name ="Point Intercept"
                            EventProcPrefix ="Point_Intercept"
                            Caption ="Point Intercept"
                            LayoutCachedLeft =135
                            LayoutCachedTop =1620
                            LayoutCachedWidth =15900
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

                                    LayoutCachedLeft =180
                                    LayoutCachedTop =1800
                                    LayoutCachedWidth =14280
                                    LayoutCachedHeight =11715
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =135
                            Top =1620
                            Width =15765
                            Height =12720
                            Name ="pgBeltShrub"
                            Caption ="1-m Belt"
                            LayoutCachedLeft =135
                            LayoutCachedTop =1620
                            LayoutCachedWidth =15900
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
                                    Width =15534
                                    Height =9504
                                    Name ="frm_LP_Belt_Transect"
                                    SourceObject ="Form.frm_LP_Belt_Transect"
                                    LinkChildFields ="Event_ID"
                                    LinkMasterFields ="Event_ID"

                                    LayoutCachedLeft =360
                                    LayoutCachedTop =1680
                                    LayoutCachedWidth =15894
                                    LayoutCachedHeight =11184
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =135
                            Top =1620
                            Width =15765
                            Height =12720
                            Name ="pgVegHeight"
                            Caption ="Vegetation Height"
                            LayoutCachedLeft =135
                            LayoutCachedTop =1620
                            LayoutCachedWidth =15900
                            LayoutCachedHeight =14340
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =300
                                    Top =1740
                                    Width =14250
                                    Height =9915
                                    Name ="frm_VH_Transect"
                                    SourceObject ="Form.frm_VH_Transect"
                                    LinkChildFields ="Event_ID"
                                    LinkMasterFields ="Event_ID"

                                    LayoutCachedLeft =300
                                    LayoutCachedTop =1740
                                    LayoutCachedWidth =14550
                                    LayoutCachedHeight =11655
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =135
                            Top =1620
                            Width =15765
                            Height =12720
                            Name ="pgGaps"
                            Caption ="Gap Intercepts"
                            LayoutCachedLeft =135
                            LayoutCachedTop =1620
                            LayoutCachedWidth =15900
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

                                    LayoutCachedLeft =180
                                    LayoutCachedTop =1680
                                    LayoutCachedWidth =13950
                                    LayoutCachedHeight =12255
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =135
                            Top =1620
                            Width =15765
                            Height =12720
                            Name ="pgSS"
                            Caption ="Soil Stability"
                            LayoutCachedLeft =135
                            LayoutCachedTop =1620
                            LayoutCachedWidth =15900
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

                                    LayoutCachedLeft =180
                                    LayoutCachedTop =1740
                                    LayoutCachedWidth =14040
                                    LayoutCachedHeight =8955
                                End
                            End
                        End
                        Begin Page
                            Visible = NotDefault
                            OverlapFlags =215
                            Left =135
                            Top =1620
                            Width =15765
                            Height =12720
                            Name ="pgSLIntercept"
                            Caption ="SL Intercept"
                            LayoutCachedLeft =135
                            LayoutCachedTop =1620
                            LayoutCachedWidth =15900
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

                                    LayoutCachedLeft =1440
                                    LayoutCachedTop =1800
                                    LayoutCachedWidth =12990
                                    LayoutCachedHeight =10455
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =60
                            Top =1620
                            Width =15840
                            Height =12720
                            Name ="pgOT"
                            Caption ="Overstory Trees"
                            LayoutCachedLeft =60
                            LayoutCachedTop =1620
                            LayoutCachedWidth =15900
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
                                    Width =11934
                                    Height =3780
                                    Name ="fsub_OT_Tree_Saplings"
                                    SourceObject ="Form.fsub_OT_Tree_Saplings"
                                    LinkChildFields ="Event_ID"
                                    LinkMasterFields ="Event_ID"

                                    LayoutCachedLeft =1440
                                    LayoutCachedTop =2040
                                    LayoutCachedWidth =13374
                                    LayoutCachedHeight =5820
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    Left =60
                                    Top =6300
                                    Width =14754
                                    Height =3780
                                    TabIndex =1
                                    Name ="fsub_OT_Census"
                                    SourceObject ="Form.fsub_OT_Census"
                                    LinkChildFields ="Event_ID"
                                    LinkMasterFields ="Event_ID"

                                    LayoutCachedLeft =60
                                    LayoutCachedTop =6300
                                    LayoutCachedWidth =14814
                                    LayoutCachedHeight =10080
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

                                    LayoutCachedLeft =3780
                                    LayoutCachedTop =1680
                                    LayoutCachedWidth =4800
                                    LayoutCachedHeight =1920
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
                                            LayoutCachedLeft =3240
                                            LayoutCachedTop =1680
                                            LayoutCachedWidth =3780
                                            LayoutCachedHeight =1920
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

                                    LayoutCachedLeft =6000
                                    LayoutCachedTop =1680
                                    LayoutCachedWidth =7440
                                    LayoutCachedHeight =1920
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
                                            LayoutCachedLeft =5100
                                            LayoutCachedTop =1680
                                            LayoutCachedWidth =5940
                                            LayoutCachedHeight =1925
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

                                    LayoutCachedLeft =8640
                                    LayoutCachedTop =1680
                                    LayoutCachedWidth =10080
                                    LayoutCachedHeight =1920
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
                                            LayoutCachedLeft =7740
                                            LayoutCachedTop =1680
                                            LayoutCachedWidth =8580
                                            LayoutCachedHeight =1925
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

                                    LayoutCachedLeft =3900
                                    LayoutCachedTop =5940
                                    LayoutCachedWidth =4919
                                    LayoutCachedHeight =6180
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
                                            LayoutCachedLeft =3360
                                            LayoutCachedTop =5940
                                            LayoutCachedWidth =3900
                                            LayoutCachedHeight =6180
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

                                    LayoutCachedLeft =6180
                                    LayoutCachedTop =5940
                                    LayoutCachedWidth =7620
                                    LayoutCachedHeight =6180
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
                                            LayoutCachedLeft =5220
                                            LayoutCachedTop =5940
                                            LayoutCachedWidth =6105
                                            LayoutCachedHeight =6185
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

                                    LayoutCachedLeft =8820
                                    LayoutCachedTop =5940
                                    LayoutCachedWidth =10260
                                    LayoutCachedHeight =6180
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
                                            LayoutCachedLeft =7920
                                            LayoutCachedTop =5940
                                            LayoutCachedWidth =8759
                                            LayoutCachedHeight =6185
                                        End
                                    End
                                End
                                Begin Rectangle
                                    SpecialEffect =0
                                    BackStyle =1
                                    OldBorderStyle =0
                                    OverlapFlags =223
                                    Left =11760
                                    Top =5820
                                    Width =3000
                                    Height =480
                                    BackColor =6750207
                                    Name ="rctNoCensus"
                                    OnClick ="[Event Procedure]"
                                    LayoutCachedLeft =11760
                                    LayoutCachedTop =5820
                                    LayoutCachedWidth =14760
                                    LayoutCachedHeight =6300
                                End
                                Begin CheckBox
                                    OverlapFlags =215
                                    Left =11880
                                    Top =5970
                                    Width =300
                                    TabIndex =8
                                    Name ="cbxNoCensus"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="No overstory species"

                                    LayoutCachedLeft =11880
                                    LayoutCachedTop =5970
                                    LayoutCachedWidth =12180
                                    LayoutCachedHeight =6210
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontFamily =0
                                            Left =12110
                                            Top =5940
                                            Width =2520
                                            Height =240
                                            FontWeight =600
                                            Name ="lblNoCensus"
                                            Caption ="No Overstory Species Found"
                                            ControlTipText ="No overstory species"
                                            LayoutCachedLeft =12110
                                            LayoutCachedTop =5940
                                            LayoutCachedWidth =14630
                                            LayoutCachedHeight =6180
                                        End
                                    End
                                End
                                Begin Rectangle
                                    Visible = NotDefault
                                    SpecialEffect =0
                                    BackStyle =1
                                    OldBorderStyle =0
                                    OverlapFlags =223
                                    Left =11220
                                    Top =1620
                                    Width =2100
                                    Height =480
                                    BackColor =6750207
                                    Name ="rctNoSaplings"
                                    OnClick ="[Event Procedure]"
                                    LayoutCachedLeft =11220
                                    LayoutCachedTop =1620
                                    LayoutCachedWidth =13320
                                    LayoutCachedHeight =2100
                                End
                                Begin CheckBox
                                    Enabled = NotDefault
                                    OverlapFlags =215
                                    Left =11340
                                    Top =1770
                                    Width =300
                                    TabIndex =9
                                    Name ="cbxNoSaplings"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="No tree saplings found"

                                    LayoutCachedLeft =11340
                                    LayoutCachedTop =1770
                                    LayoutCachedWidth =11640
                                    LayoutCachedHeight =2010
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontFamily =0
                                            Left =11570
                                            Top =1740
                                            Width =1650
                                            Height =240
                                            FontWeight =600
                                            Name ="lblNoSaplings"
                                            Caption ="No Saplings Found"
                                            ControlTipText ="No tree saplings found"
                                            LayoutCachedLeft =11570
                                            LayoutCachedTop =1740
                                            LayoutCachedWidth =13220
                                            LayoutCachedHeight =1980
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            Visible = NotDefault
                            OverlapFlags =247
                            Left =135
                            Top =1620
                            Width =15765
                            Height =12720
                            Name ="pgFuels"
                            Caption ="Fuels"
                            LayoutCachedLeft =135
                            LayoutCachedTop =1620
                            LayoutCachedWidth =15900
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
                                    Name ="FuelsDate"
                                    ControlSource ="Fuels_Date"
                                    Format ="Short Date"
                                    FontName ="Tahoma"
                                    InputMask ="99/99/0000;0;_"

                                    LayoutCachedLeft =4320
                                    LayoutCachedTop =1740
                                    LayoutCachedWidth =5280
                                    LayoutCachedHeight =1980
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
                                            LayoutCachedLeft =2700
                                            LayoutCachedTop =1740
                                            LayoutCachedWidth =4320
                                            LayoutCachedHeight =1980
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
                                    TabIndex =1
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                                    Name ="FuelsObserver"
                                    ControlSource ="Fuels_Observer"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tlu_Contacts.Contact_ID, tlu_Contacts.Last_Name, tlu_Contacts.First_Name "
                                        "FROM tlu_Contacts WHERE (((tlu_Contacts.Active)=1)); "
                                    ColumnWidths ="0;990;990"
                                    FontName ="Tahoma"

                                    LayoutCachedLeft =6420
                                    LayoutCachedTop =1740
                                    LayoutCachedWidth =7800
                                    LayoutCachedHeight =1980
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
                                            LayoutCachedLeft =5580
                                            LayoutCachedTop =1740
                                            LayoutCachedWidth =6420
                                            LayoutCachedHeight =1985
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
                                    TabIndex =2
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                                    Name ="FuelsRecorder"
                                    ControlSource ="Fuels_Recorder"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tlu_Contacts.Contact_ID, tlu_Contacts.Last_Name, tlu_Contacts.First_Name "
                                        "FROM tlu_Contacts WHERE (((tlu_Contacts.Active)=1)); "
                                    ColumnWidths ="0;990;990"
                                    FontName ="Tahoma"

                                    LayoutCachedLeft =8940
                                    LayoutCachedTop =1740
                                    LayoutCachedWidth =10320
                                    LayoutCachedHeight =1980
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
                                            LayoutCachedLeft =8100
                                            LayoutCachedTop =1740
                                            LayoutCachedWidth =8940
                                            LayoutCachedHeight =1985
                                        End
                                    End
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    Left =360
                                    Top =2160
                                    Width =13140
                                    Height =4680
                                    TabIndex =3
                                    Name ="fsub_Fuels_LD"
                                    SourceObject ="Form.fsub_Fuels_LD"
                                    LinkChildFields ="Event_ID"
                                    LinkMasterFields ="Event_ID"

                                    LayoutCachedLeft =360
                                    LayoutCachedTop =2160
                                    LayoutCachedWidth =13500
                                    LayoutCachedHeight =6840
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    Left =4020
                                    Top =7260
                                    Width =5940
                                    Height =3960
                                    TabIndex =4
                                    Name ="fsub_Fuels_1000"
                                    SourceObject ="Form.fsub_Fuels_1000"
                                    LinkChildFields ="Event_ID"
                                    LinkMasterFields ="Event_ID"

                                    LayoutCachedLeft =4020
                                    LayoutCachedTop =7260
                                    LayoutCachedWidth =9960
                                    LayoutCachedHeight =11220
                                End
                                Begin Rectangle
                                    SpecialEffect =0
                                    BackStyle =1
                                    OldBorderStyle =0
                                    OverlapFlags =255
                                    Left =10080
                                    Top =7860
                                    Width =2760
                                    Height =480
                                    BackColor =6750207
                                    Name ="rctNo1000hrA"
                                    OnClick ="[Event Procedure]"
                                    LayoutCachedLeft =10080
                                    LayoutCachedTop =7860
                                    LayoutCachedWidth =12840
                                    LayoutCachedHeight =8340
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =10200
                                    Top =8010
                                    Width =300
                                    TabIndex =5
                                    Name ="cbxNo1000hrA"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="No 1000-hr fuels found in transect A"

                                    LayoutCachedLeft =10200
                                    LayoutCachedTop =8010
                                    LayoutCachedWidth =10500
                                    LayoutCachedHeight =8250
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontFamily =0
                                            Left =10428
                                            Top =7980
                                            Width =2310
                                            Height =240
                                            FontWeight =600
                                            Name ="lblNo1000hrA"
                                            Caption ="No A 1000-hr Fuels Found"
                                            ControlTipText ="No 1000-hr fuels found in transect A"
                                            LayoutCachedLeft =10428
                                            LayoutCachedTop =7980
                                            LayoutCachedWidth =12738
                                            LayoutCachedHeight =8220
                                        End
                                    End
                                End
                                Begin Rectangle
                                    SpecialEffect =0
                                    BackStyle =1
                                    OldBorderStyle =0
                                    OverlapFlags =255
                                    Left =10080
                                    Top =8460
                                    Width =2760
                                    Height =480
                                    BackColor =6750207
                                    Name ="rctNo1000hrB"
                                    OnClick ="[Event Procedure]"
                                    LayoutCachedLeft =10080
                                    LayoutCachedTop =8460
                                    LayoutCachedWidth =12840
                                    LayoutCachedHeight =8940
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =10200
                                    Top =8610
                                    Width =300
                                    TabIndex =6
                                    Name ="cbxNo1000hrB"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="No 1000-hr fuels found in transect B"

                                    LayoutCachedLeft =10200
                                    LayoutCachedTop =8610
                                    LayoutCachedWidth =10500
                                    LayoutCachedHeight =8850
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontFamily =0
                                            Left =10428
                                            Top =8580
                                            Width =2310
                                            Height =240
                                            FontWeight =600
                                            Name ="lblNo1000hrB"
                                            Caption ="No B 1000-hr Fuels Found"
                                            ControlTipText ="No 1000-hr fuels found in transect B"
                                            LayoutCachedLeft =10428
                                            LayoutCachedTop =8580
                                            LayoutCachedWidth =12738
                                            LayoutCachedHeight =8820
                                        End
                                    End
                                End
                                Begin Rectangle
                                    SpecialEffect =0
                                    BackStyle =1
                                    OldBorderStyle =0
                                    OverlapFlags =255
                                    Left =10080
                                    Top =9060
                                    Width =2760
                                    Height =480
                                    BackColor =6750207
                                    Name ="rctNo1000hrC"
                                    OnClick ="[Event Procedure]"
                                    LayoutCachedLeft =10080
                                    LayoutCachedTop =9060
                                    LayoutCachedWidth =12840
                                    LayoutCachedHeight =9540
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =10200
                                    Top =9210
                                    Width =300
                                    TabIndex =7
                                    Name ="cbxNo1000hrC"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="No 1000-hr fuels found in transect C"

                                    LayoutCachedLeft =10200
                                    LayoutCachedTop =9210
                                    LayoutCachedWidth =10500
                                    LayoutCachedHeight =9450
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontFamily =0
                                            Left =10428
                                            Top =9180
                                            Width =2310
                                            Height =240
                                            FontWeight =600
                                            Name ="lblNo1000hrC"
                                            Caption ="No C 1000-hr Fuels Found"
                                            ControlTipText ="No 1000-hr fuels found in transect C"
                                            LayoutCachedLeft =10428
                                            LayoutCachedTop =9180
                                            LayoutCachedWidth =12738
                                            LayoutCachedHeight =9420
                                        End
                                    End
                                End
                                Begin Rectangle
                                    SpecialEffect =0
                                    BackStyle =1
                                    OldBorderStyle =0
                                    OverlapFlags =255
                                    Left =10080
                                    Top =9660
                                    Width =2760
                                    Height =480
                                    BackColor =6750207
                                    Name ="rctNo1000hrD"
                                    OnClick ="[Event Procedure]"
                                    LayoutCachedLeft =10080
                                    LayoutCachedTop =9660
                                    LayoutCachedWidth =12840
                                    LayoutCachedHeight =10140
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =10200
                                    Top =9810
                                    Width =300
                                    TabIndex =8
                                    Name ="cbxNo1000hrD"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="No 1000-hr fuels found in transect D"

                                    LayoutCachedLeft =10200
                                    LayoutCachedTop =9810
                                    LayoutCachedWidth =10500
                                    LayoutCachedHeight =10050
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontFamily =0
                                            Left =10428
                                            Top =9780
                                            Width =2325
                                            Height =240
                                            FontWeight =600
                                            Name ="lblNo1000hrD"
                                            Caption ="No D 1000-hr Fuels Found"
                                            ControlTipText ="No 1000-hr fuels found in transect D"
                                            LayoutCachedLeft =10428
                                            LayoutCachedTop =9780
                                            LayoutCachedWidth =12753
                                            LayoutCachedHeight =10020
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =135
                            Top =1620
                            Width =15765
                            Height =12720
                            Name ="pgImpact"
                            Caption ="Site Impact"
                            LayoutCachedLeft =135
                            LayoutCachedTop =1620
                            LayoutCachedWidth =15900
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

                                    LayoutCachedLeft =180
                                    LayoutCachedTop =1680
                                    LayoutCachedWidth =12810
                                    LayoutCachedHeight =11520
                                End
                            End
                        End
                        Begin Page
                            Visible = NotDefault
                            OverlapFlags =247
                            Left =135
                            Top =1620
                            Width =15765
                            Height =12720
                            Name ="pgCoords_and_loc_details"
                            Caption ="Presence Cover Density"
                            LayoutCachedLeft =135
                            LayoutCachedTop =1620
                            LayoutCachedWidth =15900
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

                                    LayoutCachedLeft =180
                                    LayoutCachedTop =1665
                                    LayoutCachedWidth =13890
                                    LayoutCachedHeight =13680
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
                    Name ="btnCoord"
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
                    Name ="btnComments"
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
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =255
                    Left =10080
                    Top =7260
                    Width =2760
                    Height =480
                    BackColor =6750207
                    Name ="rctNo1000hr"
                    OnClick ="[Event Procedure]"
                    LayoutCachedLeft =10080
                    LayoutCachedTop =7260
                    LayoutCachedWidth =12840
                    LayoutCachedHeight =7740
                End
                Begin CheckBox
                    OverlapFlags =255
                    Left =10200
                    Top =7410
                    Width =300
                    TabIndex =15
                    Name ="cbxNo1000hr"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="No 1000-hr fuels found"

                    LayoutCachedLeft =10200
                    LayoutCachedTop =7410
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =7650
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =10430
                            Top =7380
                            Width =2115
                            Height =240
                            FontWeight =600
                            Name ="lblNo1000hr"
                            Caption ="No 1000-hr Fuels Found"
                            ControlTipText ="No 1000-hr fuels found"
                            LayoutCachedLeft =10430
                            LayoutCachedTop =7380
                            LayoutCachedWidth =12545
                            LayoutCachedHeight =7620
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =13860
                    Top =180
                    Width =960
                    Height =900
                    FontSize =11
                    TabIndex =16
                    Name ="btnPlotQAQC"
                    Caption ="Plot Check!"
                    StatusBarText ="Check Field Data!"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00ddddddddddddddddddd0000000000ddddd0ffffffffff0dd ,
                        0xdd0fff88fffff0dddd0ff8188ffff0dddd0f811188fff0dddd0f11f118fff0dd ,
                        0xdd0fffff178ff0dddd0ffffff188f0dddd0fffffff18f0dddd0ffffffff1f0dd ,
                        0xdd0ffffffffff0dddd0ff000000ff0ddddd000f888000ddddddddd0000dddddd ,
                        0xdddddddddddddddd000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    FontName ="Calibri"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Check Field Data!"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638

                    CursorOnHover =1
                    LayoutCachedLeft =13860
                    LayoutCachedTop =180
                    LayoutCachedWidth =14820
                    LayoutCachedHeight =1080
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Shape =2
                    Gradient =12
                    BackColor =5066944
                    BackThemeColorIndex =5
                    BorderColor =5066944
                    BorderThemeColorIndex =5
                    ThemeFontIndex =1
                    HoverColor =15709952
                    PressedColor =15709952
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =24
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
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
' MODULE:       frm_Data_Entry
' Level:        Form module
' Version:      1.10
' Description:  data functions & procedures specific to field data entry
'
' Data source:  tbl_Locations
' Data access:  edit; allow additions off except for new records
' Pages:        none
' Functions:    none
' References:   fxnSwitchboardIsOpen, fxnGUIDGen
' Source/date:  John R. Boetsch, June 2006
' Adapted:      Bonnie Campbell, 2/2/2016
' Revisions:    RDB - unknown  - 1.00 - initial version
'               BLC - 2/3/2016 - 1.01 - added documentation, enabled seedlings & saplings for
'                                       oak scrub plots, revised to use transect overlay vs.
'                                       message box
'               BLC - 3/7/2016 - 1.02 - fixed bugs setting "Overstory-Census" vs. "OverstoryTree-Census",
'                                       and leaving rctNoShrubs visible for oak scrub plots
'               BLC - 3/21/2016 - 1.03 - fixed bug where Fuels tab was visible on grassland/shrubland plots
'                                        added, added 1000hr fuel A-D no data collected checkboxes
'               BLC - 3/21/2016 - 1.04 - revised Form_Load to:
'                                        expose: Gap Intercepts - grassland/shrubland
'                                                Soil Stability - grassland/shrubland, woodland
'                                        hide:   SL Intercept - oak scrub
'               BLC - 4/13/2016 - 1.05 - changed form properties to avoid taskbar overlap (Scrollbars = Both, Border Style = Sizeable)
'                                        allows users to resize & scroll to expose taskbar apps/documents
'                                        original values (Scrollbars = Neither, Border Style = Thin)
'                                        added refresh for underlying subforms for conditional formatting
'               BLC - 3/22/2017 - 1.06 - added documentation, error handling, btnPlotQAQC
'               BLC - 3/24/2017 - 1.07 - added CallingForm property, added TempVar("ParkCode")
'               BLC - 3/31/2017 - 1.08 - added RemoveTemplateQueries to clear queries on form close
'               BLC - 8/10/2017 - 1.09 - revised to open PlotCheck w/ WindowMode as
'                                        acWindowNormal vs. acDialog to prevent form from
'                                        behaving as a modal dialog
'               BLC - 2/1/2018  - 1.10 - revised to run PlotCheckSelect vs PlotCheck form
' =================================
'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Private m_CallingForm As String

'---------------------
' Event Declarations
'---------------------

'---------------------
' Properties
'---------------------

'---------------------
' Methods
'---------------------

' ---------------------------------
' SUB:          Form_Open
' Description:  Handles form opening actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 22, 2017 - for NCPN tools
' Revisions:
'   JRB, 6/x/2006  - initial version
'   RDB, unknown   - ?
'   BLC, 2/2/2016  - added documentation
'   BLC, 3/22/2017 - added btnPlotQAQC initalization
'   BLC, 3/24/2017 - removed btnPlotQAQC initialization & added TempVar("ParkCode")
' ---------------------------------
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
    
    'set park
    SetTempVar "ParkCode", Nz(Me.Unit_Code, "")
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[Form_frm_Data_Entry])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_Load
' Description:  Handles form loading actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch - June, 2006
' Adapted:      Bonnie Campbell, March 22, 2017 - for NCPN tools
' Revisions:
'   JRB, 6/x/2006  - initial version
'   RDB, unknown   - ?
'   BLC, 2/2/2016  - added documentation, enabled seedling & saplings data entry
'                    for oak scrub plots
'   BLC, 3/7/2016  - fixed issues where NoShrubs rectangle showed for oak scrub plots when
'                    it should not & enabling pgFuels for woodland & forest plots
'   BLC, 3/16/2016 - fixed bugs where: Fuels tab was visible on grassland/shrubland plots,
'                    SL Intercept tab visible for other than oak scrub plots
'   BLC, 3/21/2016 - handled transect A-D 1000hr fuels
'   BLC, 3/23/2016 - revised Form_Load to:
'                    expose: Gap Intercepts - grassland/shrubland
'                            Soil Stability - grassland/shrubland, woodland
'                    hide:   SL Intercept - oak scrub
'                    added more documentation for tabs exposed/hidden
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler

'------------------------
' set no data checkboxes
'------------------------
    Dim dNoDataEvent As Scripting.Dictionary
    Dim dNoDataTransect As Scripting.Dictionary

    'fetch no data info & set checkboxes
    'event level values
    Set dNoDataEvent = GetNoDataCollected(Me.Event_ID, "E")
    
    With dNoDataEvent
        Me.cbxNoSaplings.Value = .Item("OverstoryTree-Sapling")
        Me.cbxNoCensus.Value = .Item("OverstoryTree-Census")
    
        Me.cbxNo1000hr.Value = .Item("Fuel-1000hr")
        Me.cbxNo1000hrA.Value = .Item("Fuel-1000hr-A")
        Me.cbxNo1000hrB.Value = .Item("Fuel-1000hr-B")
        Me.cbxNo1000hrC.Value = .Item("Fuel-1000hr-C")
        Me.cbxNo1000hrD.Value = .Item("Fuel-1000hr-D")
        
        Me.frm_Site_Impact.Form.Controls("cbxNoDisturbance").Value = .Item("SiteImpact-Disturbance")
        Me.frm_Site_Impact.Form.Controls("cbxNoSpecies").Value = .Item("SiteImpact-Exotic")
    End With
    
    'transect level values -> see LP_Belt_Transect
   
    Me.rctNoSaplings.Visible = (Me.fsub_OT_Tree_Saplings.Form.RecordsetClone.RecordCount = 0)
    Me.rctNoCensus.Visible = (Me.fsub_OT_Census.Form.RecordsetClone.RecordCount = 0)
    
    Me.rctNo1000hr.Visible = (Me.fsub_Fuels_1000.Form.RecordsetClone.RecordCount = 0)

    'A-D are set via Check1000hrFuels (more granular than RecordCount alone)
    'A-D highlighting is displayed when no records exist
    If rctNo1000hr.Visible = True Then
        Me.rctNo1000hrA.Visible = True
        Me.rctNo1000hrB.Visible = True
        Me.rctNo1000hrC.Visible = True
        Me.rctNo1000hrD.Visible = True
    End If
       
    Me.frm_Site_Impact.Form.Controls("rctNoDisturbance").Visible = (Me.frm_Site_Impact.Form.Controls("Disturbance Details").Form.RecordsetClone.RecordCount = 0)
    Me.frm_Site_Impact.Form.Controls("rctNoSpecies").Visible = (Me.frm_Site_Impact.Form.Controls("fsub_Dist_Exotic").Form.RecordsetClone.RecordCount = 0)
    
    'disable checkboxes if records exist
    Me.cbxNoSaplings.Enabled = (Me.fsub_OT_Tree_Saplings.Form.RecordsetClone.RecordCount = 0)
    Me.cbxNoCensus.Enabled = (Me.fsub_OT_Census.Form.RecordsetClone.RecordCount = 0)

    Me.cbxNo1000hr.Enabled = (Me.fsub_Fuels_1000.Form.RecordsetClone.RecordCount = 0)
    
    'A-D are set via Check1000hrFuels (more granular than RecordCount alone)
    Check1000hrFuels

    Me.frm_Site_Impact.Form.Controls("cbxNoDisturbance").Enabled = (Me.frm_Site_Impact.Form.Controls("Disturbance Details").Form.RecordsetClone.RecordCount = 0)
    Me.frm_Site_Impact.Form.Controls("cbxNoSpecies").Enabled = (Me.frm_Site_Impact.Form.Controls("fsub_Dist_Exotic").Form.RecordsetClone.RecordCount = 0)

    '---------------------------------------
    ' display proper tabs based on veg type
    '---------------------------------------
    Dim Veg_Type As Variant

    Veg_Type = DLookup("[Vegetation_Type]", "tbl_Locations", "[Location_ID] = '" & Me!Location_ID & "'")
    
    '---------------------
    ' tab displays
    '---------------------
    ' SS = soil stability, Gaps = Gap Intercepts,
    ' all (exceptions below): hide - SLIntercept, Gaps
    '                         expose - Photos, Point Intercept, 1-m Belt, Overstory Trees, Site Impact,
    '                                  Fuels, Soil Stability
    '
    '       X = exposed  O = hidden
    '
    '                               DEFAULT     Forest  Grassland/Shrubland     Oak Scrub   Woodland
    '   Photos (pgPhotos)             X           X             X                   X           X
    '   Point Intercept               X           X             X                   X           X
    '   1-m Belt (pgBeltShrub)        X           X             X                   X           X
    '   Gap Intercepts (pgGaps)       O           O             X                   O           O
    '   Overstory Trees (pgOT)        X           X             X                   X           X
    '   Site Impact (pgImpact)        X           X             X                   X           X
    '   Soil Stability (pgSS)         X           O             X                   O           X
    '   Fuels (pgFuels)               X           X             O                   O           X
    '----------------------
    
    'defaults
    '-- hide  --
    '[deprecated, no longer used tabs]
    Me!pgSLIntercept.Visible = False
    Me!pgCoords_and_loc_details.Visible = False
    'normal tabs
    Me!pgGaps.Visible = False
    Me!pgVegHeight.Visible = False
    
    '-- expose --
    Me!pgPhotos.Visible = True
    Me![Point Intercept].Visible = True
    Me!pgBeltShrub.Visible = True
    Me!pgOT.Visible = True
    Me.pgImpact.Visible = True
    Me!pgFuels.Visible = True
    Me!pgSS.Visible = True
    
    'handle exceptions
    If Not IsNull(Veg_Type) Then
        Select Case Veg_Type
            
            Case "forest"   'hide SS & Gaps, TICA no fuels special case
                'hide SS & Gaps (above)
                Me!pgSS.Visible = False
                
                ' Modified to hide fuels form for TICA 1 [HMT, 3/13/2015]
                ' TICA 1 is a special case of a forest plot that does not have fuels data collected.
                If (Me!Unit_Code = "TICA") And (Me!Plot_ID = 1) Then
                  Me!pgFuels.Visible = False
                End If
            
            Case "grassland/shrubland"  'hide fuels, expose SS & Gaps
                'hide fuels
                Me!pgFuels.Visible = False
                            
                'expose SS (above) & Gaps
                Me!pgGaps.Visible = True
                Me!pgVegHeight.Visible = True
                            
            Case "oak scrub"    'oak plots   hide Gaps, SS & fuels,  expose
                'hide fuels, SS & Gaps (above)
                Me!pgFuels.Visible = False
                Me!pgSS.Visible = False
                
                'hide shrubs
                Me.frm_LP_Belt_Transect.Controls("cbxNoShrubs").Visible = False
                Me.frm_LP_Belt_Transect.Controls("lblNoShrubs").Visible = False
                Me.frm_LP_Belt_Transect.Controls("rctNoShrubs").Visible = False
            
                'expose SL intercept for oak scrub --> NOW hide it (3/23/2016)
                Me!pgSLIntercept.Visible = False
            
            Case "woodland"
                'hide gaps (above) & OT Census crown class
                Me!fsub_OT_Census.Form!Crown_Class.Visible = False
                Me!fsub_OT_Census.Form!Crown_Class_Label.Visible = False
                'expose SS (above)

        End Select
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

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[Form_frm_Data_Entry])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_Current
' Description:  Handles form current actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch - June, 2006
' Adapted:      Bonnie Campbell, March 22, 2017 - for NCPN tools
' Revisions:
'   JRB, 6/x/2006  - initial version
'   BLC, 3/22/2017 - added documentation, error handling
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler

    Update_Loc_Info
    If Not IsNull(Me!txtUnit_Code) Then
    '  MsgBox DLookup("[ParkState]", "tlu_Parks", "[ParkCode] = '" & Me!txtUnit_Code & "'")
      Me!frm_Quadrat_Transect.Form!fsub_Quadrat.Form!fsub_Quadrat_Shrubs.Form!State_Code = DLookup("[ParkState]", "tlu_Parks", "[ParkCode] = '" & Me!txtUnit_Code & "'")
      Me!frm_Quadrat_Transect.Form!fsub_Quadrat.Form!fsub_Species.Form!State_Code = DLookup("[ParkState]", "tlu_Parks", "[ParkCode] = '" & Me!txtUnit_Code & "'")
      Me!SiteDisplay = cboLocation_ID.Column(1)  ' Display the site number in heading
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[Form_frm_Data_Entry])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_BeforeUpdate
' Description:  Handles form before update actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch - June, 2006
' Adapted:      Bonnie Campbell, March 22, 2017 - for NCPN tools
' Revisions:
'   JRB, 6/x/2006  - initial version
'   BLC, 3/22/2017 - added documentation, error handling
' ---------------------------------
Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

  If IsNull(Me!Start_Date) Then
        ' ask user if (s)he wants to enter data or cancel and close form
        If MsgBox("Visit date is missing - do you want to enter the missing data?", vbYesNo, "Date missing") = vbNo Then
            Me.Undo
        End If
  End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeUpdate[Form_frm_Data_Entry])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_BeforeInsert
' Description:  Handles form before insert actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch - June, 2006
' Adapted:      Bonnie Campbell, March 22, 2017 - for NCPN tools
' Revisions:
'   JRB, 6/x/2006  - initial version
'   BLC, 3/22/2017 - added documentation, error handling
' ---------------------------------
Private Sub Form_BeforeInsert(Cancel As Integer)
On Error GoTo Err_Handler

        Dim db As dao.Database
        Dim Versions As dao.Recordset
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

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeInsert[Form_frm_Data_Entry])"
    End Select
    Resume Exit_Handler
End Sub

' =================================
'   No Data Checkboxes
' =================================
' ---------------------------------
' SUB:          cbxNoSaplings_Click
' Description:  Handles No Saplings checkbox actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 3, 2016 - for NCPN tools
' Revisions:
'   BLC, 2/3/2016  - initial version
'   BLC, 4/13/2016 - added requery of related subform to clear/set conditional formatting on change
' ---------------------------------
Private Sub cbxNoSaplings_Click()
On Error GoTo Err_Handler

    'set dictionary & db value (abs is used to drive 1 = true, 0 = false since -1 is true in access/vba)
    SetNoDataCollected Me.Event_ID, "E", "OverstoryTree-Sapling", Abs(Me.cbxNoSaplings.Value)
    
    'refresh the subform to clear conditional formatting
    Me.fsub_OT_Tree_Saplings.Requery

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxNoSaplings_Click[Form_frm_Data_Entry])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxNoCensus_Click
' Description:  Handles No Overstory Species checkbox actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 3, 2016 - for NCPN tools
' Revisions:
'   BLC, 2/3/2016  - initial version
'   BLC, 3/7/2016  - fixed bug setting "Overstory-Census" vs. "OverstoryTree-Census"
'   BLC, 4/13/2016 - added requery of related subform to clear/set conditional formatting on change
' ---------------------------------
Private Sub cbxNoCensus_Click()
On Error GoTo Err_Handler

    'set dictionary & db value (abs is used to drive 1 = true, 0 = false since -1 is true in access/vba)
    SetNoDataCollected Me.Event_ID, "E", "OverstoryTree-Census", Abs(Me.cbxNoCensus.Value)

    'refresh the subform to clear conditional formatting
    Me.fsub_OT_Census.Requery
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxNoOverstorySpecies_Click[Form_frm_Data_Entry])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxNo1000hr_Click
' Description:  Handles checkbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, February 9, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC, 2/9/2016  - initial version
'   BLC, 3/18/2016 - added A-D click to set these when 1000hr checkbox is checked
' ----------------------------------
Private Sub cbxNo1000hr_Click()
On Error GoTo Err_Handler

    'set dictionary & db value (abs is used to drive 1 = true, 0 = false since -1 is true in access/vba)
    SetNoDataCollected Me.Event_ID, "E", "Fuel-1000hr", Abs(Me.cbxNo1000hr.Value)

    'set A-D if checked
    If Abs(Me.cbxNo1000hr) = 1 Then
        Me.cbxNo1000hrA.Value = 1
        Me.cbxNo1000hrB.Value = 1
        Me.cbxNo1000hrC.Value = 1
        Me.cbxNo1000hrD.Value = 1
        SetNoDataCollected Me.Event_ID, "E", "Fuel-1000hr-A", 1
        SetNoDataCollected Me.Event_ID, "E", "Fuel-1000hr-B", 1
        SetNoDataCollected Me.Event_ID, "E", "Fuel-1000hr-C", 1
        SetNoDataCollected Me.Event_ID, "E", "Fuel-1000hr-D", 1
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxNoDisturbance_Click[Form_frm_Data_Entry])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxNo1000hrA_Click
' Description:  Handles checkbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 18, 2016 - for NCPN tools
' Revisions:
'   BLC, 3/18/2016  - initial version
' ----------------------------------
Private Sub cbxNo1000hrA_Click()
On Error GoTo Err_Handler

    'set dictionary & db value (abs is used to drive 1 = true, 0 = false since -1 is true in access/vba)
    SetNoDataCollected Me.Event_ID, "E", "Fuel-1000hr-A", Abs(Me.cbxNo1000hrA.Value)

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxNo1000hrA_Click[Form_frm_Data_Entry])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxNo1000hrB_Click
' Description:  Handles checkbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 18, 2016 - for NCPN tools
' Revisions:
'   BLC, 3/18/2016  - initial version
' ----------------------------------
Private Sub cbxNo1000hrB_Click()
On Error GoTo Err_Handler

    'set dictionary & db value (abs is used to drive 1 = true, 0 = false since -1 is true in access/vba)
    SetNoDataCollected Me.Event_ID, "E", "Fuel-1000hr-B", Abs(Me.cbxNo1000hrB.Value)

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxNo1000hrB_Click[Form_frm_Data_Entry])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxNo1000hrC_Click
' Description:  Handles checkbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 18, 2016 - for NCPN tools
' Revisions:
'   BLC, 3/18/2016  - initial version
' ----------------------------------
Private Sub cbxNo1000hrC_Click()
On Error GoTo Err_Handler

    'set dictionary & db value (abs is used to drive 1 = true, 0 = false since -1 is true in access/vba)
    SetNoDataCollected Me.Event_ID, "E", "Fuel-1000hr-C", Abs(Me.cbxNo1000hrC.Value)

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxNo1000hrC_Click[Form_frm_Data_Entry])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxNo1000hrD_Click
' Description:  Handles checkbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 18, 2016 - for NCPN tools
' Revisions:
'   BLC, 3/18/2016  - initial version
' ----------------------------------
Private Sub cbxNo1000hrD_Click()
On Error GoTo Err_Handler

    'set dictionary & db value (abs is used to drive 1 = true, 0 = false since -1 is true in access/vba)
    SetNoDataCollected Me.Event_ID, "E", "Fuel-1000hr-D", Abs(Me.cbxNo1000hrD.Value)

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxNo1000hrD_Click[Form_frm_Data_Entry])"
    End Select
    Resume Exit_Handler
End Sub



' ---------------------------------
' SUB:          rctNoSaplings_Click
' Description:  Handles No Saplings rectangle actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 3, 2016 - for NCPN tools
' Revisions:
'   BLC, 2/3/2016  - initial version
' ---------------------------------
Private Sub rctNoSaplings_Click()
On Error GoTo Err_Handler

    'activates No Saplings checkbox
    cbxNoSaplings_Click

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - rctNoSaplings_Click[Form_frm_Data_Entry])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          rctNoCensus_Click
' Description:  Handles No overstory census rectangle actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 3, 2016 - for NCPN tools
' Revisions:
'   BLC, 2/3/2016  - initial version
' ---------------------------------
Private Sub rctNoCensus_Click()
On Error GoTo Err_Handler

    'activates No Overstory Species checkbox
    cbxNoCensus_Click

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - rctNoCensus_Click[Form_frm_Data_Entry])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          rctNo1000hr_Click
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
Private Sub rctNo1000hr_Click()
On Error GoTo Err_Handler

    cbxNo1000hr_Click

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - rctNo1000hr_Click[Form_frm_Data_Entry])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          rctNo1000hrA_Click
' Description:  Handles rectangular box click actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 18, 2016 - for NCPN tools
' Revisions:
'   BLC, 3/18/2016  - initial version
' ----------------------------------
Private Sub rctNo1000hrA_Click()
On Error GoTo Err_Handler

    cbxNo1000hrA_Click

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - rctNo1000hrA_Click[Form_frm_Data_Entry])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          rctNo1000hrB_Click
' Description:  Handles rectangular box click actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 18, 2016 - for NCPN tools
' Revisions:
'   BLC, 3/18/2016  - initial version
' ----------------------------------
Private Sub rctNo1000hrB_Click()
On Error GoTo Err_Handler

    cbxNo1000hrB_Click

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - rctNo1000hrB_Click[Form_frm_Data_Entry])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          rctNo1000hrC_Click
' Description:  Handles rectangular box click actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 18, 2016 - for NCPN tools
' Revisions:
'   BLC, 3/18/2016  - initial version
' ----------------------------------
Private Sub rctNo1000hrC_Click()
On Error GoTo Err_Handler

    cbxNo1000hrC_Click

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - rctNo1000hrC_Click[Form_frm_Data_Entry])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          rctNo1000hrD_Click
' Description:  Handles rectangular box click actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 18, 2016 - for NCPN tools
' Revisions:
'   BLC, 3/18/2016  - initial version
' ----------------------------------
Private Sub rctNo1000hrD_Click()
On Error GoTo Err_Handler

    cbxNo1000hrD_Click

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - rctNo1000hrD_Click[Form_frm_Data_Entry])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cboLocation_ID_AfterUpdate()
' Update_Loc_Info
End Sub

' ---------------------------------
' SUB:          txtStart_Date_AfterUpdate
' Description:  Handles form closing actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch - June, 2006
' Adapted:      Bonnie Campbell, March 22, 2017 - for NCPN tools
' Revisions:
'   JRB, 6/x/2006  - initial version
'   BLC, 3/22/2017 - added documentation, error handling
' ---------------------------------
Private Sub txtStart_Date_AfterUpdate()
        Dim db As dao.Database
        Dim Events As dao.Recordset
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
      GoTo Exit_Handler
    End If
    Events.Close

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - txtStart_Date_AfterUpdate[Form_frm_Data_Entry])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          pgTabs_Change
' Description:  Handles tab change actions
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
'   BLC, 2/3/2016  - added documentation, revised transect number reminder to
'                    use transect number overlay
'   BLC, 3/7/2016  - fixed Error #2467 expression entered refers to object that is closed or does not exist when
'                    tabbing to 1-m belt transect
' ---------------------------------
Private Sub pgTabs_Change()
On Error GoTo Err_Handler

  Dim TransectNumber As Integer

  Select Case Me.pgTabs.Value  'RDB: Display a silly message so the field crews know where they are
    Case 0 'Tab: Photos
    Case 1 'Tab: Point Intercept
      If IsNull(Me!frm_LP_Transect.Form!Transect) Then
        TransectNumber = 1
      Else
        TransectNumber = Me!frm_LP_Transect.Form!Transect
      End If
    Case 2 'Tab: 1-m Belt
      If IsNull(Me!frm_LP_Belt_Transect.Form!Transect) Then
        TransectNumber = 1
      Else
        TransectNumber = Me!frm_LP_Belt_Transect.Form!Transect
      End If
     
     Case 3 'Vegetation Height
      If IsNull(Me!frm_VH_Transect.Form!Transect) Then
        TransectNumber = 1
      Else
        TransectNumber = Me!frm_VH_Transect.Form!Transect
      End If
    Case 4 'Tab: Gap Intercepts
      If IsNull(Me!frm_Canopy_Transect.Form!Transect) Then
        TransectNumber = 1
      Else
        TransectNumber = Me!frm_Canopy_Transect.Form!Transect
      End If
    Case 5 'Tab: Soil Stability
    Case 6 'Tab:
      If IsNull(Me!frm_SL_Transect.Form!Transect) Then
        TransectNumber = 1
      Else
        TransectNumber = Me!frm_SL_Transect.Form!Transect
      End If
    Case 7 'Tab:
    Case 8 'Tab: Overstory Trees
    Case 9 'Tab: Site Impact
  End Select
  
    '---------------------------
    'display overlay - 2/3/2016 - BLC
    '---------------------------
    'MsgBox "You are on transect " & TransectNumber & ".", 0, "Transect Verify"
    DoCmd.OpenForm "frm_Transect_Overlay", OpenArgs:=TransectNumber
    '---------------------------

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - pgTabs_Change[Form_frm_Data_Entry])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnCoord_Click
' Description:  Handles coordinate button actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch - June, 2006
' Adapted:      Bonnie Campbell, March 22, 2017 - for NCPN tools
' Revisions:
'   JRB, 6/x/2006  - initial version
'   BLC, 3/22/2017 - added documentation, error handling,
'                    renamed buttonCoord to btnCoord
' ---------------------------------
Private Sub btnCoord_Click()
On Error GoTo Err_Handler

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Location_Modify"
    
    stLinkCriteria = "[Location_ID]=" & "'" & Me![txtLocation_ID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , acDialog
    Update_Loc_Info
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnCoord_Click[Form_frm_Data_Entry])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnClose_CLick
' Description:  Handles close button actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch - June, 2006
' Adapted:      Bonnie Campbell, March 22, 2017 - for NCPN tools
' Revisions:
'   JRB, 6/x/2006  - initial version
'   BLC, 3/22/2017 - added documentation, renamed cmdClose to btnClose
' ---------------------------------
Private Sub btnClose_Click()
    On Error GoTo Err_Handler

    DoCmd.RunCommand acCmdSaveRecord
    DoCmd.Close , , acSaveNo

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnClose_Click[Form_frm_Data_Entry])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnComments_Click
' Description:  Handles comment button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch - June, 2006
' Adapted:      Bonnie Campbell, March 22, 2017 - for NCPN tools
' Revisions:
'   JRB, 6/x/2006  - initial version
'   BLC, 3/22/2017 - added documentation, error handling,
'                    renamed buttonComments to btnComments
' ---------------------------------
Private Sub btnComments_Click()
On Error GoTo Err_Handler

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim RevisitComments As dao.Recordset
    Dim db As dao.Database
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
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnComments_Click[Form_frm_Data_Entry])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnPlotQAQC_Click
' Description:  Handles Plot QA/QC button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch - June, 2006
' Adapted:      Bonnie Campbell, March 22, 2017 - for NCPN tools
' Revisions:
'   JRB, 6/x/2006  - initial version
'   BLC, 3/22/2017 - added documentation, error handling
'   BLC, 3/31/2017 - added sample date open arg for PlotCheck
'   BLC, 8/10/2017 - revised from opening with WindowMode as acDialog to acWindowNormal
'                    to prevent PlotCheck from opening as a modal dialog box w/o controls
'   BLC, 2/1/2018  - revised to run PlotCheckSelect vs PlotCheck form
' ---------------------------------
Private Sub btnPlotQAQC_Click()
On Error GoTo Err_Handler
    
    'minimize calling form
    ToggleForm Me.Name, -1
    
    'commit form changes first
    DoCmd.RunCommand acCmdSaveRecord
    Me.Requery
    Me.Refresh
    
'    'commit form changes first
'    DoCmd.RunCommand acCmdSaveRecord
'
'    'close form & re-open
'    DoCmd.Close acForm, "frm_Data_Entry", acSaveYes
'    DoCmd.OpenForm "frm_Data_Entry", , , TempVars("CriteriaLoc") & " AND " & TempVars("CriteriaEvent"), , , TempVars("CriteriaEvent")
'    DoCmd.SelectObject acForm, "frm_Data_Entry"
'    DoCmd.Minimize

    'pass along form name, plot ID, sample date (WindowMode acWindowNormal vs. acDialog)
    DoCmd.OpenForm "PlotCheckSelect", acNormal, , , , acWindowNormal, Me.Name & _
                                                            "|" & Me.Plot_ID & _
                                                            "|" & Me.txtStart_date

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnPlotQAQC_Click[Form_frm_Data_Entry])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_Close
' Description:  Handles form closing actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch - June, 2006
' Adapted:      Bonnie Campbell, March 22, 2017 - for NCPN tools
' Revisions:
'   JRB, 6/x/2006  - initial version
'   BLC, 3/22/2017 - added documentation, error handling
'   BLC, 3/31/2017 - added clearing for qc queries
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

    If IsLoaded("frm_Data_Gateway") Then
        Forms("frm_Data_Gateway").Requery
    End If
    
    'clear queries
    RemoveTemplateQueries

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[Form_frm_Data_Entry])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Update_Loc_Info
' Description:  Updates associated location information when Location_ID is updated
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   GetCriteriaString
' Source/date:  Simon Kingston - Sept. 2006
' Adapted:      Bonnie Campbell, March 22, 2017 - for NCPN tools
' Revisions:
'   SK, 9/x/2006  - initial version
'   BLC, 3/22/2017 - added documentation, error handling
' ---------------------------------
Public Sub Update_Loc_Info()
On Error GoTo Err_Handler
    
    Dim strXY As Variant
    Dim strCriteria As String
    
    If IsNull(Me!txtLocation_ID) Then
        Me!txtXY = Null
        Me!txtUnit_Code = Null
    Else
        strCriteria = GetCriteriaString("Location_ID=", "tbl_Locations", "Location_ID", Me.Name, "txtLocation_ID")
        
        strXY = "E: " & Nz(DLookup("E_Coord", "tbl_Locations", strCriteria), "")
        strXY = strXY & "  N: " & Nz(DLookup("N_Coord", "tbl_Locations", strCriteria), "")
        Me!txtXY = strXY
        Me!txtUnit_Code = DLookup("Unit_Code", "tbl_Locations", strCriteria)
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[Form_frm_Data_Entry])"
    End Select
    Resume Exit_Handler
End Sub
