Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
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
    Width =13500
    DatasheetFontHeight =9
    ItemSuffix =226
    Left =405
    Top =45
    Right =13650
    Bottom =8280
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x8b370bc14b2ee340
    End
    RecordSource ="qry_Quadrat"
    Caption ="frm_Canopy_Transect"
    OnCurrent ="[Event Procedure]"
    BeforeInsert ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =255
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            TextAlign =2
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Line
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
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11460
                    Top =60
                    Width =630
                    Height =180
                    ColumnWidth =2310
                    Name ="Quadrat_ID"
                    ControlSource ="Quadrat_ID"
                    StatusBarText ="Unique record identifier - primary key"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =12300
                    Top =60
                    Width =630
                    Height =180
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Transect_ID"
                    ControlSource ="Transect_ID"
                    StatusBarText ="M. Link to tbl_Locations (Loc_ID)"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1380
                    Top =60
                    Width =360
                    ColumnWidth =465
                    FontWeight =700
                    TabIndex =2
                    Name ="Quadrat"
                    ControlSource ="Quadrat"
                    StatusBarText ="Transect number - 1, 2, or 3"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =0
                            Left =60
                            Top =60
                            Width =1320
                            Height =240
                            FontWeight =700
                            Name ="Transect_Label"
                            Caption ="10 m2 Quadrat"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =1740
                    Top =60
                    Width =306
                    Height =306
                    TabIndex =5
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
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2100
                    Top =60
                    Width =306
                    Height =306
                    TabIndex =6
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
                    Left =5760
                    Top =60
                    Width =1620
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Observer"
                    ControlSource ="Observer"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Contacts.Contact_ID, tlu_Contacts.Last_Name, tlu_Contacts.First_Name "
                        "FROM tlu_Contacts; "
                    ColumnWidths ="0;810;840"
                    OnKeyDown ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =0
                            Left =4920
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
                    Left =8640
                    Top =60
                    Width =1620
                    TabIndex =4
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Recorder"
                    ControlSource ="Recorder"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Contacts.Contact_ID, tlu_Contacts.Last_Name, tlu_Contacts.First_Name "
                        "FROM tlu_Contacts; "
                    ColumnWidths ="0;750;795"
                    OnKeyDown ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =0
                            Left =7800
                            Top =60
                            Width =840
                            Height =245
                            FontWeight =700
                            Name ="Recorder_Label"
                            Caption ="Recorder"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =600
                    Top =480
                    Width =11760
                    Height =2100
                    TabIndex =7
                    Name ="fsub_Species"
                    SourceObject ="Form.fsub_Species"
                    LinkChildFields ="Quadrat_ID"
                    LinkMasterFields ="Quadrat_ID"

                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    AccessKey =32
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11880
                    Top =3540
                    Width =540
                    TabIndex =84
                    Name ="Cover_Shrub"
                    ControlSource ="Cover_Shrub"
                    StatusBarText ="Percentage shrub & dwarf-shrub cover in 10 m2 quadrat"
                    UnicodeAccessKey =32

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =93
                            TextAlign =0
                            Left =10080
                            Top =3540
                            Width =1800
                            Height =240
                            Name ="Label33"
                            Caption ="shrub & dwarf-shrub"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11880
                    Top =3780
                    Width =540
                    TabIndex =85
                    Name ="Cover_Annual"
                    ControlSource ="Cover_Annual"
                    StatusBarText ="Percentage annual grass cover in 10 m2 quadrat"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =95
                            TextAlign =0
                            Left =10080
                            Top =3780
                            Width =1800
                            Height =240
                            Name ="Label34"
                            Caption ="annual grass"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11880
                    Top =4020
                    Width =540
                    TabIndex =86
                    Name ="Cover_Perennial"
                    ControlSource ="Cover_Perennial"
                    StatusBarText ="Percentage perennial grass cover in 10 m2 quadrat"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =95
                            TextAlign =0
                            Left =10080
                            Top =4020
                            Width =1800
                            Height =240
                            Name ="Label35"
                            Caption ="perennial grass"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11880
                    Top =4260
                    Width =540
                    TabIndex =87
                    Name ="Cover_Forbs"
                    ControlSource ="Cover_Forbs"
                    StatusBarText ="Percentage forbs/herbs cover in 10 m2 quadrat"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =95
                            TextAlign =0
                            Left =10080
                            Top =4260
                            Width =1800
                            Height =240
                            Name ="Label36"
                            Caption ="forbs/herbs"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11880
                    Top =4980
                    Width =540
                    TabIndex =89
                    Name ="Cover_Soil"
                    ControlSource ="Cover_Soil"
                    StatusBarText ="Percentage bare soil (loose) cover in 10 m2 quadrat"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =93
                            TextAlign =0
                            Left =10080
                            Top =4980
                            Width =1800
                            Height =240
                            Name ="Label37"
                            Caption ="bare soil (loose)"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11880
                    Top =5220
                    Width =540
                    TabIndex =90
                    Name ="Cover_Litter"
                    ControlSource ="Cover_Litter"
                    StatusBarText ="Percentage litter cover in 10 m2 quadrat"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =95
                            TextAlign =0
                            Left =10080
                            Top =5220
                            Width =1800
                            Height =240
                            Name ="Label38"
                            Caption ="litter"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11880
                    Top =5460
                    Width =540
                    TabIndex =91
                    Name ="Cover_Woody"
                    ControlSource ="Cover_Woody"
                    StatusBarText ="Percentage woody debris (>2.5cm) cover in 10 m2 quadrat"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =95
                            TextAlign =0
                            Left =10080
                            Top =5460
                            Width =1800
                            Height =240
                            Name ="Label39"
                            Caption ="woody debris (>2.5cm)"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11880
                    Top =5700
                    Width =540
                    TabIndex =92
                    Name ="Cover_Small_Rock"
                    ControlSource ="Cover_Small_Rock"
                    StatusBarText ="Percentage small rock cover (2-20mm) in 10 m2 quadrat"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =95
                            TextAlign =0
                            Left =10080
                            Top =5700
                            Width =1800
                            Height =240
                            Name ="Label40"
                            Caption ="small rock ((2-20mm)"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11880
                    Top =5940
                    Width =540
                    TabIndex =93
                    Name ="Cover_Large_Rock"
                    ControlSource ="Cover_Large_Rock"
                    StatusBarText ="Percentage large rock (>2cm) cover in 10 m2 quadrat"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =95
                            TextAlign =0
                            Left =10080
                            Top =5940
                            Width =1800
                            Height =240
                            Name ="Label41"
                            Caption ="large rock (>2cm)"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11880
                    Top =6180
                    Width =540
                    TabIndex =94
                    Name ="Cover_Bedrock"
                    ControlSource ="Cover_Bedrock"
                    StatusBarText ="Percentage bedrock cover in 10 m2 quadrat"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =95
                            TextAlign =0
                            Left =10080
                            Top =6180
                            Width =1800
                            Height =240
                            Name ="Label42"
                            Caption ="bedrock"
                        End
                    End
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =93
                    TextAlign =0
                    Left =10080
                    Top =3060
                    Width =1800
                    Height =240
                    FontWeight =700
                    Name ="Label43"
                    Caption ="Total live vegetation"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =11880
                    Top =3060
                    Width =540
                    Height =240
                    Name ="Box44"
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =0
                    Left =9900
                    Top =2820
                    Width =2700
                    Height =240
                    FontSize =10
                    FontWeight =700
                    Name ="Label45"
                    Caption ="% Cover in 10 m2 quadrat"
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =0
                    Left =10080
                    Top =4500
                    Width =1800
                    Height =240
                    FontWeight =700
                    Name ="Label46"
                    Caption ="Surface features"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =11880
                    Top =4500
                    Width =540
                    Height =240
                    Name ="Box47"
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2110
                    Top =3360
                    TabIndex =8
                    Name ="Crust_Q1"
                    ControlSource ="Crust_Q1"
                    StatusBarText ="Undifferentiated crust nested quadrat 0.01 m2"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =93
                            TextAlign =0
                            Left =370
                            Top =3300
                            Width =1620
                            Height =240
                            Name ="Label50"
                            Caption ="Crust, undifferentiated"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2470
                    Top =3360
                    TabIndex =9
                    Name ="Crust_Q2"
                    ControlSource ="Crust_Q2"
                    StatusBarText ="Undifferentiated crust nested quadrat 0.1 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2830
                    Top =3360
                    TabIndex =10
                    Name ="Crust_Q3"
                    ControlSource ="Crust_Q3"
                    StatusBarText ="Undifferentiated crust nested quadrat 1 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =3190
                    Top =3360
                    TabIndex =11
                    Name ="Crust_Q4"
                    ControlSource ="Crust_Q4"
                    StatusBarText ="Undifferentiated crust nested quadrat 10 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =223
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3430
                    Top =3300
                    Width =825
                    TabIndex =12
                    Name ="Crust_Cover_1M2"
                    ControlSource ="Crust_Cover_1M2"
                    StatusBarText ="Percent undifferentiated crust cover in percentage 10 m2 quadrat"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =0
                    Left =370
                    Top =3060
                    Width =1620
                    Height =240
                    FontWeight =700
                    Name ="Label55"
                    Caption ="Surface features"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =255
                    Left =1990
                    Top =3300
                    Width =360
                    Height =240
                    Name ="Box56"
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =255
                    Left =1975
                    Top =3060
                    Width =375
                    Height =240
                    FontWeight =700
                    Name ="Label57"
                    Caption =".01"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =255
                    Left =2350
                    Top =3300
                    Width =360
                    Height =240
                    Name ="Box58"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =255
                    Left =2710
                    Top =3300
                    Width =360
                    Height =240
                    Name ="Box59"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =255
                    Left =3070
                    Top =3300
                    Width =360
                    Height =240
                    Name ="Box60"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =255
                    Left =1990
                    Top =3540
                    Width =360
                    Height =240
                    Name ="Box61"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =255
                    Left =2350
                    Top =3540
                    Width =360
                    Height =240
                    Name ="Box62"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =255
                    Left =2710
                    Top =3540
                    Width =360
                    Height =240
                    Name ="Box63"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =255
                    Left =3070
                    Top =3540
                    Width =360
                    Height =240
                    Name ="Box64"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =127
                    Left =1990
                    Top =3780
                    Width =360
                    Height =240
                    Name ="Box68"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =127
                    Left =2350
                    Top =3780
                    Width =360
                    Height =240
                    Name ="Box69"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =127
                    Left =2710
                    Top =3780
                    Width =360
                    Height =240
                    Name ="Box70"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =127
                    Left =3070
                    Top =3780
                    Width =360
                    Height =240
                    Name ="Box71"
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =127
                    Left =2350
                    Top =3060
                    Width =375
                    Height =240
                    FontWeight =700
                    Name ="Label72"
                    Caption ="0.1"
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =255
                    Left =2710
                    Top =3060
                    Width =375
                    Height =240
                    FontWeight =700
                    Name ="Label73"
                    Caption ="1"
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =255
                    Left =3070
                    Top =3060
                    Width =375
                    Height =240
                    FontWeight =700
                    Name ="Label74"
                    Caption ="10"
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =255
                    Left =3430
                    Top =2880
                    Width =825
                    Height =420
                    FontWeight =700
                    Name ="Label75"
                    Caption ="% cover 1 m2"
                End
                Begin Label
                    BorderWidth =1
                    OverlapFlags =119
                    TextAlign =0
                    Left =1990
                    Top =2820
                    Width =1440
                    Height =240
                    FontWeight =700
                    Name ="Label76"
                    Caption ="Nested quadrats"
                End
                Begin Label
                    BorderWidth =1
                    OverlapFlags =87
                    TextAlign =1
                    Left =370
                    Top =2820
                    Width =1440
                    Height =240
                    FontWeight =700
                    Name ="Label77"
                    Caption ="Presence:"
                End
                Begin CheckBox
                    OverlapFlags =255
                    Left =2110
                    Top =3600
                    TabIndex =13
                    Name ="Cyan_Q1"
                    ControlSource ="Cyan_Q1"
                    StatusBarText ="Cyanobacteria nested quadrat 0.01 m2"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =127
                            TextAlign =0
                            Left =370
                            Top =3540
                            Width =1620
                            Height =240
                            Name ="Label78"
                            Caption ="cyanobacteria"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =255
                    Left =2470
                    Top =3600
                    TabIndex =14
                    Name ="Cyan_Q2"
                    ControlSource ="Cyan_Q2"
                    StatusBarText ="Cyanobacteria nested quadrat 0.1 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =255
                    Left =2830
                    Top =3600
                    TabIndex =15
                    Name ="Cyan_Q3"
                    ControlSource ="Cyan_Q3"
                    StatusBarText ="Cyanobacteria nested quadrat 1 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =255
                    Left =3190
                    Top =3600
                    TabIndex =16
                    Name ="Cyan_Q4"
                    ControlSource ="Cyan_Q4"
                    StatusBarText ="Cyanobacteria nested quadrat 10 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =255
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3430
                    Top =3540
                    Width =825
                    TabIndex =17
                    Name ="Cyan_Cover_1M2"
                    ControlSource ="Cyan_Cover_1M2"
                    StatusBarText ="Percent cyanobacteria cover percentage in 10 m2 quadrat"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =127
                    Left =1990
                    Top =4020
                    Width =360
                    Height =240
                    Name ="Box83"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =127
                    Left =2350
                    Top =4020
                    Width =360
                    Height =240
                    Name ="Box84"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =127
                    Left =2710
                    Top =4020
                    Width =360
                    Height =240
                    Name ="Box85"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =127
                    Left =3070
                    Top =4020
                    Width =360
                    Height =240
                    Name ="Box86"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =93
                    Left =1995
                    Top =4680
                    Width =360
                    Height =240
                    Name ="Box87"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =2355
                    Top =4680
                    Width =360
                    Height =240
                    Name ="Box88"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =2715
                    Top =4680
                    Width =360
                    Height =240
                    Name ="Box89"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =3075
                    Top =4680
                    Width =360
                    Height =240
                    Name ="Box90"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =1995
                    Top =4920
                    Width =360
                    Height =240
                    Name ="Box91"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =2355
                    Top =4920
                    Width =360
                    Height =240
                    Name ="Box92"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =2715
                    Top =4920
                    Width =360
                    Height =240
                    Name ="Box93"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =3075
                    Top =4920
                    Width =360
                    Height =240
                    Name ="Box94"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =93
                    Left =6840
                    Top =3300
                    Width =360
                    Height =240
                    Name ="Box99"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =7200
                    Top =3300
                    Width =360
                    Height =240
                    Name ="Box100"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =7560
                    Top =3300
                    Width =360
                    Height =240
                    Name ="Box101"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =7920
                    Top =3300
                    Width =360
                    Height =240
                    Name ="Box102"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =6840
                    Top =3540
                    Width =360
                    Height =240
                    Name ="Box103"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =7200
                    Top =3540
                    Width =360
                    Height =240
                    Name ="Box104"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =7560
                    Top =3540
                    Width =360
                    Height =240
                    Name ="Box105"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =7920
                    Top =3540
                    Width =360
                    Height =240
                    Name ="Box106"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =6840
                    Top =3780
                    Width =360
                    Height =240
                    Name ="Box107"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =7200
                    Top =3780
                    Width =360
                    Height =240
                    Name ="Box108"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =7560
                    Top =3780
                    Width =360
                    Height =240
                    Name ="Box109"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =7920
                    Top =3780
                    Width =360
                    Height =240
                    Name ="Box110"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =6840
                    Top =4020
                    Width =360
                    Height =240
                    Name ="Box111"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =7200
                    Top =4020
                    Width =360
                    Height =240
                    Name ="Box112"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =7560
                    Top =4020
                    Width =360
                    Height =240
                    Name ="Box113"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =7920
                    Top =4020
                    Width =360
                    Height =240
                    Name ="Box114"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =6840
                    Top =4260
                    Width =360
                    Height =240
                    Name ="Box115"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =7200
                    Top =4260
                    Width =360
                    Height =240
                    Name ="Box116"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =7560
                    Top =4260
                    Width =360
                    Height =240
                    Name ="Box117"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =7920
                    Top =4260
                    Width =360
                    Height =240
                    Name ="Box118"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =6840
                    Top =4500
                    Width =360
                    Height =240
                    Name ="Box119"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =7200
                    Top =4500
                    Width =360
                    Height =240
                    Name ="Box120"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =7560
                    Top =4500
                    Width =360
                    Height =240
                    Name ="Box121"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =7920
                    Top =4500
                    Width =360
                    Height =240
                    Name ="Box122"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =6840
                    Top =4740
                    Width =360
                    Height =240
                    Name ="Box123"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =7200
                    Top =4740
                    Width =360
                    Height =240
                    Name ="Box124"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =7560
                    Top =4740
                    Width =360
                    Height =240
                    Name ="Box125"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =7920
                    Top =4740
                    Width =360
                    Height =240
                    Name ="Box126"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =6840
                    Top =4980
                    Width =360
                    Height =240
                    Name ="Box127"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =7200
                    Top =4980
                    Width =360
                    Height =240
                    Name ="Box128"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =7560
                    Top =4980
                    Width =360
                    Height =240
                    Name ="Box129"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =7920
                    Top =4980
                    Width =360
                    Height =240
                    Name ="Box130"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =6840
                    Top =5220
                    Width =360
                    Height =240
                    Name ="Box131"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =7200
                    Top =5220
                    Width =360
                    Height =240
                    Name ="Box132"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =7560
                    Top =5220
                    Width =360
                    Height =240
                    Name ="Box133"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    Left =7920
                    Top =5220
                    Width =360
                    Height =240
                    Name ="Box134"
                End
                Begin CheckBox
                    OverlapFlags =255
                    Left =2125
                    Top =3840
                    TabIndex =18
                    Name ="Lichen_Q1"
                    ControlSource ="Lichen_Q1"
                    StatusBarText ="Lichen nested quadrat 0.01 m2"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =127
                            TextAlign =0
                            Left =375
                            Top =3780
                            Width =1620
                            Height =240
                            Name ="Label135"
                            Caption ="Lichen"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =255
                    Left =2470
                    Top =3840
                    TabIndex =19
                    Name ="Lichen_Q2"
                    ControlSource ="Lichen_Q2"
                    StatusBarText ="Lichen nested quadrat 0.1 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =255
                    Left =2830
                    Top =3840
                    TabIndex =20
                    Name ="Lichen_Q3"
                    ControlSource ="Lichen_Q3"
                    StatusBarText ="Lichen nested quadrat 1 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =255
                    Left =3190
                    Top =3840
                    TabIndex =21
                    Name ="Lichen_Q4"
                    ControlSource ="Lichen_Q4"
                    StatusBarText ="Lichen nested quadrat 10 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =255
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3430
                    Top =3780
                    Width =825
                    TabIndex =22
                    Name ="Lichen_Cover_1M2"
                    ControlSource ="Lichen_Cover_1M2"
                    StatusBarText ="Percent lichen cover in percentage 10 m2 quadrat"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =93
                    TextAlign =0
                    Left =5220
                    Top =3060
                    Width =1605
                    Height =240
                    FontWeight =700
                    Name ="Label140"
                    Caption ="Disturbance"
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =95
                    Left =6825
                    Top =3060
                    Width =375
                    Height =240
                    FontWeight =700
                    Name ="Label141"
                    Caption =".01"
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =95
                    Left =7200
                    Top =3060
                    Width =375
                    Height =240
                    FontWeight =700
                    Name ="Label142"
                    Caption ="0.1"
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =223
                    Left =7560
                    Top =3060
                    Width =375
                    Height =240
                    FontWeight =700
                    Name ="Label143"
                    Caption ="1"
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =223
                    Left =7920
                    Top =3060
                    Width =375
                    Height =240
                    FontWeight =700
                    Name ="Label144"
                    Caption ="10"
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =223
                    Left =8280
                    Top =2880
                    Width =840
                    Height =420
                    FontWeight =700
                    Name ="Label145"
                    Caption ="% cover 10 m2"
                End
                Begin Label
                    BorderWidth =1
                    OverlapFlags =87
                    TextAlign =0
                    Left =6840
                    Top =2820
                    Width =1440
                    Height =240
                    FontWeight =700
                    Name ="Label146"
                    Caption ="Nested quadrats"
                End
                Begin Label
                    BorderWidth =1
                    OverlapFlags =87
                    TextAlign =1
                    Left =5220
                    Top =2820
                    Width =1440
                    Height =240
                    FontWeight =700
                    Name ="Label147"
                    Caption ="Presence:"
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =2115
                    Top =4080
                    TabIndex =23
                    Name ="Moss_Q1"
                    ControlSource ="Moss_Q1"
                    StatusBarText ="Moss nested quadrat 0.01 m2"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =127
                            TextAlign =0
                            Left =380
                            Top =4020
                            Width =1620
                            Height =240
                            Name ="Label148"
                            Caption ="Moss"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =2460
                    Top =4080
                    TabIndex =24
                    Name ="Moss_Q2"
                    ControlSource ="Moss_Q2"
                    StatusBarText ="Moss nested quadrat 0.1 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =2820
                    Top =4080
                    TabIndex =25
                    Name ="Moss_Q3"
                    ControlSource ="Moss_Q3"
                    StatusBarText ="Moss nested quadrat 1 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =255
                    Left =3180
                    Top =4080
                    TabIndex =26
                    Name ="Moss_Q4"
                    ControlSource ="Moss_Q4"
                    StatusBarText ="Moss nested quadrat 10 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =255
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3435
                    Top =4020
                    Width =825
                    TabIndex =27
                    Name ="Moss_Cover_1M2"
                    ControlSource ="Moss_Cover_1M2"
                    StatusBarText ="Percent moss cover percentage in 10 m2 quadrat"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =2195
                    Top =4740
                    TabIndex =28
                    Name ="Lscat_Q1"
                    ControlSource ="Lscat_Q1"
                    StatusBarText ="Livestock scat nested quadrat 0.01 m2"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =95
                            TextAlign =0
                            Left =375
                            Top =4680
                            Width =1620
                            Height =240
                            Name ="Label154"
                            Caption ="scat, livestock"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =2455
                    Top =4740
                    TabIndex =29
                    Name ="Lscat_Q2"
                    ControlSource ="Lscat_Q2"
                    StatusBarText ="Livestock scat nested quadrat 0.1 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =2815
                    Top =4740
                    TabIndex =30
                    Name ="Lscat_Q3"
                    ControlSource ="Lscat_Q3"
                    StatusBarText ="Livestock scat nested quadrat 1 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =3175
                    Top =4740
                    TabIndex =31
                    Name ="Lscat_Q4"
                    ControlSource ="Lscat_Q4"
                    StatusBarText ="Livestock scat nested quadrat 10 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3435
                    Top =4680
                    Width =825
                    TabIndex =32
                    Name ="Lscat_Cover"
                    ControlSource ="Lscat_Cover"
                    StatusBarText ="Percent livestock scat cover percentage in 10 m2 quadrat"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =2195
                    Top =4980
                    TabIndex =33
                    Name ="Wscat_Q1"
                    ControlSource ="Wscat_Q1"
                    StatusBarText ="Wildlife scat nested quadrat 0.01 m2"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =87
                            TextAlign =0
                            Left =375
                            Top =4920
                            Width =1620
                            Height =240
                            Name ="Label159"
                            Caption ="scat, wildlife"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =2455
                    Top =4980
                    TabIndex =34
                    Name ="Wscat_Q2"
                    ControlSource ="Wscat_Q2"
                    StatusBarText ="Wildlife scat nested quadrat 0.1 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =2815
                    Top =4980
                    TabIndex =35
                    Name ="Wscat_Q3"
                    ControlSource ="Wscat_Q3"
                    StatusBarText ="Wildlife scat nested quadrat 1 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =3175
                    Top =4980
                    TabIndex =36
                    Name ="Wscat_Q4"
                    ControlSource ="Wscat_Q4"
                    StatusBarText ="Wildlife scat nested quadrat 10 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3435
                    Top =4920
                    Width =825
                    TabIndex =37
                    Name ="Wscat_Cover"
                    ControlSource ="Wscat_Cover"
                    StatusBarText ="Percent wildlife scat cover percentage in 10 m2 quadrat"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =6960
                    Top =3360
                    TabIndex =38
                    Name ="Ant_Q1"
                    ControlSource ="Ant_Q1"
                    StatusBarText ="Ant mound nested quadrat 0.01 m2"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =95
                            TextAlign =0
                            Left =5220
                            Top =3300
                            Width =1605
                            Height =240
                            Name ="Label164"
                            Caption ="ant mound"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =7320
                    Top =3360
                    TabIndex =39
                    Name ="Ant_Q2"
                    ControlSource ="Ant_Q2"
                    StatusBarText ="Ant mound nested quadrat 0.1 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =7680
                    Top =3360
                    TabIndex =40
                    Name ="Ant_Q3"
                    ControlSource ="Ant_Q3"
                    StatusBarText ="Ant mound nested quadrat 1 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =8040
                    Top =3360
                    TabIndex =41
                    Name ="Ant_Q4"
                    ControlSource ="Ant_Q4"
                    StatusBarText ="Ant mound nested quadrat 10 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =223
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8280
                    Top =3300
                    Width =840
                    TabIndex =42
                    Name ="Ant_Cover"
                    ControlSource ="Ant_Cover"
                    StatusBarText ="Percent ant mound cover percentage in 10 m2 quadrat"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =6960
                    Top =3600
                    TabIndex =43
                    Name ="Bicycle_Q1"
                    ControlSource ="Bicycle_Q1"
                    StatusBarText ="Bicycle disturbance nested quadrat 0.01 m2"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =95
                            TextAlign =0
                            Left =5220
                            Top =3540
                            Width =1605
                            Height =240
                            Name ="Label169"
                            Caption ="bicycle"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =7320
                    Top =3600
                    TabIndex =44
                    Name ="Bicycle_Q2"
                    ControlSource ="Bicycle_Q2"
                    StatusBarText ="Bicycle disturbance nested quadrat 0.1 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =7680
                    Top =3600
                    TabIndex =45
                    Name ="Bicycle_Q3"
                    ControlSource ="Bicycle_Q3"
                    StatusBarText ="Bicycle disturbance nested quadrat 1 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =8040
                    Top =3600
                    TabIndex =46
                    Name ="Bicycle_Q4"
                    ControlSource ="Bicycle_Q4"
                    StatusBarText ="Bicycle disturbance nested quadrat 10 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =223
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8280
                    Top =3540
                    Width =840
                    TabIndex =47
                    Name ="Bicycle_Cover"
                    ControlSource ="Bicycle_Cover"
                    StatusBarText ="Percent bicycle disturbance cover percentage in 10 m2 quadrat"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =6960
                    Top =3840
                    TabIndex =48
                    Name ="Htrail_Q1"
                    ControlSource ="Htrail_Q1"
                    StatusBarText ="Human track/trail nested quadrat 0.01 m2"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =95
                            TextAlign =0
                            Left =5220
                            Top =3780
                            Width =1605
                            Height =240
                            Name ="Label174"
                            Caption ="human track/trail"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =7320
                    Top =3840
                    TabIndex =49
                    Name ="Htrail_Q2"
                    ControlSource ="Htrail_Q2"
                    StatusBarText ="Human track/trail nested quadrat 0.1 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =7680
                    Top =3840
                    TabIndex =50
                    Name ="Htrail_Q3"
                    ControlSource ="Htrail_Q3"
                    StatusBarText ="Human track/trail nested quadrat 1 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =8040
                    Top =3840
                    TabIndex =51
                    Name ="Htrail_Q4"
                    ControlSource ="Htrail_Q4"
                    StatusBarText ="Human track/trail nested quadrat 10 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =223
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8280
                    Top =3780
                    Width =840
                    TabIndex =52
                    Name ="Htrail_Cover"
                    ControlSource ="Htrail_Cover"
                    StatusBarText ="Human track/trail disturbance cover percentage in 10 m2 quadrat"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =6960
                    Top =4080
                    TabIndex =53
                    Name ="Ltrail_Q1"
                    ControlSource ="Ltrail_Q1"
                    StatusBarText ="Livestock track/trail nested quadrat 0.01 m2"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =95
                            TextAlign =0
                            Left =5220
                            Top =4020
                            Width =1605
                            Height =240
                            Name ="Label179"
                            Caption ="livestock track/trail"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =7320
                    Top =4080
                    TabIndex =54
                    Name ="Ltrail_Q2"
                    ControlSource ="Ltrail_Q2"
                    StatusBarText ="Livestock track/trail nested quadrat 0.1 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =7680
                    Top =4080
                    TabIndex =55
                    Name ="Ltrail_Q3"
                    ControlSource ="Ltrail_Q3"
                    StatusBarText ="Livestock track/trail nested quadrat 1 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =8040
                    Top =4080
                    TabIndex =56
                    Name ="Ltrail_Q4"
                    ControlSource ="Ltrail_Q4"
                    StatusBarText ="Livestock track/trail nested quadrat 10 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =223
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8280
                    Top =4020
                    Width =840
                    TabIndex =57
                    Name ="Ltrail_Cover"
                    ControlSource ="Ltrail_Cover"
                    StatusBarText ="Livestock track/trail disturbance cover percentage in 10 m2 quadrat"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =6960
                    Top =4320
                    TabIndex =58
                    Name ="Vehicle_Q1"
                    ControlSource ="Vehicle_Q1"
                    StatusBarText ="Motor vehicle track nested quadrat 0.01 m2"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =95
                            TextAlign =0
                            Left =5220
                            Top =4260
                            Width =1605
                            Height =240
                            Name ="Label184"
                            Caption ="motor vehicle"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =7320
                    Top =4320
                    TabIndex =59
                    Name ="Vehicle_Q2"
                    ControlSource ="Vehicle_Q2"
                    StatusBarText ="Motor vehicle track nested quadrat 0.1 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =7680
                    Top =4320
                    TabIndex =60
                    Name ="Vehicle_Q3"
                    ControlSource ="Vehicle_Q3"
                    StatusBarText ="Motor vehicle track nested quadrat 1 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =8040
                    Top =4320
                    TabIndex =61
                    Name ="Vehicle_Q4"
                    ControlSource ="Vehicle_Q4"
                    StatusBarText ="Motor vehicle track nested quadrat 10 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =223
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8280
                    Top =4260
                    Width =840
                    TabIndex =62
                    Name ="Vehicle_Cover"
                    ControlSource ="Vehicle_Cover"
                    StatusBarText ="Motor vehicle disturbance cover percentage in 10 m2 quadrat"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =6960
                    Top =4560
                    TabIndex =63
                    Name ="Wildlife_Ex_Q1"
                    ControlSource ="Wildlife_Ex_Q1"
                    StatusBarText ="Wildlife excavation nested quadrat 0.01 m2"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =95
                            TextAlign =0
                            Left =5220
                            Top =4500
                            Width =1605
                            Height =240
                            Name ="Label189"
                            Caption ="wildlife excavation"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =7320
                    Top =4560
                    TabIndex =64
                    Name ="Wildlife_Ex_Q2"
                    ControlSource ="Wildlife_Ex_Q2"
                    StatusBarText ="Wildlife excavation nested quadrat 0.1 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =7680
                    Top =4560
                    TabIndex =65
                    Name ="Wildlife_Ex_Q3"
                    ControlSource ="Wildlife_Ex_Q3"
                    StatusBarText ="Wildlife excavation nested quadrat 1 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =8040
                    Top =4560
                    TabIndex =66
                    Name ="Wildlife_Ex_Q4"
                    ControlSource ="Wildlife_Ex_Q4"
                    StatusBarText ="Wildlife excavation nested quadrat 10 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =223
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8280
                    Top =4500
                    Width =840
                    TabIndex =67
                    Name ="Wildlife_Ex_Cover"
                    ControlSource ="Wildlife_Ex_Cover"
                    StatusBarText ="Wildlife excavation disturbance cover percentage in 10 m2 quadrat"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =6960
                    Top =4800
                    TabIndex =68
                    Name ="Wtrail_Q1"
                    ControlSource ="Wtrail_Q1"
                    StatusBarText ="Wildlife track/trail nested quadrat 0.01 m2"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =95
                            TextAlign =0
                            Left =5220
                            Top =4740
                            Width =1605
                            Height =240
                            Name ="Label194"
                            Caption ="wildlife track"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =7320
                    Top =4800
                    TabIndex =69
                    Name ="Wtrail_Q2"
                    ControlSource ="Wtrail_Q2"
                    StatusBarText ="Wildlife track/trail nested quadrat 0.1 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =7680
                    Top =4800
                    TabIndex =70
                    Name ="Wtrail_Q3"
                    ControlSource ="Wtrail_Q3"
                    StatusBarText ="Wildlife track/trail nested quadrat 1 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =8040
                    Top =4800
                    TabIndex =71
                    Name ="Wtrail_Q4"
                    ControlSource ="Wtrail_Q4"
                    StatusBarText ="Wildlife track/trail nested quadrat 10 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =223
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8280
                    Top =4740
                    Width =840
                    TabIndex =72
                    Name ="Wtrail_Cover"
                    ControlSource ="Wtrail_Cover"
                    StatusBarText ="Wildlife track/trail disturbance cover percentage in 10 m2 quadrat"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =6960
                    Top =5040
                    TabIndex =73
                    Name ="Other_Q1"
                    ControlSource ="Other_Q1"
                    StatusBarText ="Other anthropogenic nested quadrat 0.01 m2"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =95
                            TextAlign =0
                            Left =5220
                            Top =4980
                            Width =1605
                            Height =240
                            Name ="Label199"
                            Caption ="other anthropogenic"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =7320
                    Top =5040
                    TabIndex =74
                    Name ="Other_Q2"
                    ControlSource ="Other_Q2"
                    StatusBarText ="Other anthropogenic nested quadrat 0.1 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =7680
                    Top =5040
                    TabIndex =75
                    Name ="Other_Q3"
                    ControlSource ="Other_Q3"
                    StatusBarText ="Other anthropogenic nested quadrat 1 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =8040
                    Top =5040
                    TabIndex =76
                    Name ="Other_Q4"
                    ControlSource ="Other_Q4"
                    StatusBarText ="Other anthropogenic nested quadrat 10 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =223
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8280
                    Top =4980
                    Width =840
                    TabIndex =77
                    Name ="Other_Cover"
                    ControlSource ="Other_Cover"
                    StatusBarText ="Other anthropogenic disturbance coverpercentage  in 10 m2 quadrat"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =6960
                    Top =5280
                    TabIndex =78
                    Name ="Undiff_Q1"
                    ControlSource ="Undiff_Q1"
                    StatusBarText ="Undifferentiated disturbance nested quadrat 0.01 m2"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =87
                            TextAlign =0
                            Left =5220
                            Top =5220
                            Width =1605
                            Height =240
                            Name ="Label204"
                            Caption ="undifferentiated"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =7320
                    Top =5280
                    TabIndex =79
                    Name ="Undiff_Q2"
                    ControlSource ="Undiff_Q2"
                    StatusBarText ="Undifferentiated disturbance nested quadrat 0.1 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =7680
                    Top =5280
                    TabIndex =80
                    Name ="Undiff_Q3"
                    ControlSource ="Undiff_Q3"
                    StatusBarText ="Undifferentiated disturbance nested quadrat 1 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =8040
                    Top =5280
                    TabIndex =81
                    Name ="Undiff_Q4"
                    ControlSource ="Undiff_Q4"
                    StatusBarText ="Undifferentiated disturbance nested quadrat 10 m2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8280
                    Top =5220
                    Width =840
                    TabIndex =82
                    Name ="Undiff_Cover"
                    ControlSource ="Undiff_Cover"
                    StatusBarText ="Undifferentiated disturbance cover percentage in 10 m2 quadrat"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =10080
                    Top =7500
                    Width =2700
                    Height =720
                    TabIndex =97
                    Name ="Comments"
                    ControlSource ="Comments"
                    StatusBarText ="10 m2 quadrat comments"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =0
                            Left =10080
                            Top =7260
                            Width =1020
                            Height =240
                            FontWeight =700
                            Name ="Label209"
                            Caption ="Comments:"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =87
                    Left =360
                    Top =5940
                    Width =9240
                    Height =2265
                    TabIndex =98
                    Name ="fsub_Quadrat_Shrubs"
                    SourceObject ="Form.fsub_Quadrat_Shrubs"
                    LinkChildFields ="Quadrat_ID"
                    LinkMasterFields ="Quadrat_ID"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =0
                            Left =360
                            Top =5640
                            Width =5730
                            Height =300
                            FontSize =10
                            FontWeight =700
                            Name ="fsub_Quadrat_Shrubs Label"
                            Caption ="Number of Live Shrubs/Trees Rooted in 10 m2 Quadrat"
                            EventProcPrefix ="fsub_Quadrat_Shrubs_Label"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11880
                    Top =6900
                    Width =540
                    TabIndex =99
                    Name ="Total_Cover_Percent"

                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =1
                    Left =10080
                    Top =6900
                    Width =1800
                    Height =240
                    FontWeight =700
                    Name ="Label214"
                    Caption ="Total Cover Percent"
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =119
                    Left =3435
                    Top =4260
                    Width =825
                    Height =420
                    FontWeight =700
                    Name ="Label215"
                    Caption ="% cover 10 m2"
                End
                Begin Line
                    BorderWidth =1
                    OverlapFlags =119
                    Left =375
                    Top =4260
                    Width =0
                    Height =420
                    Name ="Line216"
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11880
                    Top =6420
                    Width =540
                    TabIndex =95
                    Name ="Cover_Biocrust"
                    ControlSource ="Cover_Biocrust"
                    StatusBarText ="Percentage biocrust cover in 10 m2 quadrat"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =0
                            Left =10080
                            Top =6420
                            Width =1800
                            Height =240
                            Name ="Label218"
                            Caption ="biocrust"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11880
                    Top =6660
                    Width =540
                    TabIndex =96
                    Name ="Cover_Crust_Undiff"
                    ControlSource ="Cover_Crust_Undiff"
                    StatusBarText ="Percentage undifferentiated crust cover in 10 m2 quadrat"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =0
                            Left =10080
                            Top =6660
                            Width =1800
                            Height =240
                            Name ="Label219"
                            Caption ="crust, undifferentiated"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11880
                    Top =3300
                    Width =540
                    TabIndex =83
                    Name ="Cover_Tree"
                    ControlSource ="Cover_Tree"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =0
                            Left =10080
                            Top =3300
                            Width =1800
                            Height =240
                            Name ="Label223"
                            Caption ="tree"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11880
                    Top =4740
                    Width =540
                    TabIndex =88
                    Name ="Cover_Plant_Basal"
                    ControlSource ="Cover_Plant_Basal"
                    StatusBarText ="Percentage undifferentiated crust cover in 10 m2 quadrat"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =0
                            Left =10080
                            Top =4740
                            Width =1800
                            Height =240
                            Name ="Label225"
                            Caption ="plant basal cover"
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
Public Function Calc1MCover() As Integer
' Calculate total cover for quality control - 4/01/2007 - Russ DenBleyker
' Northern Colorado Plateau Network
    On Error GoTo Err_Handler
    
    Dim Cover1M As Integer
   
    Cover1M = 0

    If Not IsNull(Me!Crust_Cover_1M2) And Me!Crust_Cover_1M2 <> "T" Then
      Cover1M = Cover1M + Me!Crust_Cover_1M2
    End If
    If Not IsNull(Me!Cyan_Cover_1M2) And Me!Cyan_Cover_1M2 <> "T" Then
      Cover1M = Cover1M + Me!Cyan_Cover_1M2
    End If
    If Not IsNull(Me!Lichen_Cover_1M2) And Me!Lichen_Cover_1M2 <> "T" Then
      Cover1M = Cover1M + Me!Lichen_Cover_1M2
    End If
    If Not IsNull(Me!Moss_Cover_1M2) And Me!Moss_Cover_1M2 <> "T" Then
      Cover1M = Cover1M + Me!Moss_Cover_1M2
    End If

    Calc1MCover = Cover1M
Exit_Procedure_1M:
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (Calc1MCover)"
            Resume Exit_Procedure_1M
    End Select

End Function
Public Function CalcTotalCover() As Integer
' Calculate total cover for quality control - 4/01/2007 - Russ DenBleyker
' Northern Colorado Plateau Network
    On Error GoTo Err_Handler
    
    Dim TotalCover As Integer
   
    TotalCover = 0

    If Not IsNull(Me!Lscat_Cover) And Me!Lscat_Cover <> "T" Then
      TotalCover = TotalCover + Me!Lscat_Cover
    End If
    If Not IsNull(Me!Wscat_Cover) And Me!Wscat_Cover <> "T" Then
      TotalCover = TotalCover + Me!Wscat_Cover
    End If
    If Not IsNull(Me!Ant_Cover) And Me!Ant_Cover <> "T" Then
      TotalCover = TotalCover + Me!Ant_Cover
    End If
    If Not IsNull(Me!Bicycle_Cover) And Me!Bicycle_Cover <> "T" Then
      TotalCover = TotalCover + Me!Bicycle_Cover
    End If
    If Not IsNull(Me!Htrail_Cover) And Me!Htrail_Cover <> "T" Then
      TotalCover = TotalCover + Me!Htrail_Cover
    End If
    If Not IsNull(Me!Ltrail_Cover) And Me!Ltrail_Cover <> "T" Then
      TotalCover = TotalCover + Me!Ltrail_Cover
    End If
    If Not IsNull(Me!Vehicle_Cover) And Me!Vehicle_Cover <> "T" Then
      TotalCover = TotalCover + Me!Vehicle_Cover
    End If
    If Not IsNull(Me!Wildlife_Ex_Cover) And Me!Wildlife_Ex_Cover <> "T" Then
      TotalCover = TotalCover + Me!Wildlife_Ex_Cover
    End If
    If Not IsNull(Me!Wtrail_Cover) And Me!Wtrail_Cover <> "T" Then
      TotalCover = TotalCover + Me!Wtrail_Cover
    End If
    If Not IsNull(Me!Other_Cover) And Me!Other_Cover <> "T" Then
      TotalCover = TotalCover + Me!Other_Cover
    End If
    If Not IsNull(Me!Undiff_Cover) And Me!Undiff_Cover <> "T" Then
      TotalCover = TotalCover + Me!Undiff_Cover
    End If
    If Not IsNull(Me!Cover_Plant_Basal) And Me!Cover_Plant_Basal <> "T" Then
      TotalCover = TotalCover + Me!Cover_Plant_Basal
    End If
    If Not IsNull(Me!Cover_Soil) And Me!Cover_Soil <> "T" Then
      TotalCover = TotalCover + Me!Cover_Soil
    End If
    If Not IsNull(Me!Cover_Litter) And Me!Cover_Litter <> "T" Then
      TotalCover = TotalCover + Me!Cover_Litter
    End If
    If Not IsNull(Me!Cover_Woody) And Me!Cover_Woody <> "T" Then
      TotalCover = TotalCover + Me!Cover_Woody
    End If
    If Not IsNull(Me!Cover_Small_Rock) And Me!Cover_Small_Rock <> "T" Then
      TotalCover = TotalCover + Me!Cover_Small_Rock
    End If
    If Not IsNull(Me!Cover_Large_Rock) And Me!Cover_Large_Rock <> "T" Then
      TotalCover = TotalCover + Me!Cover_Large_Rock
    End If
    If Not IsNull(Me!Cover_Bedrock) And Me!Cover_Bedrock <> "T" Then
      TotalCover = TotalCover + Me!Cover_Bedrock
    End If
    If Not IsNull(Me!Cover_Biocrust) And Me!Cover_Biocrust <> "T" Then
      TotalCover = TotalCover + Me!Cover_Biocrust
    End If
    If Not IsNull(Me!Cover_Crust_Undiff) And Me!Cover_Crust_Undiff <> "T" Then
      TotalCover = TotalCover + Me!Cover_Crust_Undiff
    End If
    CalcTotalCover = TotalCover
Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (CalcTotalCover)"
            Resume Exit_Procedure
    End Select

End Function

Private Sub Ant_Cover_AfterUpdate()
  Me!Total_Cover_Percent = CalcTotalCover()

End Sub

Private Sub Ant_Q1_AfterUpdate()
  If Me!Ant_Q1 = -1 Then
    Me!Ant_Q2 = 0
    Me!Ant_Q3 = 0
    Me!Ant_Q4 = 0
  End If
End Sub

Private Sub Ant_Q2_AfterUpdate()
  If Me!Ant_Q2 = -1 Then
    Me!Ant_Q1 = 0
    Me!Ant_Q3 = 0
    Me!Ant_Q4 = 0
  End If
End Sub

Private Sub Ant_Q3_AfterUpdate()
  If Me!Ant_Q3 = -1 Then
    Me!Ant_Q2 = 0
    Me!Ant_Q1 = 0
    Me!Ant_Q4 = 0
  End If
End Sub

Private Sub Ant_Q4_AfterUpdate()
  If Me!Ant_Q4 = -1 Then
    Me!Ant_Q2 = 0
    Me!Ant_Q3 = 0
    Me!Ant_Q1 = 0
  End If
End Sub

Private Sub Bicycle_Cover_AfterUpdate()
  Me!Total_Cover_Percent = CalcTotalCover()

End Sub

Private Sub Bicycle_Q1_AfterUpdate()
  If Me!Bicycle_Q1 = -1 Then
    Me!Bicycle_Q2 = 0
    Me!Bicycle_Q3 = 0
    Me!Bicycle_Q4 = 0
  End If
End Sub

Private Sub Bicycle_Q2_AfterUpdate()
  If Me!Bicycle_Q2 = -1 Then
    Me!Bicycle_Q1 = 0
    Me!Bicycle_Q3 = 0
    Me!Bicycle_Q4 = 0
  End If
End Sub

Private Sub Bicycle_Q3_AfterUpdate()
  If Me!Bicycle_Q3 = -1 Then
    Me!Bicycle_Q2 = 0
    Me!Bicycle_Q1 = 0
    Me!Bicycle_Q4 = 0
  End If
End Sub

Private Sub Bicycle_Q4_AfterUpdate()
  If Me!Bicycle_Q4 = -1 Then
    Me!Bicycle_Q2 = 0
    Me!Bicycle_Q3 = 0
    Me!Bicycle_Q1 = 0
  End If
End Sub

Private Sub ButtonNext_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub ButtonPrevious_Click()
On Error GoTo Err_ButtonPrevious_Click
  Dim intQuadrat As Byte

  If Me!Quadrat = 1 Then
    MsgBox "Already on first Quadrat"
  Else
    intQuadrat = Me!Quadrat
    DoCmd.GoToRecord , , acPrevious
    Me!Quadrat = intQuadrat - 1
  End If
  
Exit_ButtonPrevious_Click:
    Exit Sub

Err_ButtonPrevious_Click:
    MsgBox Err.Description
    Resume Exit_ButtonPrevious_Click
    
End Sub
Private Sub ButtonNext_Click()
On Error GoTo Err_ButtonNext_Click

  Dim intQuadrat As Byte

  If Me!Quadrat = 5 Then
    MsgBox "Five Quadrats maximum!"
  Else
    intQuadrat = Me!Quadrat
    DoCmd.GoToRecord , , acNext
    Me!Quadrat = intQuadrat + 1
  End If

Exit_ButtonNext_Click:
    Exit Sub

Err_ButtonNext_Click:
    MsgBox Err.Description
    Resume Exit_ButtonNext_Click
    
End Sub

Private Sub ButtonPrevious_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Cover_Bedrock_AfterUpdate()
  Me!Total_Cover_Percent = CalcTotalCover()

End Sub

Private Sub Cover_Biocrust_AfterUpdate()
  Me!Total_Cover_Percent = CalcTotalCover()

End Sub

Private Sub Cover_Crust_Undiff_AfterUpdate()
  Me!Total_Cover_Percent = CalcTotalCover()

End Sub

Private Sub Cover_Large_Rock_AfterUpdate()
  Me!Total_Cover_Percent = CalcTotalCover()

End Sub

Private Sub Cover_Litter_AfterUpdate()
  Me!Total_Cover_Percent = CalcTotalCover()

End Sub

Private Sub Cover_Plant_Basal_AfterUpdate()
  Me!Total_Cover_Percent = CalcTotalCover()
End Sub

Private Sub Cover_Small_Rock_AfterUpdate()
  Me!Total_Cover_Percent = CalcTotalCover()

End Sub

Private Sub Cover_Soil_AfterUpdate()
  Me!Total_Cover_Percent = CalcTotalCover()

End Sub

Private Sub Cover_Woody_AfterUpdate()
  Me!Total_Cover_Percent = CalcTotalCover()

End Sub

Private Sub Crust_Cover_1M2_AfterUpdate()
  If Calc1MCover() > 100 Then
    MsgBox "1 M2 cover cannot exceed 100 percent."
    DoCmd.CancelEvent
    SendKeys "{ESC}"
  End If
End Sub

Private Sub Crust_Q1_AfterUpdate()
  If Me!Crust_Q1 = -1 Then
    Me!Crust_Q2 = 0
    Me!Crust_Q3 = 0
    Me!Crust_Q4 = 0
  End If
End Sub

Private Sub Crust_Q2_AfterUpdate()
  If Me!Crust_Q2 = -1 Then
    Me!Crust_Q1 = 0
    Me!Crust_Q3 = 0
    Me!Crust_Q4 = 0
  End If
End Sub

Private Sub Crust_Q3_AfterUpdate()
  If Me!Crust_Q3 = -1 Then
    Me!Crust_Q2 = 0
    Me!Crust_Q1 = 0
    Me!Crust_Q4 = 0
  End If
End Sub

Private Sub Crust_Q4_AfterUpdate()
  If Me!Crust_Q4 = -1 Then
    Me!Crust_Q2 = 0
    Me!Crust_Q3 = 0
    Me!Crust_Q1 = 0
  End If
End Sub

Private Sub Cyan_Cover_1M2_AfterUpdate()
  If Calc1MCover() > 100 Then
    MsgBox "1 M2 cover cannot exceed 100 percent."
    DoCmd.CancelEvent
    SendKeys "{ESC}"
  End If
End Sub

Private Sub Cyan_Q1_AfterUpdate()
  If Me!Cyan_Q1 = -1 Then
    Me!Cyan_Q2 = 0
    Me!Cyan_Q3 = 0
    Me!Cyan_Q4 = 0
  End If
End Sub

Private Sub Cyan_Q2_AfterUpdate()
  If Me!Cyan_Q2 = -1 Then
    Me!Cyan_Q1 = 0
    Me!Cyan_Q3 = 0
    Me!Cyan_Q4 = 0
  End If
End Sub

Private Sub Cyan_Q3_AfterUpdate()
  If Me!Cyan_Q3 = -1 Then
    Me!Cyan_Q2 = 0
    Me!Cyan_Q1 = 0
    Me!Cyan_Q4 = 0
  End If
End Sub

Private Sub Cyan_Q4_AfterUpdate()
  If Me!Cyan_Q4 = -1 Then
    Me!Cyan_Q2 = 0
    Me!Cyan_Q1 = 0
    Me!Cyan_Q3 = 0
  End If
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    On Error GoTo Err_Handler
    If IsNull(Me.Parent!Visit_Date) Then
      MsgBox "You must enter Visit Date first."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
      GoTo Exit_Procedure
    End If
    ' Create the GUID primary key value
    If IsNull(Me!Quadrat_ID) Then
        If GetDataType("tbl_Quadrat", "Quadrat_ID") = dbText Then
            Me.Quadrat_ID = fxnGUIDGen
        End If
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub



Private Sub Form_Current()
  Me!Total_Cover_Percent = CalcTotalCover()

End Sub


Private Sub Htrail_Cover_AfterUpdate()
  Me!Total_Cover_Percent = CalcTotalCover()

End Sub

Private Sub Htrail_Q1_AfterUpdate()
  If Me!Htrail_Q1 = -1 Then
    Me!Htrail_Q2 = 0
    Me!Htrail_Q3 = 0
    Me!Htrail_Q4 = 0
  End If
End Sub

Private Sub Htrail_Q2_AfterUpdate()
  If Me!Htrail_Q2 = -1 Then
    Me!Htrail_Q1 = 0
    Me!Htrail_Q3 = 0
    Me!Htrail_Q4 = 0
  End If
End Sub

Private Sub Htrail_Q3_AfterUpdate()
  If Me!Htrail_Q3 = -1 Then
    Me!Htrail_Q2 = 0
    Me!Htrail_Q1 = 0
    Me!Htrail_Q4 = 0
  End If
End Sub

Private Sub Htrail_Q4_AfterUpdate()
  If Me!Htrail_Q4 = -1 Then
    Me!Htrail_Q2 = 0
    Me!Htrail_Q3 = 0
    Me!Htrail_Q1 = 0
  End If
End Sub

Private Sub Lichen_Cover_1M2_AfterUpdate()
  If Calc1MCover() > 100 Then
    MsgBox "1 M2 cover cannot exceed 100 percent."
    DoCmd.CancelEvent
    SendKeys "{ESC}"
  End If
End Sub

Private Sub Lichen_Q1_AfterUpdate()
  If Me!Lichen_Q1 = -1 Then
    Me!Lichen_Q2 = 0
    Me!Lichen_Q3 = 0
    Me!Lichen_Q4 = 0
  End If
End Sub

Private Sub Lichen_Q2_AfterUpdate()
  If Me!Lichen_Q2 = -1 Then
    Me!Lichen_Q1 = 0
    Me!Lichen_Q3 = 0
    Me!Lichen_Q4 = 0
  End If
End Sub

Private Sub Lichen_Q3_AfterUpdate()
  If Me!Lichen_Q3 = -1 Then
    Me!Lichen_Q2 = 0
    Me!Lichen_Q1 = 0
    Me!Lichen_Q4 = 0
  End If
End Sub

Private Sub Lichen_Q4_AfterUpdate()
  If Me!Lichen_Q4 = -1 Then
    Me!Lichen_Q2 = 0
    Me!Lichen_Q3 = 0
    Me!Lichen_Q1 = 0
  End If
End Sub

Private Sub Lscat_Cover_AfterUpdate()
  Me!Total_Cover_Percent = CalcTotalCover()

End Sub

Private Sub Lscat_Q1_AfterUpdate()
  If Me!Lscat_Q1 = -1 Then
    Me!Lscat_Q2 = 0
    Me!Lscat_Q3 = 0
    Me!Lscat_Q4 = 0
  End If
End Sub

Private Sub Lscat_Q2_AfterUpdate()
  If Me!Lscat_Q2 = -1 Then
    Me!Lscat_Q1 = 0
    Me!Lscat_Q3 = 0
    Me!Lscat_Q4 = 0
  End If
End Sub

Private Sub Lscat_Q3_AfterUpdate()
  If Me!Lscat_Q3 = -1 Then
    Me!Lscat_Q2 = 0
    Me!Lscat_Q1 = 0
    Me!Lscat_Q4 = 0
  End If
End Sub

Private Sub Lscat_Q4_AfterUpdate()
  If Me!Lscat_Q4 = -1 Then
    Me!Lscat_Q2 = 0
    Me!Lscat_Q3 = 0
    Me!Lscat_Q1 = 0
  End If
End Sub

Private Sub Ltrail_Cover_AfterUpdate()
  Me!Total_Cover_Percent = CalcTotalCover()

End Sub

Private Sub Ltrail_Q1_AfterUpdate()
  If Me!Ltrail_Q1 = -1 Then
    Me!Ltrail_Q2 = 0
    Me!Ltrail_Q3 = 0
    Me!Ltrail_Q4 = 0
  End If
End Sub

Private Sub Ltrail_Q2_AfterUpdate()
  If Me!Ltrail_Q2 = -1 Then
    Me!Ltrail_Q1 = 0
    Me!Ltrail_Q3 = 0
    Me!Ltrail_Q4 = 0
  End If
End Sub

Private Sub Ltrail_Q3_AfterUpdate()
  If Me!Ltrail_Q3 = -1 Then
    Me!Ltrail_Q2 = 0
    Me!Ltrail_Q1 = 0
    Me!Ltrail_Q4 = 0
  End If
End Sub

Private Sub Ltrail_Q4_AfterUpdate()
  If Me!Ltrail_Q4 = -1 Then
    Me!Ltrail_Q2 = 0
    Me!Ltrail_Q3 = 0
    Me!Ltrail_Q1 = 0
  End If
End Sub

Private Sub Moss_Cover_1M2_AfterUpdate()
  If Calc1MCover() > 100 Then
    MsgBox "1 M2 cover cannot exceed 100 percent."
    DoCmd.CancelEvent
    SendKeys "{ESC}"
  End If
End Sub

Private Sub Moss_Q1_AfterUpdate()
  If Me!Moss_Q1 = -1 Then
    Me!Moss_Q2 = 0
    Me!Moss_Q3 = 0
    Me!Moss_Q4 = 0
  End If
End Sub

Private Sub Moss_Q2_AfterUpdate()
  If Me!Moss_Q2 = -1 Then
    Me!Moss_Q1 = 0
    Me!Moss_Q3 = 0
    Me!Moss_Q4 = 0
  End If
End Sub

Private Sub Moss_Q3_AfterUpdate()
  If Me!Moss_Q3 = -1 Then
    Me!Moss_Q2 = 0
    Me!Moss_Q1 = 0
    Me!Moss_Q4 = 0
  End If
End Sub

Private Sub Moss_Q4_AfterUpdate()
  If Me!Moss_Q4 = -1 Then
    Me!Moss_Q2 = 0
    Me!Moss_Q3 = 0
    Me!Moss_Q1 = 0
  End If
End Sub

Private Sub Observer_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Other_Cover_AfterUpdate()
  Me!Total_Cover_Percent = CalcTotalCover()

End Sub

Private Sub Other_Q1_AfterUpdate()
  If Me!Other_Q1 = -1 Then
    Me!Other_Q2 = 0
    Me!Other_Q3 = 0
    Me!Other_Q4 = 0
  End If
End Sub

Private Sub Other_Q2_AfterUpdate()
  If Me!Other_Q2 = -1 Then
    Me!Other_Q1 = 0
    Me!Other_Q3 = 0
    Me!Other_Q4 = 0
  End If
End Sub

Private Sub Other_Q3_AfterUpdate()
  If Me!Other_Q3 = -1 Then
    Me!Other_Q2 = 0
    Me!Other_Q1 = 0
    Me!Other_Q4 = 0
  End If
End Sub

Private Sub Other_Q4_AfterUpdate()
  If Me!Other_Q4 = -1 Then
    Me!Other_Q2 = 0
    Me!Other_Q3 = 0
    Me!Other_Q1 = 0
  End If
End Sub

Private Sub Recorder_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Undiff_Cover_AfterUpdate()
  Me!Total_Cover_Percent = CalcTotalCover()

End Sub

Private Sub Undiff_Q1_AfterUpdate()
  If Me!Undiff_Q1 = -1 Then
    Me!Undiff_Q2 = 0
    Me!Undiff_Q3 = 0
    Me!Undiff_Q4 = 0
  End If
End Sub

Private Sub Undiff_Q2_AfterUpdate()
  If Me!Undiff_Q2 = -1 Then
    Me!Undiff_Q1 = 0
    Me!Undiff_Q3 = 0
    Me!Undiff_Q4 = 0
  End If
End Sub

Private Sub Undiff_Q3_AfterUpdate()
  If Me!Undiff_Q3 = -1 Then
    Me!Undiff_Q2 = 0
    Me!Undiff_Q1 = 0
    Me!Undiff_Q4 = 0
  End If
End Sub

Private Sub Undiff_Q4_AfterUpdate()
  If Me!Undiff_Q4 = -1 Then
    Me!Undiff_Q2 = 0
    Me!Undiff_Q3 = 0
    Me!Undiff_Q1 = 0
  End If
End Sub

Private Sub Vehicle_Cover_AfterUpdate()
  Me!Total_Cover_Percent = CalcTotalCover()

End Sub

Private Sub Vehicle_Q1_AfterUpdate()
  If Me!Vehicle_Q1 = -1 Then
    Me!Vehicle_Q2 = 0
    Me!Vehicle_Q3 = 0
    Me!Vehicle_Q4 = 0
  End If
End Sub

Private Sub Vehicle_Q2_AfterUpdate()
  If Me!Vehicle_Q2 = -1 Then
    Me!Vehicle_Q1 = 0
    Me!Vehicle_Q3 = 0
    Me!Vehicle_Q4 = 0
  End If
End Sub

Private Sub Vehicle_Q3_AfterUpdate()
  If Me!Vehicle_Q3 = -1 Then
    Me!Vehicle_Q2 = 0
    Me!Vehicle_Q1 = 0
    Me!Vehicle_Q4 = 0
  End If
End Sub

Private Sub Vehicle_Q4_AfterUpdate()
  If Me!Vehicle_Q4 = -1 Then
    Me!Vehicle_Q2 = 0
    Me!Vehicle_Q3 = 0
    Me!Vehicle_Q1 = 0
  End If
End Sub

Private Sub Wildlife_Ex_Cover_AfterUpdate()
  Me!Total_Cover_Percent = CalcTotalCover()

End Sub

Private Sub Wildlife_Ex_Q1_AfterUpdate()
  If Me!Wildlife_Ex_Q1 = -1 Then
    Me!Wildlife_Ex_Q2 = 0
    Me!Wildlife_Ex_Q3 = 0
    Me!Wildlife_Ex_Q4 = 0
  End If
End Sub

Private Sub Wildlife_Ex_Q2_AfterUpdate()
  If Me!Wildlife_Ex_Q2 = -1 Then
    Me!Wildlife_Ex_Q1 = 0
    Me!Wildlife_Ex_Q3 = 0
    Me!Wildlife_Ex_Q4 = 0
  End If
End Sub

Private Sub Wildlife_Ex_Q3_AfterUpdate()
  If Me!Wildlife_Ex_Q3 = -1 Then
    Me!Wildlife_Ex_Q2 = 0
    Me!Wildlife_Ex_Q1 = 0
    Me!Wildlife_Ex_Q4 = 0
  End If
End Sub

Private Sub Wildlife_Ex_Q4_AfterUpdate()
  If Me!Wildlife_Ex_Q4 = -1 Then
    Me!Wildlife_Ex_Q2 = 0
    Me!Wildlife_Ex_Q3 = 0
    Me!Wildlife_Ex_Q1 = 0
  End If
End Sub

Private Sub Wscat_Cover_AfterUpdate()
  Me!Total_Cover_Percent = CalcTotalCover()

End Sub

Private Sub Wscat_Q1_AfterUpdate()
  If Me!Wscat_Q1 = -1 Then
    Me!Wscat_Q2 = 0
    Me!Wscat_Q3 = 0
    Me!Wscat_Q4 = 0
  End If
End Sub

Private Sub Wscat_Q2_AfterUpdate()
  If Me!Wscat_Q2 = -1 Then
    Me!Wscat_Q1 = 0
    Me!Wscat_Q3 = 0
    Me!Wscat_Q4 = 0
  End If
End Sub

Private Sub Wscat_Q3_AfterUpdate()
  If Me!Wscat_Q3 = -1 Then
    Me!Wscat_Q2 = 0
    Me!Wscat_Q1 = 0
    Me!Wscat_Q4 = 0
  End If
End Sub

Private Sub Wscat_Q4_AfterUpdate()
  If Me!Wscat_Q4 = -1 Then
    Me!Wscat_Q2 = 0
    Me!Wscat_Q3 = 0
    Me!Wscat_Q1 = 0
  End If
End Sub

Private Sub Wtrail_Cover_AfterUpdate()
  Me!Total_Cover_Percent = CalcTotalCover()

End Sub

Private Sub Wtrail_Q1_AfterUpdate()
  If Me!Wtrail_Q1 = -1 Then
    Me!Wtrail_Q2 = 0
    Me!Wtrail_Q3 = 0
    Me!Wtrail_Q4 = 0
  End If
End Sub

Private Sub Wtrail_Q2_AfterUpdate()
  If Me!Wtrail_Q2 = -1 Then
    Me!Wtrail_Q1 = 0
    Me!Wtrail_Q3 = 0
    Me!Wtrail_Q4 = 0
  End If
End Sub

Private Sub Wtrail_Q3_AfterUpdate()
  If Me!Wtrail_Q3 = -1 Then
    Me!Wtrail_Q2 = 0
    Me!Wtrail_Q1 = 0
    Me!Wtrail_Q4 = 0
  End If
End Sub

Private Sub Wtrail_Q4_AfterUpdate()
  If Me!Wtrail_Q4 = -1 Then
    Me!Wtrail_Q2 = 0
    Me!Wtrail_Q3 = 0
    Me!Wtrail_Q1 = 0
  End If
End Sub
