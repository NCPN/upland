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
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =13680
    DatasheetFontHeight =9
    ItemSuffix =298
    Left =300
    Top =510
    Right =14325
    Bottom =8685
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x49aadcf15012e340
    End
    RecordSource ="tbl_Soil_Stability"
    Caption ="fsub_Soil_Stability"
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
            Height =7200
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9960
                    Top =180
                    Width =510
                    ColumnWidth =2310
                    Name ="Soil_ID"
                    ControlSource ="Soil_ID"
                    StatusBarText ="Unique record identifier - primary key"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10530
                    Top =180
                    Width =510
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Event_ID"
                    ControlSource ="Event_ID"
                    StatusBarText ="Foreign key to tbl_Events"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7620
                    Top =180
                    Width =1035
                    ColumnWidth =1035
                    TabIndex =4
                    Name ="Visit_Date"
                    ControlSource ="Visit_Date"
                    Format ="Short Date"
                    StatusBarText ="Date of visit."
                    InputMask ="99/99/0000;0;_"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7020
                            Top =180
                            Width =540
                            Height =240
                            FontWeight =700
                            Name ="Visit_Date_Label"
                            Caption ="Date"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1200
                    Top =1800
                    Width =839
                    Height =300
                    ColumnWidth =600
                    TabIndex =6
                    Name ="T1500_Pos"
                    ControlSource ="T1500_Pos"
                    StatusBarText ="Transect 1 5 minute position"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =2040
                            Top =1800
                            Width =839
                            Height =300
                            Name ="T1500_Pos_Label"
                            Caption ="5:00"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1200
                    Top =3840
                    Width =839
                    Height =299
                    ColumnWidth =600
                    TabIndex =18
                    Name ="T1545_Pos"
                    ControlSource ="T1545_Pos"
                    StatusBarText ="Transect 1 5:45 minute position"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =2040
                            Top =3840
                            Width =839
                            Height =299
                            Name ="T1545_Pos_Label"
                            Caption ="5:45"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1200
                    Top =2100
                    Width =839
                    Height =299
                    ColumnWidth =600
                    TabIndex =30
                    Name ="T2630_Pos"
                    ControlSource ="T2630_Pos"
                    StatusBarText ="Transect 1 6:30 minute position"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =2040
                            Top =2100
                            Width =839
                            Height =299
                            Name ="T1630_Pos_Label"
                            Caption ="6:30"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1200
                    Top =4140
                    Width =839
                    Height =299
                    ColumnWidth =600
                    TabIndex =42
                    Name ="T2715_Pos"
                    ControlSource ="T2715_Pos"
                    StatusBarText ="Transect 1 7:15 minute position"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =2040
                            Top =4140
                            Width =839
                            Height =299
                            Name ="T1715_Pos_Label"
                            Caption ="7:15"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1200
                    Top =2400
                    Width =839
                    Height =299
                    ColumnWidth =600
                    TabIndex =54
                    Name ="T3800_Pos"
                    ControlSource ="T3800_Pos"
                    StatusBarText ="Transect 1 8:00 minute position"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =2040
                            Top =2400
                            Width =839
                            Height =299
                            Name ="T1800_Pos_Label"
                            Caption ="8:00"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1200
                    Top =4440
                    Width =839
                    Height =299
                    ColumnWidth =600
                    TabIndex =66
                    Name ="T3845_Pos"
                    ControlSource ="T3845_Pos"
                    StatusBarText ="Transect 1 8:45 minute position"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =2040
                            Top =4440
                            Width =839
                            Height =299
                            Name ="T1845_Pos_Label"
                            Caption ="8:45"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4980
                    Top =1800
                    Width =839
                    Height =299
                    ColumnWidth =600
                    TabIndex =10
                    Name ="T1515_Pos"
                    ControlSource ="T1515_Pos"
                    StatusBarText ="Transect 2 5:15 minute position"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =5820
                            Top =1800
                            Width =839
                            Height =299
                            Name ="T2515_Pos_Label"
                            Caption ="5:15"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4980
                    Top =3840
                    Width =839
                    Height =299
                    ColumnWidth =600
                    TabIndex =22
                    Name ="T1600_Pos"
                    ControlSource ="T1600_Pos"
                    StatusBarText ="Transect 2 6:00 minute position"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =5820
                            Top =3840
                            Width =839
                            Height =299
                            Name ="T2600_Pos_Label"
                            Caption ="6:00"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4980
                    Top =2100
                    Width =839
                    Height =299
                    ColumnWidth =600
                    TabIndex =34
                    Name ="T2645_Pos"
                    ControlSource ="T2645_Pos"
                    StatusBarText ="Transect 2 6:45 minute position"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =5820
                            Top =2100
                            Width =839
                            Height =299
                            Name ="T2645_Pos_Label"
                            Caption ="6:45"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4980
                    Top =4140
                    Width =839
                    Height =299
                    ColumnWidth =900
                    TabIndex =46
                    Name ="T2730_Pos"
                    ControlSource ="T2730_Pos"
                    StatusBarText ="Transect 2 7:30 minute Rating  1-6"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =5820
                            Top =4140
                            Width =839
                            Height =299
                            Name ="T2730_Rating_Label"
                            Caption ="7:30"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4980
                    Top =2400
                    Width =839
                    Height =299
                    ColumnWidth =600
                    TabIndex =58
                    Name ="T3815_Pos"
                    ControlSource ="T3815_Pos"
                    StatusBarText ="Transect 2 8:15 minute position"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =5820
                            Top =2400
                            Width =839
                            Height =299
                            Name ="T2815_Pos_Label"
                            Caption ="8:15"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4980
                    Top =4440
                    Width =839
                    Height =299
                    ColumnWidth =600
                    TabIndex =70
                    Name ="T3900_Pos"
                    ControlSource ="T3900_Pos"
                    StatusBarText ="Transect 2 9:00 minute position"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =5820
                            Top =4440
                            Width =839
                            Height =299
                            Name ="T2900_Pos_Label"
                            Caption ="9:00"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8760
                    Top =1800
                    Width =839
                    Height =299
                    ColumnWidth =600
                    TabIndex =14
                    Name ="T1530_Pos"
                    ControlSource ="T1530_Pos"
                    StatusBarText ="Transect 3 5:30 minute position"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =9600
                            Top =1800
                            Width =839
                            Height =299
                            Name ="T3530_Pos_Label"
                            Caption ="5:30"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8760
                    Top =3840
                    Width =839
                    Height =299
                    ColumnWidth =600
                    TabIndex =26
                    Name ="T1615_Pos"
                    ControlSource ="T1615_Pos"
                    StatusBarText ="Transect 3 6:15 minute position"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =9600
                            Top =3840
                            Width =839
                            Height =299
                            Name ="T3515_Pos_Label"
                            Caption ="6:15"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8760
                    Top =2100
                    Width =839
                    Height =299
                    ColumnWidth =600
                    TabIndex =38
                    Name ="T2700_Pos"
                    ControlSource ="T2700_Pos"
                    StatusBarText ="Transect 3 7:00 minute position"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =9600
                            Top =2100
                            Width =839
                            Height =299
                            Name ="T3700_Pos_Label"
                            Caption ="7:00"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8760
                    Top =4140
                    Width =839
                    Height =299
                    ColumnWidth =600
                    TabIndex =50
                    Name ="T2745_Pos"
                    ControlSource ="T2745_Pos"
                    StatusBarText ="Transect 3 7:45 minute position"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =9600
                            Top =4140
                            Width =839
                            Height =299
                            Name ="T3745_Pos_Label"
                            Caption ="7:45"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8760
                    Top =2400
                    Width =839
                    Height =299
                    ColumnWidth =600
                    TabIndex =62
                    Name ="T3830_Pos"
                    ControlSource ="T3830_Pos"
                    StatusBarText ="Transect 3 8:30 minute position"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =9600
                            Top =2400
                            Width =839
                            Height =299
                            Name ="T3830_Pos_Label"
                            Caption ="8:30"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8760
                    Top =4440
                    Width =839
                    Height =299
                    ColumnWidth =600
                    TabIndex =74
                    Name ="T3915_Pos"
                    ControlSource ="T3915_Pos"
                    StatusBarText ="Transect 3 9:15 minute position"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =9600
                            Top =4440
                            Width =839
                            Height =299
                            Name ="T3915_Pos_Label"
                            Caption ="9:15"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =2520
                    Top =5700
                    Width =7200
                    Height =645
                    TabIndex =78
                    Name ="Comments"
                    ControlSource ="Comments"
                    StatusBarText ="Soil stability comments"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2520
                            Top =5460
                            Width =1020
                            Height =240
                            FontWeight =700
                            Name ="Comments_Label"
                            Caption ="Comments"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =1965
                    Left =1080
                    Top =180
                    Width =2040
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Observer"
                    ControlSource ="Observer"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Contacts.Contact_ID, tlu_Contacts.Last_Name, tlu_Contacts.First_Name "
                        "FROM tlu_Contacts WHERE (((tlu_Contacts.Active)=1)); "
                    ColumnWidths ="0;975;990"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =180
                            Width =840
                            Height =245
                            FontWeight =700
                            Name ="Label123"
                            Caption ="Observer"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =1965
                    Left =4500
                    Top =180
                    Width =2040
                    TabIndex =3
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
                            OverlapFlags =85
                            Left =3600
                            Top =180
                            Width =840
                            Height =245
                            FontWeight =700
                            Name ="Recorder_Label"
                            Caption ="Recorder"
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =990
                    Left =1380
                    Top =660
                    Width =1200
                    TabIndex =5
                    Name ="Sample_Type"
                    ControlSource ="Sample_Type"
                    RowSourceType ="Value List"
                    RowSource ="1;\"Surface\""
                    ColumnWidths ="0;990"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =180
                            Top =660
                            Width =1200
                            Height =245
                            FontWeight =700
                            Name ="Sample Type_Label"
                            Caption ="Sample Type"
                            EventProcPrefix ="Sample_Type_Label"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2310
                    Left =2880
                    Top =1800
                    Width =839
                    Height =300
                    TabIndex =7
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"4\""
                    Name ="T1500_Veg"
                    ControlSource ="T1500_Veg"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Cover_Class.Cover_Class, tlu_Cover_Class.Cover_Description FROM tlu_C"
                        "over_Class; "
                    ColumnWidths ="360;1950"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ListWidth =1440
                    Left =3720
                    Top =1800
                    Width =839
                    Height =300
                    TabIndex =8
                    Name ="T1500_Rating"
                    ControlSource ="T1500_Rating"
                    RowSourceType ="Value List"
                    RowSource ="1;2;3;4;5;6"
                    ColumnWidths ="1440"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2280
                    Left =2880
                    Top =3840
                    Width =839
                    Height =300
                    TabIndex =19
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"4\""
                    Name ="T1545_Veg"
                    ControlSource ="T1545_Veg"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Cover_Class.Cover_Class, tlu_Cover_Class.Cover_Description FROM tlu_C"
                        "over_Class; "
                    ColumnWidths ="330;1950"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ListWidth =240
                    Left =3720
                    Top =3840
                    Width =839
                    Height =300
                    TabIndex =20
                    Name ="T1545_Rating"
                    ControlSource ="T1545_Rating"
                    RowSourceType ="Value List"
                    RowSource ="1;2;3;4;5;6"
                    ColumnWidths ="240"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =2040
                    Top =1560
                    Width =840
                    Height =240
                    FontWeight =700
                    Name ="Label136"
                    Caption ="Dip Time"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =1200
                    Top =1560
                    Width =840
                    Height =240
                    FontWeight =700
                    Name ="Label137"
                    Caption ="Pos."
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =2880
                    Top =1560
                    Width =840
                    Height =240
                    FontWeight =700
                    Name ="Label138"
                    Caption ="Veg"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =3720
                    Top =1560
                    Width =840
                    Height =240
                    FontWeight =700
                    Name ="Label139"
                    Caption ="Rating"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =180
                    Top =1800
                    Width =1020
                    Height =299
                    FontWeight =700
                    Name ="Label140"
                    Caption ="Transect 1"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2250
                    Left =2880
                    Top =2100
                    Width =840
                    Height =300
                    TabIndex =31
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"4\""
                    Name ="T2630_Veg"
                    ControlSource ="T2630_Veg"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Cover_Class.Cover_Class, tlu_Cover_Class.Cover_Description FROM tlu_C"
                        "over_Class; "
                    ColumnWidths ="300;1950"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ListWidth =255
                    Left =3720
                    Top =2100
                    Width =840
                    Height =300
                    TabIndex =32
                    Name ="T2630_Rating"
                    ControlSource ="T2630_Rating"
                    RowSourceType ="Value List"
                    RowSource ="1;2;3;4;5;6"
                    ColumnWidths ="255"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2295
                    Left =2880
                    Top =4140
                    Width =840
                    Height =300
                    TabIndex =43
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"4\""
                    Name ="T2715_Veg"
                    ControlSource ="T2715_Veg"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Cover_Class.Cover_Class, tlu_Cover_Class.Cover_Description FROM tlu_C"
                        "over_Class; "
                    ColumnWidths ="345;1950"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ListWidth =210
                    Left =3720
                    Top =4140
                    Width =840
                    Height =300
                    TabIndex =44
                    Name ="T2715_Rating"
                    ControlSource ="T2715_Rating"
                    RowSourceType ="Value List"
                    RowSource ="1;2;3;4;5;6"
                    ColumnWidths ="210"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2265
                    Left =2880
                    Top =2400
                    Width =840
                    Height =300
                    TabIndex =55
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"4\""
                    Name ="T3800_Veg"
                    ControlSource ="T3800_Veg"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Cover_Class.Cover_Class, tlu_Cover_Class.Cover_Description FROM tlu_C"
                        "over_Class; "
                    ColumnWidths ="315;1950"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ListWidth =195
                    Left =3720
                    Top =2400
                    Width =840
                    Height =300
                    TabIndex =56
                    Name ="T3800_Rating"
                    ControlSource ="T3800_Rating"
                    RowSourceType ="Value List"
                    RowSource ="1;2;3;4;5;6"
                    ColumnWidths ="195"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2250
                    Left =2880
                    Top =4440
                    Width =840
                    Height =300
                    TabIndex =67
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"4\""
                    Name ="T3845_Veg"
                    ControlSource ="T3845_Veg"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Cover_Class.Cover_Class, tlu_Cover_Class.Cover_Description FROM tlu_C"
                        "over_Class; "
                    ColumnWidths ="300;1950"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ListWidth =195
                    Left =3720
                    Top =4440
                    Width =840
                    Height =300
                    TabIndex =68
                    Name ="T3845_Rating"
                    ControlSource ="T3845_Rating"
                    RowSourceType ="Value List"
                    RowSource ="1;2;3;4;5;6"
                    ColumnWidths ="195"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =180
                    Top =2100
                    Width =1020
                    Height =299
                    FontWeight =700
                    Name ="Label163"
                    Caption ="Transect 2"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    Left =180
                    Top =2400
                    Width =1020
                    Height =299
                    FontWeight =700
                    Name ="Label169"
                    Caption ="Transect 3"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2265
                    Left =6660
                    Top =1800
                    Width =840
                    Height =300
                    TabIndex =11
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"4\""
                    Name ="T1515_Veg"
                    ControlSource ="T1515_Veg"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Cover_Class.Cover_Class, tlu_Cover_Class.Cover_Description FROM tlu_C"
                        "over_Class; "
                    ColumnWidths ="315;1950"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ListWidth =1440
                    Left =7500
                    Top =1800
                    Width =840
                    Height =300
                    TabIndex =12
                    Name ="T1515_Rating"
                    ControlSource ="T1515_Rating"
                    RowSourceType ="Value List"
                    RowSource ="1;2;3;4;5;6"
                    ColumnWidths ="1440"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2265
                    Left =6660
                    Top =3840
                    Width =840
                    Height =300
                    TabIndex =23
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"4\""
                    Name ="T1600_Veg"
                    ControlSource ="T1600_Veg"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Cover_Class.Cover_Class, tlu_Cover_Class.Cover_Description FROM tlu_C"
                        "over_Class; "
                    ColumnWidths ="315;1950"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ListWidth =210
                    Left =7500
                    Top =3840
                    Width =840
                    Height =300
                    TabIndex =24
                    Name ="T1600_Rating"
                    ControlSource ="T1600_Rating"
                    RowSourceType ="Value List"
                    RowSource ="1;2;3;4;5;6"
                    ColumnWidths ="210"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2265
                    Left =6660
                    Top =2100
                    Width =840
                    Height =300
                    TabIndex =35
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"4\""
                    Name ="T2645_Veg"
                    ControlSource ="T2645_Veg"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Cover_Class.Cover_Class, tlu_Cover_Class.Cover_Description FROM tlu_C"
                        "over_Class; "
                    ColumnWidths ="315;1950"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ListWidth =225
                    Left =7500
                    Top =2100
                    Width =840
                    Height =300
                    TabIndex =36
                    Name ="T2645_Rating"
                    ControlSource ="T2645_Rating"
                    RowSourceType ="Value List"
                    RowSource ="1;2;3;4;5;6"
                    ColumnWidths ="225"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2280
                    Left =6660
                    Top =4140
                    Width =840
                    Height =300
                    TabIndex =47
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"4\""
                    Name ="T2730_Veg"
                    ControlSource ="T2730_Veg"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Cover_Class.Cover_Class, tlu_Cover_Class.Cover_Description FROM tlu_C"
                        "over_Class; "
                    ColumnWidths ="330;1950"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ListWidth =270
                    Left =7500
                    Top =4140
                    Width =840
                    Height =300
                    TabIndex =48
                    Name ="T2730_Rating"
                    ControlSource ="T2730_Rating"
                    RowSourceType ="Value List"
                    RowSource ="1;2;3;4;5;6"
                    ColumnWidths ="270"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2295
                    Left =6660
                    Top =2400
                    Width =840
                    Height =300
                    TabIndex =59
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"4\""
                    Name ="T3815_Veg"
                    ControlSource ="T3815_Veg"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Cover_Class.Cover_Class, tlu_Cover_Class.Cover_Description FROM tlu_C"
                        "over_Class; "
                    ColumnWidths ="345;1950"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ListWidth =255
                    Left =7500
                    Top =2400
                    Width =840
                    Height =300
                    TabIndex =60
                    Name ="T3815_Rating"
                    ControlSource ="T3815_Rating"
                    RowSourceType ="Value List"
                    RowSource ="1;2;3;4;5;6"
                    ColumnWidths ="255"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2295
                    Left =6660
                    Top =4440
                    Width =840
                    Height =300
                    TabIndex =71
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"4\""
                    Name ="T3900_Veg"
                    ControlSource ="T3900_Veg"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Cover_Class.Cover_Class, tlu_Cover_Class.Cover_Description FROM tlu_C"
                        "over_Class; "
                    ColumnWidths ="345;1950"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ListWidth =255
                    Left =7500
                    Top =4440
                    Width =840
                    Height =300
                    TabIndex =72
                    Name ="T3900_Rating"
                    ControlSource ="T3900_Rating"
                    RowSourceType ="Value List"
                    RowSource ="1;2;3;4;5;6"
                    ColumnWidths ="255"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2310
                    Left =10440
                    Top =1800
                    Width =840
                    Height =300
                    TabIndex =15
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"4\""
                    Name ="T1530_Veg"
                    ControlSource ="T1530_Veg"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Cover_Class.Cover_Class, tlu_Cover_Class.Cover_Description FROM tlu_C"
                        "over_Class; "
                    ColumnWidths ="360;1950"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ListWidth =225
                    Left =11280
                    Top =1800
                    Width =840
                    Height =300
                    TabIndex =16
                    Name ="T1530_Rating"
                    ControlSource ="T1530_Rating"
                    RowSourceType ="Value List"
                    RowSource ="1;2;3;4;5;6"
                    ColumnWidths ="225"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2295
                    Left =10440
                    Top =3840
                    Width =840
                    Height =300
                    TabIndex =27
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"4\""
                    Name ="T1615_Veg"
                    ControlSource ="T1615_Veg"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Cover_Class.Cover_Class, tlu_Cover_Class.Cover_Description FROM tlu_C"
                        "over_Class; "
                    ColumnWidths ="345;1950"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ListWidth =225
                    Left =11280
                    Top =3840
                    Width =840
                    Height =300
                    TabIndex =28
                    Name ="T1615_Rating"
                    ControlSource ="T1615_Rating"
                    RowSourceType ="Value List"
                    RowSource ="1;2;3;4;5;6"
                    ColumnWidths ="225"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2280
                    Left =10440
                    Top =2100
                    Width =840
                    Height =300
                    TabIndex =39
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"4\""
                    Name ="T2700_Veg"
                    ControlSource ="T2700_Veg"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Cover_Class.Cover_Class, tlu_Cover_Class.Cover_Description FROM tlu_C"
                        "over_Class; "
                    ColumnWidths ="330;1950"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ListWidth =285
                    Left =11280
                    Top =2100
                    Width =840
                    Height =300
                    TabIndex =40
                    Name ="T2700_Rating"
                    ControlSource ="T2700_Rating"
                    RowSourceType ="Value List"
                    RowSource ="1;2;3;4;5;6"
                    ColumnWidths ="285"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2340
                    Left =10440
                    Top =4140
                    Width =840
                    Height =300
                    TabIndex =51
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"4\""
                    Name ="T2745_Veg"
                    ControlSource ="T2745_Veg"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Cover_Class.Cover_Class, tlu_Cover_Class.Cover_Description FROM tlu_C"
                        "over_Class; "
                    ColumnWidths ="390;1950"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ListWidth =255
                    Left =11280
                    Top =4140
                    Width =840
                    Height =300
                    TabIndex =52
                    Name ="T2745_Rating"
                    ControlSource ="T2745_Rating"
                    RowSourceType ="Value List"
                    RowSource ="1;2;3;4;5;6"
                    ColumnWidths ="255"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2295
                    Left =10440
                    Top =2400
                    Width =840
                    Height =300
                    TabIndex =63
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"4\""
                    Name ="T3830_Veg"
                    ControlSource ="T3830_Veg"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Cover_Class.Cover_Class, tlu_Cover_Class.Cover_Description FROM tlu_C"
                        "over_Class; "
                    ColumnWidths ="345;1950"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ListWidth =1440
                    Left =11280
                    Top =2400
                    Width =840
                    Height =300
                    TabIndex =64
                    Name ="T3830_Rating"
                    ControlSource ="T3830_Rating"
                    RowSourceType ="Value List"
                    RowSource ="1;2;3;4;5;6"
                    ColumnWidths ="1440"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2295
                    Left =10440
                    Top =4440
                    Width =840
                    Height =300
                    TabIndex =75
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"4\""
                    Name ="T3915_Veg"
                    ControlSource ="T3915_Veg"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Cover_Class.Cover_Class, tlu_Cover_Class.Cover_Description FROM tlu_C"
                        "over_Class; "
                    ColumnWidths ="345;1950"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ListWidth =285
                    Left =11280
                    Top =4440
                    Width =840
                    Height =300
                    TabIndex =76
                    Name ="T3915_Rating"
                    ControlSource ="T3915_Rating"
                    RowSourceType ="Value List"
                    RowSource ="1;2;3;4;5;6"
                    ColumnWidths ="285"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =2640
                    Top =660
                    Width =306
                    Height =246
                    TabIndex =79
                    Name ="ButtonPrevious"
                    Caption ="Command219"
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
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =3000
                    Top =660
                    Width =306
                    Height =246
                    TabIndex =80
                    Name ="ButtonNext"
                    Caption ="Command220"
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
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Next Record"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =4680
                    Top =1860
                    Width =240
                    TabIndex =9
                    Name ="T1500_Hydro"
                    ControlSource ="T1500_Hydro"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =255
                    Left =4560
                    Top =1800
                    Width =420
                    Height =300
                    BorderColor =1
                    Name ="Box225"
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =4680
                    Top =3900
                    Width =240
                    TabIndex =21
                    Name ="T1545_Hydro"
                    ControlSource ="T1545_Hydro"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =255
                    Left =4560
                    Top =3840
                    Width =420
                    Height =300
                    BorderColor =1
                    Name ="Box227"
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =4680
                    Top =2160
                    Width =240
                    TabIndex =33
                    Name ="T2630_Hydro"
                    ControlSource ="T2630_Hydro"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =255
                    Left =4560
                    Top =2100
                    Width =420
                    Height =300
                    BorderColor =1
                    Name ="Box229"
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =4680
                    Top =4200
                    Width =240
                    TabIndex =45
                    Name ="T2715_Hydro"
                    ControlSource ="T2715_Hydro"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =255
                    Left =4560
                    Top =4140
                    Width =420
                    Height =300
                    BorderColor =1
                    Name ="Box231"
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =4680
                    Top =2460
                    Width =240
                    TabIndex =57
                    Name ="T3800_Hydro"
                    ControlSource ="T3800_Hydro"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =247
                    Left =4560
                    Top =2400
                    Width =420
                    Height =300
                    BorderColor =1
                    Name ="Box233"
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =4680
                    Top =4500
                    Width =240
                    TabIndex =69
                    Name ="T3845_Hydro"
                    ControlSource ="T3845_Hydro"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =247
                    Left =4560
                    Top =4440
                    Width =420
                    Height =300
                    BorderColor =1
                    Name ="Box235"
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =8460
                    Top =1860
                    Width =240
                    TabIndex =13
                    Name ="T1515_Hydro"
                    ControlSource ="T1515_Hydro"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =255
                    Left =8340
                    Top =1800
                    Width =420
                    Height =300
                    BorderColor =1
                    Name ="Box237"
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =8460
                    Top =3900
                    Width =240
                    TabIndex =25
                    Name ="T1600_Hydro"
                    ControlSource ="T1600_Hydro"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =255
                    Left =8340
                    Top =3840
                    Width =420
                    Height =300
                    BorderColor =1
                    Name ="Box239"
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =8460
                    Top =2160
                    Width =240
                    TabIndex =37
                    Name ="T2645_Hydro"
                    ControlSource ="T2645_Hydro"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =255
                    Left =8340
                    Top =2100
                    Width =420
                    Height =300
                    BorderColor =1
                    Name ="Box241"
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =8460
                    Top =4200
                    Width =240
                    TabIndex =49
                    Name ="T2730_Hydro"
                    ControlSource ="T2730_Hydro"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =255
                    Left =8340
                    Top =4140
                    Width =420
                    Height =300
                    BorderColor =1
                    Name ="Box243"
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =8460
                    Top =2460
                    Width =240
                    TabIndex =61
                    Name ="T3815_Hydro"
                    ControlSource ="T3815_Hydro"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =247
                    Left =8340
                    Top =2400
                    Width =420
                    Height =300
                    BorderColor =1
                    Name ="Box245"
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =8460
                    Top =4500
                    Width =240
                    TabIndex =73
                    Name ="T3900_Hydro"
                    ControlSource ="T3900_Hydro"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =247
                    Left =8340
                    Top =4440
                    Width =420
                    Height =300
                    BorderColor =1
                    Name ="Box247"
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =12240
                    Top =1860
                    Width =240
                    TabIndex =17
                    Name ="T1530_Hydro"
                    ControlSource ="T1530_Hydro"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =255
                    Left =12120
                    Top =1800
                    Width =420
                    Height =300
                    BorderColor =1
                    Name ="Box249"
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =12240
                    Top =3900
                    Width =240
                    TabIndex =29
                    Name ="T1615_Hydro"
                    ControlSource ="T1615_Hydro"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =255
                    Left =12120
                    Top =3840
                    Width =420
                    Height =300
                    BorderColor =1
                    Name ="Box251"
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =12240
                    Top =2160
                    Width =240
                    TabIndex =41
                    Name ="T2700_Hydro"
                    ControlSource ="T2700_Hydro"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =255
                    Left =12120
                    Top =2100
                    Width =420
                    Height =300
                    BorderColor =1
                    Name ="Box253"
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =12240
                    Top =4200
                    Width =240
                    TabIndex =53
                    Name ="T2745_Hydro"
                    ControlSource ="T2745_Hydro"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =255
                    Left =12120
                    Top =4140
                    Width =420
                    Height =300
                    BorderColor =1
                    Name ="Box255"
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =12240
                    Top =2460
                    Width =240
                    TabIndex =65
                    Name ="T3830_Hydro"
                    ControlSource ="T3830_Hydro"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =247
                    Left =12120
                    Top =2400
                    Width =420
                    Height =300
                    BorderColor =1
                    Name ="Box257"
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =12240
                    Top =4500
                    Width =240
                    TabIndex =77
                    Name ="T3915_Hydro"
                    ControlSource ="T3915_Hydro"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =247
                    Left =12120
                    Top =4440
                    Width =420
                    Height =300
                    BorderColor =1
                    Name ="Box259"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =127
                    TextAlign =2
                    Left =4560
                    Top =1560
                    Width =420
                    Height =240
                    FontWeight =700
                    Name ="Label260"
                    Caption ="HP"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =5820
                    Top =1560
                    Width =840
                    Height =240
                    FontWeight =700
                    Name ="Label264"
                    Caption ="Dip Time"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =127
                    TextAlign =2
                    Left =4980
                    Top =1560
                    Width =840
                    Height =240
                    FontWeight =700
                    Name ="Label265"
                    Caption ="Pos."
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =6660
                    Top =1560
                    Width =840
                    Height =240
                    FontWeight =700
                    Name ="Label266"
                    Caption ="Veg"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =127
                    TextAlign =2
                    Left =7500
                    Top =1560
                    Width =840
                    Height =240
                    FontWeight =700
                    Name ="Label267"
                    Caption ="Rating"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =127
                    TextAlign =2
                    Left =8340
                    Top =1560
                    Width =420
                    Height =240
                    FontWeight =700
                    Name ="Label268"
                    Caption ="HP"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =9600
                    Top =1560
                    Width =840
                    Height =240
                    FontWeight =700
                    Name ="Label269"
                    Caption ="Dip Time"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =127
                    TextAlign =2
                    Left =8760
                    Top =1560
                    Width =840
                    Height =240
                    FontWeight =700
                    Name ="Label270"
                    Caption ="Pos."
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =10440
                    Top =1560
                    Width =840
                    Height =240
                    FontWeight =700
                    Name ="Label271"
                    Caption ="Veg"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =127
                    TextAlign =2
                    Left =11280
                    Top =1560
                    Width =840
                    Height =240
                    FontWeight =700
                    Name ="Label272"
                    Caption ="Rating"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =127
                    TextAlign =2
                    Left =12120
                    Top =1560
                    Width =420
                    Height =240
                    FontWeight =700
                    Name ="Label273"
                    Caption ="HP"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =127
                    TextAlign =2
                    Left =1200
                    Top =1260
                    Width =3780
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label274"
                    Caption ="Sample 1"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =127
                    TextAlign =2
                    Left =4980
                    Top =1260
                    Width =3780
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label275"
                    Caption ="Sample 2"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =119
                    TextAlign =2
                    Left =8760
                    Top =1260
                    Width =3780
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label276"
                    Caption ="Sample 3"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =1200
                    Top =3300
                    Width =3780
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label277"
                    Caption ="Sample 4"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =4980
                    Top =3300
                    Width =3780
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label278"
                    Caption ="Sample 5"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =8760
                    Top =3300
                    Width =3780
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label279"
                    Caption ="Sample 6"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =180
                    Top =3840
                    Width =1020
                    Height =299
                    FontWeight =700
                    Name ="Label280"
                    Caption ="Transect 1"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =180
                    Top =4140
                    Width =1020
                    Height =299
                    FontWeight =700
                    Name ="Label281"
                    Caption ="Transect 2"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    Left =180
                    Top =4440
                    Width =1020
                    Height =299
                    FontWeight =700
                    Name ="Label282"
                    Caption ="Transect 3"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =2040
                    Top =3600
                    Width =840
                    Height =240
                    FontWeight =700
                    Name ="Label283"
                    Caption ="Dip Time"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    Left =1200
                    Top =3600
                    Width =840
                    Height =240
                    FontWeight =700
                    Name ="Label284"
                    Caption ="Pos."
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =2880
                    Top =3600
                    Width =840
                    Height =240
                    FontWeight =700
                    Name ="Label285"
                    Caption ="Veg"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =127
                    TextAlign =2
                    Left =3720
                    Top =3600
                    Width =840
                    Height =240
                    FontWeight =700
                    Name ="Label286"
                    Caption ="Rating"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =127
                    TextAlign =2
                    Left =4560
                    Top =3600
                    Width =420
                    Height =240
                    FontWeight =700
                    Name ="Label287"
                    Caption ="HP"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =5820
                    Top =3600
                    Width =840
                    Height =240
                    FontWeight =700
                    Name ="Label288"
                    Caption ="Dip Time"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =119
                    TextAlign =2
                    Left =4980
                    Top =3600
                    Width =840
                    Height =240
                    FontWeight =700
                    Name ="Label289"
                    Caption ="Pos."
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =6660
                    Top =3600
                    Width =840
                    Height =240
                    FontWeight =700
                    Name ="Label290"
                    Caption ="Veg"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =127
                    TextAlign =2
                    Left =7500
                    Top =3600
                    Width =840
                    Height =240
                    FontWeight =700
                    Name ="Label291"
                    Caption ="Rating"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =127
                    TextAlign =2
                    Left =8340
                    Top =3600
                    Width =420
                    Height =240
                    FontWeight =700
                    Name ="Label292"
                    Caption ="HP"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =9600
                    Top =3600
                    Width =840
                    Height =240
                    FontWeight =700
                    Name ="Label293"
                    Caption ="Dip Time"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =119
                    TextAlign =2
                    Left =8760
                    Top =3600
                    Width =840
                    Height =240
                    FontWeight =700
                    Name ="Label294"
                    Caption ="Pos."
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =10440
                    Top =3600
                    Width =840
                    Height =240
                    FontWeight =700
                    Name ="Label295"
                    Caption ="Veg"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =127
                    TextAlign =2
                    Left =11280
                    Top =3600
                    Width =840
                    Height =240
                    FontWeight =700
                    Name ="Label296"
                    Caption ="Rating"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =119
                    TextAlign =2
                    Left =12120
                    Top =3600
                    Width =420
                    Height =240
                    FontWeight =700
                    Name ="Label297"
                    Caption ="HP"
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

Private Sub Comments_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
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
    If IsNull(Me!Soil_ID) Then
        If GetDataType("tbl_Soil_Stability", "Soil_ID") = dbText Then
            Me.Soil_ID = fxnGUIDGen
        End If
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub
Private Sub ButtonPrevious_Click()
On Error GoTo Err_ButtonPrevious_Click


    DoCmd.GoToRecord , , acPrevious

Exit_ButtonPrevious_Click:
    Exit Sub

Err_ButtonPrevious_Click:
    MsgBox Err.Description
    Resume Exit_ButtonPrevious_Click
    
End Sub
Private Sub ButtonNext_Click()
On Error GoTo Err_ButtonNext_Click


    DoCmd.GoToRecord , , acNext

Exit_ButtonNext_Click:
    Exit Sub

Err_ButtonNext_Click:
    MsgBox Err.Description
    Resume Exit_ButtonNext_Click
    
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

Private Sub Sample_Type_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T1500_Hydro_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T1500_Pos_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T1500_Rating_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T1500_Veg_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T1545_Hydro_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T1545_Pos_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T1545_Rating_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T1545_Veg_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T1630_Hydro_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T1630_Pos_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T1630_Rating_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T1630_Veg_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T1715_Hydro_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T1715_Pos_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T1715_Rating_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T1715_Veg_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T1800_Hydro_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T1800_Pos_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T1800_Rating_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T1800_Veg_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T1845_Hydro_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T1845_Pos_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T1845_Rating_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T1845_Veg_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T2515_Hydro_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T2515_Pos_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T2515_Rating_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T2515_Veg_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T2600_Hydro_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T2600_Pos_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T2600_Rating_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T2600_Veg_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T2645_Hydro_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T2645_Pos_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T2645_Rating_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T2645_Veg_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T2730_Hydro_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T2730_Pos_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T2730_Rating_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T2730_Veg_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T2815_Hydro_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T2815_Pos_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T2815_Rating_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T2815_Veg_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T2900_Hydro_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T2900_Pos_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T2900_Rating_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T2900_Veg_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T3530_Hydro_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T3530_Pos_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T3530_Rating_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T3530_Veg_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T3615_Hydro_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T3615_Pos_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T3615_Rating_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T3615_Veg_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T3700_Hydro_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T3700_Pos_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T3700_Rating_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T3700_Veg_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T3745_Hydro_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T3745_Pos_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T3745_Rating_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T3745_Veg_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T3830_Hydro_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T3830_Pos_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T3830_Rating_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T3830_Veg_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T3915_Hydro_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T3915_Pos_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T3915_Rating_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub T3915_Veg_KeyDown(KeyCode As Integer, Shift As Integer)
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
