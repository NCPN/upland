Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    KeyPreview = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =12960
    DatasheetFontHeight =9
    ItemSuffix =140
    Left =1800
    Top =3264
    Right =14664
    Bottom =7668
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x9ff2d503acd4e340
    End
    RecordSource ="tbl_Fuels"
    Caption ="fsub_Fuels_LD"
    BeforeInsert ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnKeyDown ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            TextAlign =2
            FontWeight =700
            BackColor =-2147483633
            ForeColor =-2147483630
            FontName ="Tahoma"
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
            Height =4680
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6000
                    Top =1320
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =20
                    BackColor =62207
                    Name ="2Litter_A"
                    ControlSource ="2Litter_A"
                    StatusBarText ="Litter depth at 2 meter point for transect A in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    EventProcPrefix ="Ctl2Litter_A"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =6000
                            Top =1080
                            Width =840
                            Height =240
                            Name ="2Litter_A_Label"
                            Caption ="litter"
                            EventProcPrefix ="Ctl2Litter_A_Label"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6000
                    Top =1620
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =22
                    Name ="4Litter_A"
                    ControlSource ="4Litter_A"
                    StatusBarText ="Litter depth at 4 meter point for transect A in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl4Litter_A"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6000
                    Top =1920
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =24
                    Name ="6Litter_A"
                    ControlSource ="6Litter_A"
                    StatusBarText ="Litter depth at 6 meter point for transect A in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl6Litter_A"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6000
                    Top =2220
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =26
                    Name ="8Litter_A"
                    ControlSource ="8Litter_A"
                    StatusBarText ="Litter depth at 8 meter point for transect A in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl8Litter_A"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6000
                    Top =2520
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =28
                    Name ="10Litter_A"
                    ControlSource ="10Litter_A"
                    StatusBarText ="Litter depth at 10 meter point for transect A in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl10Litter_A"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6000
                    Top =2820
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =30
                    Name ="12Litter_A"
                    ControlSource ="12Litter_A"
                    StatusBarText ="Litter depth at 12 meter point for transect A in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl12Litter_A"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6000
                    Top =3120
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =32
                    Name ="14Litter_A"
                    ControlSource ="14Litter_A"
                    StatusBarText ="Litter depth at 14 meter point for transect A in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl14Litter_A"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7680
                    Top =1320
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =34
                    Name ="2Litter_B"
                    ControlSource ="2Litter_B"
                    StatusBarText ="Litter depth at 2 meter point for transect B in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl2Litter_B"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =7680
                            Top =1080
                            Width =840
                            Height =240
                            Name ="2Litter_B_Label"
                            Caption ="litter"
                            EventProcPrefix ="Ctl2Litter_B_Label"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7680
                    Top =1620
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =36
                    Name ="4Litter_B"
                    ControlSource ="4Litter_B"
                    StatusBarText ="Litter depth at 4 meter point for transect B in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl4Litter_B"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7680
                    Top =1920
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =38
                    Name ="6Litter_B"
                    ControlSource ="6Litter_B"
                    StatusBarText ="Litter depth at 6 meter point for transect B in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl6Litter_B"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7680
                    Top =2220
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =40
                    Name ="8Litter_B"
                    ControlSource ="8Litter_B"
                    StatusBarText ="Litter depth at 8 meter point for transect B in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl8Litter_B"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7680
                    Top =2520
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =42
                    Name ="10Litter_B"
                    ControlSource ="10Litter_B"
                    StatusBarText ="Litter depth at 10 meter point for transect B in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl10Litter_B"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7680
                    Top =2820
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =44
                    Name ="12Litter_B"
                    ControlSource ="12Litter_B"
                    StatusBarText ="Litter depth at 12 meter point for transect B in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl12Litter_B"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7680
                    Top =3120
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =46
                    Name ="14Litter_B"
                    ControlSource ="14Litter_B"
                    StatusBarText ="Litter depth at 14 meter point for transect B in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl14Litter_B"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9360
                    Top =1320
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =48
                    Name ="2Litter_C"
                    ControlSource ="2Litter_C"
                    StatusBarText ="Litter depth at 2 meter point for transect C in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl2Litter_C"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =9360
                            Top =1080
                            Width =840
                            Height =240
                            Name ="2Litter_C_Label"
                            Caption ="litter"
                            EventProcPrefix ="Ctl2Litter_C_Label"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9360
                    Top =1620
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =50
                    Name ="4Litter_C"
                    ControlSource ="4Litter_C"
                    StatusBarText ="Litter depth at 4 meter point for transect C in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl4Litter_C"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9360
                    Top =1920
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =52
                    Name ="6Litter_C"
                    ControlSource ="6Litter_C"
                    StatusBarText ="Litter depth at 6 meter point for transect C in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl6Litter_C"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9360
                    Top =2220
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =54
                    Name ="8Litter_C"
                    ControlSource ="8Litter_C"
                    StatusBarText ="Litter depth at 8 meter point for transect C in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl8Litter_C"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9360
                    Top =2520
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =56
                    Name ="10Litter_C"
                    ControlSource ="10Litter_C"
                    StatusBarText ="Litter depth at 10 meter point for transect C in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl10Litter_C"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9360
                    Top =2820
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =58
                    Name ="12Litter_C"
                    ControlSource ="12Litter_C"
                    StatusBarText ="Litter depth at 12 meter point for transect C in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl12Litter_C"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9360
                    Top =3120
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =60
                    Name ="14Litter_C"
                    ControlSource ="14Litter_C"
                    StatusBarText ="Litter depth at 14 meter point for transect C in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl14Litter_C"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11040
                    Top =1320
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =62
                    Name ="Litter_D2"
                    ControlSource ="2Litter_D"
                    StatusBarText ="Litter depth at 2 meter point for transect D in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =11040
                            Top =1080
                            Width =840
                            Height =240
                            Name ="Litter_D_Label"
                            Caption ="litter"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11040
                    Top =1620
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =64
                    Name ="Litter_D4"
                    ControlSource ="4Litter_D"
                    StatusBarText ="Litter depth at 4 meter point for transect D in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11040
                    Top =1920
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =66
                    Name ="Litter_D6"
                    ControlSource ="6Litter_D"
                    StatusBarText ="Litter depth at 6 meter point for transect D in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11040
                    Top =2220
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =68
                    Name ="Litter_D8"
                    ControlSource ="8Litter_D"
                    StatusBarText ="Litter depth at 8 meter point for transect D in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11040
                    Top =2520
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =70
                    Name ="Litter_D10"
                    ControlSource ="10Litter_D"
                    StatusBarText ="Litter depth at 10 meter point for transect D in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11040
                    Top =2820
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =72
                    Name ="Litter_D12"
                    ControlSource ="12Litter_D"
                    StatusBarText ="Litter depth at 12 meter point for transect D in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11040
                    Top =3120
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =74
                    Name ="Litter_D14"
                    ControlSource ="14Litter_D"
                    StatusBarText ="Litter depth at 14 meter point for transect D in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6840
                    Top =1320
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =21
                    Name ="2Duff_A"
                    ControlSource ="2Duff_A"
                    StatusBarText ="Duff depth at 2 meter point for transect A in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl2Duff_A"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =223
                            Left =6840
                            Top =1080
                            Width =840
                            Height =240
                            Name ="2Duff_A_Label"
                            Caption ="duff"
                            EventProcPrefix ="Ctl2Duff_A_Label"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6840
                    Top =1620
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =23
                    Name ="4Duff_A"
                    ControlSource ="4Duff_A"
                    StatusBarText ="Duff depth at 4 meter point for transect A in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl4Duff_A"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6840
                    Top =1920
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =25
                    Name ="6Duff_A"
                    ControlSource ="6Duff_A"
                    StatusBarText ="Duff depth at 6 meter point for transect A in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl6Duff_A"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6840
                    Top =2220
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =27
                    Name ="8Duff_A"
                    ControlSource ="8Duff_A"
                    StatusBarText ="Duff depth at 8 meter point for transect A in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl8Duff_A"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6840
                    Top =2520
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =29
                    Name ="10Duff_A"
                    ControlSource ="10Duff_A"
                    StatusBarText ="Duff depth at 10 meter point for transect A in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl10Duff_A"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6840
                    Top =2820
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =31
                    Name ="12Duff_A"
                    ControlSource ="12Duff_A"
                    StatusBarText ="Duff depth at 12 meter point for transect A in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl12Duff_A"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6840
                    Top =3120
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =33
                    Name ="14Duff_A"
                    ControlSource ="14Duff_A"
                    StatusBarText ="Duff depth at 14 meter point for transect A in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl14Duff_A"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8520
                    Top =1320
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =35
                    Name ="2Duff_B"
                    ControlSource ="2Duff_B"
                    StatusBarText ="Duff depth at 2 meter point for transect B in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl2Duff_B"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =223
                            Left =8520
                            Top =1080
                            Width =840
                            Height =240
                            Name ="2Duff_B_Label"
                            Caption ="duff"
                            EventProcPrefix ="Ctl2Duff_B_Label"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8520
                    Top =1620
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =37
                    Name ="4Duff_B"
                    ControlSource ="4Duff_B"
                    StatusBarText ="Duff depth at 4 meter point for transect B in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl4Duff_B"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8520
                    Top =1920
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =39
                    Name ="6Duff_B"
                    ControlSource ="6Duff_B"
                    StatusBarText ="Duff depth at 6 meter point for transect B in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl6Duff_B"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8520
                    Top =2220
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =41
                    Name ="8Duff_B"
                    ControlSource ="8Duff_B"
                    StatusBarText ="Duff depth at 8 meter point for transect B in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl8Duff_B"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8520
                    Top =2520
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =43
                    Name ="10Duff_B"
                    ControlSource ="10Duff_B"
                    StatusBarText ="Duff depth at 10 meter point for transect B in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl10Duff_B"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8520
                    Top =2820
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =45
                    Name ="12Duff_B"
                    ControlSource ="12Duff_B"
                    StatusBarText ="Duff depth at 12 meter point for transect B in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl12Duff_B"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8520
                    Top =3120
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =47
                    Name ="14Duff_B"
                    ControlSource ="14Duff_B"
                    StatusBarText ="Duff depth at 14 meter point for transect B in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl14Duff_B"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10200
                    Top =1320
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =49
                    Name ="2Duff_C"
                    ControlSource ="2Duff_C"
                    StatusBarText ="Duff depth at 2 meter point for transect C in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl2Duff_C"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =223
                            Left =10200
                            Top =1080
                            Width =840
                            Height =240
                            Name ="2Duff_C_Label"
                            Caption ="duff"
                            EventProcPrefix ="Ctl2Duff_C_Label"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10200
                    Top =1620
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =51
                    Name ="4Duff_C"
                    ControlSource ="4Duff_C"
                    StatusBarText ="Duff depth at 4 meter point for transect C in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl4Duff_C"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10200
                    Top =1920
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =53
                    Name ="6Duff_C"
                    ControlSource ="6Duff_C"
                    StatusBarText ="Duff depth at 6 meter point for transect C in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl6Duff_C"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10200
                    Top =2220
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =55
                    Name ="8Duff_C"
                    ControlSource ="8Duff_C"
                    StatusBarText ="Duff depth at 8 meter point for transect C in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl8Duff_C"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10200
                    Top =2520
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =57
                    Name ="10Duff_C"
                    ControlSource ="10Duff_C"
                    StatusBarText ="Duff depth at 10 meter point for transect C in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl10Duff_C"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10200
                    Top =2820
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =59
                    Name ="12Duff_C"
                    ControlSource ="12Duff_C"
                    StatusBarText ="Duff depth at 12 meter point for transect C in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl12Duff_C"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10200
                    Top =3120
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =61
                    Name ="14Duff_C"
                    ControlSource ="14Duff_C"
                    StatusBarText ="Duff depth at 14 meter point for transect C in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl14Duff_C"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11880
                    Top =1320
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =63
                    Name ="Duff_D2"
                    ControlSource ="2Duff_D"
                    StatusBarText ="Duff depth at 2 meter point for transect D in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            Left =11880
                            Top =1080
                            Width =840
                            Height =240
                            Name ="Duff_D_Label"
                            Caption ="duff"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11880
                    Top =1620
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =65
                    Name ="Duff_D4"
                    ControlSource ="4Duff_D"
                    StatusBarText ="Duff depth at 4 meter point for transect D in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11880
                    Top =1920
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =67
                    Name ="Duff_D6"
                    ControlSource ="6Duff_D"
                    StatusBarText ="Duff depth at 6 meter point for transect D in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11880
                    Top =2220
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =69
                    Name ="Duff_D8"
                    ControlSource ="8Duff_D"
                    StatusBarText ="Duff depth at 8 meter point for transect D in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11880
                    Top =2520
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =71
                    Name ="Duff_D10"
                    ControlSource ="10Duff_D"
                    StatusBarText ="Duff depth at 10 meter point for transect D in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11880
                    Top =2820
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =73
                    Name ="Duff_D12"
                    ControlSource ="12Duff_D"
                    StatusBarText ="Duff depth at 12 meter point for transect D in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11880
                    Top =3120
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =75
                    Name ="Duff_D14"
                    ControlSource ="14Duff_D"
                    StatusBarText ="Duff depth at 14 meter point for transect D in centimeters"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    Left =6000
                    Top =840
                    Width =1680
                    Height =240
                    FontSize =10
                    Name ="Label117"
                    Caption ="A"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    Left =7680
                    Top =840
                    Width =1680
                    Height =240
                    FontSize =10
                    Name ="Label118"
                    Caption ="B"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    Left =9360
                    Top =840
                    Width =1680
                    Height =240
                    FontSize =10
                    Name ="Label119"
                    Caption ="C"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    Left =11040
                    Top =840
                    Width =1680
                    Height =240
                    FontSize =10
                    Name ="Label120"
                    Caption ="D"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    Left =6000
                    Top =600
                    Width =6720
                    Height =240
                    Name ="Label121"
                    Caption ="depth (cm)"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =223
                    Left =5100
                    Top =600
                    Width =900
                    Height =720
                    Name ="Label122"
                    Caption ="point on transect (m)"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =223
                    Left =5100
                    Top =1320
                    Width =900
                    Height =300
                    Name ="Label123"
                    Caption ="2"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =223
                    Left =5100
                    Top =1620
                    Width =900
                    Height =300
                    Name ="Label124"
                    Caption ="4"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =223
                    Left =5100
                    Top =1920
                    Width =900
                    Height =300
                    Name ="Label125"
                    Caption ="6"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =223
                    Left =5100
                    Top =2220
                    Width =900
                    Height =300
                    Name ="Label126"
                    Caption ="8"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =223
                    Left =5100
                    Top =2520
                    Width =900
                    Height =300
                    Name ="Label127"
                    Caption ="10"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =223
                    Left =5100
                    Top =2820
                    Width =900
                    Height =300
                    Name ="Label128"
                    Caption ="12"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =215
                    Left =5100
                    Top =3120
                    Width =900
                    Height =300
                    Name ="Label130"
                    Caption ="14"
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =0
                    Left =5100
                    Top =300
                    Width =2220
                    Height =300
                    FontSize =10
                    Name ="Label131"
                    Caption ="Litter and Duff Depth"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =120
                    Top =120
                    Width =330
                    Height =300
                    TabIndex =76
                    Name ="Fuels_ID"
                    ControlSource ="Fuels_ID"
                    StatusBarText ="Unique record identifier - primary key"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =540
                    Top =120
                    Width =330
                    Height =300
                    TabIndex =77
                    Name ="Event_ID"
                    ControlSource ="Event_ID"
                    StatusBarText ="Foreign key to tbl_Events"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1485
                    Top =3660
                    Width =810
                    Height =300
                    Name ="Bearing_A"
                    StatusBarText ="Bearing of the plot slope + 180 in degrees"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =1485
                            Top =3420
                            Width =810
                            Height =240
                            FontWeight =400
                            Name ="Bearing_A_Label"
                            Caption ="A"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =255
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2280
                    Top =3660
                    Width =810
                    Height =300
                    TabIndex =2
                    Name ="Bearing_B"
                    StatusBarText ="Bearing of transect 1 in degrees"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =223
                            Left =2280
                            Top =3420
                            Width =810
                            Height =240
                            FontWeight =400
                            Name ="Bearing_B_Label"
                            Caption ="B"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =127
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3105
                    Top =3660
                    Width =810
                    Height =300
                    TabIndex =4
                    Name ="Bearing_C"
                    StatusBarText ="Bearing of transect 3 + 180"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =3105
                            Top =3420
                            Width =810
                            Height =240
                            FontWeight =400
                            Name ="Bearing_C_Label"
                            Caption ="C"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3945
                    Top =3660
                    Width =810
                    Height =300
                    TabIndex =6
                    Name ="Bearing_D"
                    StatusBarText ="Bearing of the plot slope"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =3945
                            Top =3420
                            Width =810
                            Height =240
                            FontWeight =400
                            Name ="Bearing_D_Label"
                            Caption ="D"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =1
                    OverlapFlags =127
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1485
                    Top =3960
                    Width =810
                    Height =300
                    TabIndex =1
                    Name ="Slope_A"
                    StatusBarText ="Slope of transect A to nearest half percent"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            Left =180
                            Top =3660
                            Width =1290
                            Height =300
                            FontWeight =400
                            Name ="Slope_A_Label"
                            Caption ="bearing (deg)"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =1
                    OverlapFlags =255
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2280
                    Top =3960
                    Width =810
                    Height =300
                    TabIndex =3
                    Name ="Slope_B"
                    StatusBarText ="Slope of transect B to nearest half percent"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =119
                            Left =180
                            Top =3960
                            Width =1290
                            Height =300
                            FontWeight =400
                            Name ="Slope_B_Label"
                            Caption ="slope (%)"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =1
                    OverlapFlags =119
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3105
                    Top =3960
                    Width =810
                    Height =300
                    TabIndex =5
                    Name ="Slope_C"
                    StatusBarText ="Slope of transect C to nearest half percent"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =1
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3945
                    Top =3960
                    Width =810
                    Height =300
                    TabIndex =7
                    Name ="Slope_D"
                    StatusBarText ="Slope of transect C to nearest half percent"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1440
                    Top =1020
                    Width =1080
                    Height =300
                    TabIndex =8
                    Name ="1HR_A"
                    ControlSource ="1HR_A"
                    StatusBarText ="One hour fuel intercept for transect A"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl1HR_A"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =1440
                            Top =600
                            Width =1080
                            Height =420
                            Name ="1HR_A_Label"
                            Caption ="    1-hr      (0-0.25 in)"
                            EventProcPrefix ="Ctl1HR_A_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1440
                    Top =1320
                    Width =1080
                    Height =300
                    TabIndex =11
                    Name ="1HR_B"
                    ControlSource ="1HR_B"
                    StatusBarText ="One hour fuel intercept for transect B"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl1HR_B"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =223
                            Left =540
                            Top =600
                            Width =900
                            Height =420
                            Name ="1HR_B_Label"
                            Caption ="transect"
                            EventProcPrefix ="Ctl1HR_B_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1440
                    Top =1620
                    Width =1080
                    Height =300
                    TabIndex =14
                    Name ="1HR_C"
                    ControlSource ="1HR_C"
                    StatusBarText ="One hour fuel intercept for transect C"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl1HR_C"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =223
                            Left =540
                            Top =1020
                            Width =900
                            Height =300
                            Name ="1HR_C_Label"
                            Caption ="A"
                            EventProcPrefix ="Ctl1HR_C_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1440
                    Top =1920
                    Width =1080
                    Height =300
                    TabIndex =17
                    Name ="DI_1HR"
                    ControlSource ="1HR_D"
                    StatusBarText ="One hour fuel intercept for transect D"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2520
                    Top =1020
                    Width =1080
                    Height =300
                    TabIndex =9
                    Name ="10HR_A"
                    ControlSource ="10HR_A"
                    StatusBarText ="Ten hour fuel intercept for transect A"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl10HR_A"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            Left =2520
                            Top =600
                            Width =1080
                            Height =420
                            Name ="10HR_A_Label"
                            Caption ="     10-hr     (0.25-1 in)"
                            EventProcPrefix ="Ctl10HR_A_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2520
                    Top =1320
                    Width =1080
                    Height =300
                    TabIndex =12
                    Name ="10HR_B"
                    ControlSource ="10HR_B"
                    StatusBarText ="Ten hour fuel intercept for transect B"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl10HR_B"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =223
                            Left =540
                            Top =1620
                            Width =900
                            Height =300
                            Name ="10HR_B_Label"
                            Caption ="C"
                            EventProcPrefix ="Ctl10HR_B_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2520
                    Top =1620
                    Width =1080
                    Height =300
                    TabIndex =15
                    Name ="10HR_C"
                    ControlSource ="10HR_C"
                    StatusBarText ="Ten hour fuel intercept for transect C"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl10HR_C"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =215
                            Left =540
                            Top =1920
                            Width =900
                            Height =299
                            Name ="LabelD10"
                            Caption ="D"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2520
                    Top =1920
                    Width =1080
                    Height =300
                    TabIndex =18
                    Name ="DI_10HR"
                    ControlSource ="10HR_D"
                    StatusBarText ="Ten hour fuel intercept for transect D"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3600
                    Top =1020
                    Width =1080
                    Height =300
                    TabIndex =10
                    Name ="100HR_A"
                    ControlSource ="100HR_A"
                    StatusBarText ="Hundred hour fuel intercept for transect A"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl100HR_A"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            Left =3600
                            Top =600
                            Width =1080
                            Height =420
                            Name ="100HR_A_Label"
                            Caption ="     100-hr    (1-3 in)"
                            EventProcPrefix ="Ctl100HR_A_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3600
                    Top =1320
                    Width =1080
                    Height =300
                    TabIndex =13
                    Name ="100HR_B"
                    ControlSource ="100HR_B"
                    StatusBarText ="Hundred hour fuel intercept for transect B"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl100HR_B"

                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3600
                    Top =1620
                    Width =1080
                    Height =300
                    TabIndex =16
                    Name ="100HR_C"
                    ControlSource ="100HR_C"
                    StatusBarText ="Hundred hour fuel intercept for transect C"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl100HR_C"

                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3600
                    Top =1920
                    Width =1080
                    Height =300
                    TabIndex =19
                    Name ="DI_100HR"
                    ControlSource ="100HR_D"
                    StatusBarText ="Hundred hour fuel intercept for transect D"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin Label
                    OverlapFlags =85
                    Left =1620
                    Top =3120
                    Width =1440
                    Height =240
                    Name ="Label45"
                    Caption ="fuels transect"
                End
                Begin Label
                    OverlapFlags =85
                    Left =1560
                    Top =300
                    Width =1440
                    Height =240
                    Name ="Label46"
                    Caption ="# of intercepts"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =215
                    Left =540
                    Top =1320
                    Width =900
                    Height =300
                    Name ="Label133"
                    Caption ="B"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =900
                    Top =2280
                    Width =606
                    Height =288
                    TabIndex =78
                    Name ="ButtonA1"
                    Caption ="+ 1"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    OnKeyDown ="[Event Procedure]"

                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1620
                    Top =2280
                    Width =606
                    Height =288
                    TabIndex =79
                    Name ="ButtonA5"
                    Caption ="+ 5"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    OnKeyDown ="[Event Procedure]"

                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2340
                    Top =2280
                    Width =606
                    Height =288
                    TabIndex =80
                    Name ="ButtonS1"
                    Caption ="- 1"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    OnKeyDown ="[Event Procedure]"

                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3060
                    Top =2280
                    Width =606
                    Height =288
                    TabIndex =81
                    Name ="ButtonS5"
                    Caption ="- 5"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    OnKeyDown ="[Event Procedure]"

                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3780
                    Top =2280
                    Width =606
                    Height =288
                    TabIndex =82
                    Name ="ButtonZero"
                    Caption ="0"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    OnKeyDown ="[Event Procedure]"

                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4920
                    Top =3840
                    Width =1185
                    Height =300
                    TabIndex =83
                    Name ="ButtonTransect"
                    Caption ="Edit Transect"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
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
' MODULE:       Form_fsub_Fuels_LD
' Level:        Form module
' Version:      1.01
' Description:  data functions & procedures specific to adding litter & duff fuels data
'
' Source/date:  Russ DenBleyker, unknown
' Revisions:    RDB - unknown  - 1.00 - initial version
'               BLC - 3/16/2016 - 1.01 - added documentation, litter & duff change & form load
'                                        subroutines for handling litter & duff value highlighting (backcolor = #FFF200, yellow)
' =================================

' ---------------------------------
' SUB:          Form_Load
' Description:  handles form loading actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Russ DenBleyker, unknown
' Revisions:
'       RDB, unknown - initial version
'       BLC, 3/16/2016 - added error handling & documentation
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim Locations As DAO.Recordset
    Dim strSQL As String
    
    ' Set up necessary fields
    Set db = CurrentDb
    strSQL = "SELECT * FROM [tbl_Locations] WHERE [Location_ID] = '" & Me.Parent!Location_ID & "'"
    Set Locations = db.OpenRecordset(strSQL)
    
    Me!Bearing_A = Locations!Bearing_A
    Me!Bearing_B = Locations!Bearing_B
    Me!Bearing_C = Locations!Bearing_C
    Me!Bearing_D = Locations!Bearing_D
    Me!Slope_A = Locations!Slope_A
    Me!Slope_B = Locations!Slope_B
    Me!Slope_C = Locations!Slope_C
    Me!Slope_D = Locations!Slope_D
    
    'set values for oak scrub
    If Not IsNull(Locations!Vegetation_Type) And Locations!Vegetation_Type = "oak scrub" Then
      Me!Bearing_D_Label.Visible = False
      Me!Bearing_D.Visible = False
      Me!Slope_D.Visible = False
      Me!DI_1HR.Visible = False
      Me!DI_10HR.Visible = False
      Me!DI_100HR.Visible = False
      Me!LabelD10.Visible = False
      Me!Label120.Caption = " "
      Me!Litter_D_Label.Caption = " "
      Me!Duff_D_Label.Caption = " "
      Me!Duff_D2.Enabled = False
      Me!Duff_D2.Locked = True
      Me!Litter_D2.Enabled = False
      Me!Litter_D2.Locked = True
      Me!Duff_D4.Enabled = False
      Me!Duff_D4.Locked = True
      Me!Litter_D4.Enabled = False
      Me!Litter_D4.Locked = True
      Me!Duff_D6.Enabled = False
      Me!Duff_D6.Locked = True
      Me!Litter_D6.Enabled = False
      Me!Litter_D6.Locked = True
      Me!Duff_D8.Enabled = False
      Me!Duff_D8.Locked = True
      Me!Litter_D8.Enabled = False
      Me!Litter_D8.Locked = True
      Me!Duff_D10.Enabled = False
      Me!Duff_D10.Locked = True
      Me!Litter_D10.Enabled = False
      Me!Litter_D10.Locked = True
      Me!Duff_D12.Enabled = False
      Me!Duff_D12.Locked = True
      Me!Litter_D12.Enabled = False
      Me!Litter_D12.Locked = True
      Me!Duff_D14.Enabled = False
      Me!Duff_D14.Locked = True
      Me!Litter_D14.Enabled = False
      Me!Litter_D14.Locked = True
    End If
    Locations.Close
    Set Locations = Nothing

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub



'=========================
' Litter & Duff Changes
'=========================

' ---------------------------------
' SUB:          Ctl2Litter_A_Change
' Description:  handles ctl2Litter_A actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Ctl2Litter_A_Change()
On Error GoTo Err_Handler

    'clear highlight if not null
    If Not IsNull(Ctl2Litter_A) Then
        Ctl2Litter_A.BackColor = RGB(255, 255, 255)
    End If
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Ctl2Litter_A_Change[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub


' ---------------------------------
' SUB:          Form_KeyDown
' Description:  handles form's key down actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, August 2015
' Revisions:    BLC, 8/21/2014 - initial version
' ---------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Handler

    'capture ESC & let user determine if fields should be cleared
    CaptureEscapeKey KeyCode
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_KeyDown[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

'======================
' KeyDown Events
'======================

Private Sub Bearing_A_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Bearing_B_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Bearing_C_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Bearing_D_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub ButtonA1_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub ButtonA5_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub ButtonS1_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub ButtonS5_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub ButtonZero_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl100HR_A_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl100HR_B_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl100HR_C_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl10Duff_A_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl10Duff_B_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl10Duff_C_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl10HR_A_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl10HR_B_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl10HR_C_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl10Litter_A_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl10Litter_B_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl10Litter_C_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl12Duff_A_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl12Duff_B_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl12Duff_C_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl12Litter_A_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl12Litter_B_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl12Litter_C_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl14Duff_A_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl14Duff_B_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl14Duff_C_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl14Litter_A_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl14Litter_B_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl14Litter_C_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl1HR_A_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl1HR_B_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl1HR_C_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl2Duff_A_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl2Duff_B_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl2Duff_C_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl2Litter_A_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl2Litter_B_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl2Litter_C_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl4Duff_A_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl4Duff_B_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl4Duff_C_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl4Litter_A_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl4Litter_B_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl4Litter_C_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl6Duff_A_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl6Duff_B_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl6Duff_C_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl6Litter_A_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl6Litter_B_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl6Litter_C_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl8Duff_A_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl8Duff_B_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl8Duff_C_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl8Litter_A_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl8Litter_B_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Ctl8Litter_C_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub DI_100HR_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub DI_10HR_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub DI_1HR_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Duff_D10_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Duff_D12_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Duff_D14_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Duff_D2_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Duff_D4_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Duff_D6_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Duff_D8_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub


'======================
' Click Events
'======================
Private Sub ButtonA1_Click()
  If InStr(1, Screen.PreviousControl.name, "HR") > 0 Then
    If IsNull(Screen.PreviousControl.Value) Then
      Screen.PreviousControl.Value = 1
    Else
      Screen.PreviousControl.Value = Screen.PreviousControl.Value + 1
    End If
  End If
  Screen.PreviousControl.SetFocus
End Sub

Private Sub ButtonA5_Click()
  If InStr(1, Screen.PreviousControl.name, "HR") > 0 Then
    If IsNull(Screen.PreviousControl.Value) Then
      Screen.PreviousControl.Value = 5
    Else
      Screen.PreviousControl.Value = Screen.PreviousControl.Value + 5
    End If
  End If
  Screen.PreviousControl.SetFocus
End Sub

Private Sub ButtonS1_Click()
  If InStr(1, Screen.PreviousControl.name, "HR") > 0 Then
    If IsNull(Screen.PreviousControl.Value) Then
      Screen.PreviousControl.Value = 0
    ElseIf Screen.PreviousControl.Value - 1 < 0 Then
      MsgBox "Total cannot be negative.", , "Fuels Intercepts"
      Exit Sub
    Else
      Screen.PreviousControl.Value = Screen.PreviousControl.Value - 1
    End If
  End If
  Screen.PreviousControl.SetFocus
End Sub

Private Sub ButtonS5_Click()
  If InStr(1, Screen.PreviousControl.name, "HR") > 0 Then
    If IsNull(Screen.PreviousControl.Value) Then
      Screen.PreviousControl.Value = 0
    ElseIf Screen.PreviousControl.Value - 5 < 0 Then
      MsgBox "Total cannot be negative.", , "Fuels Intercepts"
      Exit Sub
    Else
      Screen.PreviousControl.Value = Screen.PreviousControl.Value - 5
    End If
  End If
  Screen.PreviousControl.SetFocus
End Sub

Private Sub ButtonZero_Click()
  If InStr(1, Screen.PreviousControl.name, "HR") > 0 Then
    Screen.PreviousControl.Value = 0
  End If
  Screen.PreviousControl.SetFocus
End Sub

Private Sub Litter_D10_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Litter_D12_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Litter_D14_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Litter_D2_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Litter_D4_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Litter_D6_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Litter_D8_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Slope_A_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Slope_B_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Slope_C_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Slope_D_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub ButtonTransect_Click()
On Error GoTo Err_ButtonTransect_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim db As DAO.Database
    Dim Locations As DAO.Recordset
    Dim strSQL As String

    stDocName = "frm_Edit_Fuel_Transect"
    
    stLinkCriteria = "[Location_ID]=" & "'" & Me.Parent!Location_ID & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , acDialog

    ' Set up changed fields
    Set db = CurrentDb
    strSQL = "SELECT * FROM [tbl_Locations] WHERE [Location_ID] = '" & Me.Parent!Location_ID & "'"
    Set Locations = db.OpenRecordset(strSQL)
    Me!Bearing_A = Locations!Bearing_A
    Me!Bearing_B = Locations!Bearing_B
    Me!Bearing_C = Locations!Bearing_C
    Me!Bearing_D = Locations!Bearing_D
    Me!Slope_A = Locations!Slope_A
    Me!Slope_B = Locations!Slope_B
    Me!Slope_C = Locations!Slope_C
    Me!Slope_D = Locations!Slope_D
    Locations.Close
    Set Locations = Nothing
    Me.Requery

Exit_ButtonTransect_Click:
    Exit Sub

Err_ButtonTransect_Click:
    MsgBox Err.Description
    Resume Exit_ButtonTransect_Click
    
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    On Error GoTo Err_Handler
    If IsNull(Me!Event_ID) Then
      MsgBox "You must enter event information first."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
      GoTo Exit_Procedure
    End If
    ' Create the GUID primary key value
    If IsNull(Me!Fuels_ID) Then
        If GetDataType("tbl_Fuels", "Fuels_ID") = dbText Then
            Me!Fuels_ID = fxnGUIDGen
        End If
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub
