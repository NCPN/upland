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
    Top =3270
    Right =9030
    Bottom =7665
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
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6000
                    Top =1320
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =13
                    BackColor =65535
                    Name ="Litter_A2"
                    ControlSource ="2Litter_A"
                    StatusBarText ="Litter depth at 2 meter point for transect A in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =6000
                    LayoutCachedTop =1320
                    LayoutCachedWidth =6840
                    LayoutCachedHeight =1620
                    BorderThemeColorIndex =0
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
                            LayoutCachedLeft =6000
                            LayoutCachedTop =1080
                            LayoutCachedWidth =6840
                            LayoutCachedHeight =1320
                            BorderThemeColorIndex =0
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6000
                    Top =1620
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =15
                    BackColor =65535
                    Name ="Litter_A4"
                    ControlSource ="4Litter_A"
                    StatusBarText ="Litter depth at 4 meter point for transect A in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =6000
                    LayoutCachedTop =1620
                    LayoutCachedWidth =6840
                    LayoutCachedHeight =1920
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6000
                    Top =1920
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =17
                    BackColor =65535
                    Name ="Litter_A6"
                    ControlSource ="6Litter_A"
                    StatusBarText ="Litter depth at 6 meter point for transect A in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =6000
                    LayoutCachedTop =1920
                    LayoutCachedWidth =6840
                    LayoutCachedHeight =2220
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6000
                    Top =2220
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =19
                    BackColor =65535
                    Name ="Litter_A8"
                    ControlSource ="8Litter_A"
                    StatusBarText ="Litter depth at 8 meter point for transect A in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =6000
                    LayoutCachedTop =2220
                    LayoutCachedWidth =6840
                    LayoutCachedHeight =2520
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6000
                    Top =2520
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =21
                    BackColor =65535
                    Name ="Litter_A10"
                    ControlSource ="10Litter_A"
                    StatusBarText ="Litter depth at 10 meter point for transect A in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =6000
                    LayoutCachedTop =2520
                    LayoutCachedWidth =6840
                    LayoutCachedHeight =2820
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6000
                    Top =2820
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =23
                    BackColor =65535
                    Name ="Litter_A12"
                    ControlSource ="12Litter_A"
                    StatusBarText ="Litter depth at 12 meter point for transect A in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =6000
                    LayoutCachedTop =2820
                    LayoutCachedWidth =6840
                    LayoutCachedHeight =3120
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6000
                    Top =3120
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =25
                    BackColor =65535
                    Name ="Litter_A14"
                    ControlSource ="14Litter_A"
                    StatusBarText ="Litter depth at 14 meter point for transect A in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =6000
                    LayoutCachedTop =3120
                    LayoutCachedWidth =6840
                    LayoutCachedHeight =3420
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7680
                    Top =1320
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =27
                    BackColor =65535
                    Name ="Litter_B2"
                    ControlSource ="2Litter_B"
                    StatusBarText ="Litter depth at 2 meter point for transect B in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =7680
                    LayoutCachedTop =1320
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =1620
                    BorderThemeColorIndex =0
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
                            LayoutCachedLeft =7680
                            LayoutCachedTop =1080
                            LayoutCachedWidth =8520
                            LayoutCachedHeight =1320
                            BorderThemeColorIndex =0
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7680
                    Top =1620
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =29
                    BackColor =65535
                    Name ="Litter_B4"
                    ControlSource ="4Litter_B"
                    StatusBarText ="Litter depth at 4 meter point for transect B in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =7680
                    LayoutCachedTop =1620
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =1920
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7680
                    Top =1920
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =31
                    BackColor =65535
                    Name ="Litter_B6"
                    ControlSource ="6Litter_B"
                    StatusBarText ="Litter depth at 6 meter point for transect B in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =7680
                    LayoutCachedTop =1920
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =2220
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7680
                    Top =2220
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =33
                    BackColor =65535
                    Name ="Litter_B8"
                    ControlSource ="8Litter_B"
                    StatusBarText ="Litter depth at 8 meter point for transect B in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =7680
                    LayoutCachedTop =2220
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =2520
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7680
                    Top =2520
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =35
                    BackColor =65535
                    Name ="Litter_B10"
                    ControlSource ="10Litter_B"
                    StatusBarText ="Litter depth at 10 meter point for transect B in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =7680
                    LayoutCachedTop =2520
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =2820
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7680
                    Top =2820
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =37
                    BackColor =65535
                    Name ="Litter_B12"
                    ControlSource ="12Litter_B"
                    StatusBarText ="Litter depth at 12 meter point for transect B in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =7680
                    LayoutCachedTop =2820
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =3120
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7680
                    Top =3120
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =39
                    BackColor =65535
                    Name ="Litter_B14"
                    ControlSource ="14Litter_B"
                    StatusBarText ="Litter depth at 14 meter point for transect B in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =7680
                    LayoutCachedTop =3120
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =3420
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9360
                    Top =1320
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =41
                    BackColor =65535
                    Name ="Litter_C2"
                    ControlSource ="2Litter_C"
                    StatusBarText ="Litter depth at 2 meter point for transect C in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =9360
                    LayoutCachedTop =1320
                    LayoutCachedWidth =10200
                    LayoutCachedHeight =1620
                    BorderThemeColorIndex =0
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
                            LayoutCachedLeft =9360
                            LayoutCachedTop =1080
                            LayoutCachedWidth =10200
                            LayoutCachedHeight =1320
                            BorderThemeColorIndex =0
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9360
                    Top =1620
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =43
                    BackColor =65535
                    Name ="Litter_C4"
                    ControlSource ="4Litter_C"
                    StatusBarText ="Litter depth at 4 meter point for transect C in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =9360
                    LayoutCachedTop =1620
                    LayoutCachedWidth =10200
                    LayoutCachedHeight =1920
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9360
                    Top =1920
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =45
                    BackColor =65535
                    Name ="Litter_C6"
                    ControlSource ="6Litter_C"
                    StatusBarText ="Litter depth at 6 meter point for transect C in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =9360
                    LayoutCachedTop =1920
                    LayoutCachedWidth =10200
                    LayoutCachedHeight =2220
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9360
                    Top =2220
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =47
                    BackColor =65535
                    Name ="Litter_C8"
                    ControlSource ="8Litter_C"
                    StatusBarText ="Litter depth at 8 meter point for transect C in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =9360
                    LayoutCachedTop =2220
                    LayoutCachedWidth =10200
                    LayoutCachedHeight =2520
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9360
                    Top =2520
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =49
                    BackColor =65535
                    Name ="Litter_C10"
                    ControlSource ="10Litter_C"
                    StatusBarText ="Litter depth at 10 meter point for transect C in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =9360
                    LayoutCachedTop =2520
                    LayoutCachedWidth =10200
                    LayoutCachedHeight =2820
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9360
                    Top =2820
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =51
                    BackColor =65535
                    Name ="Litter_C12"
                    ControlSource ="12Litter_C"
                    StatusBarText ="Litter depth at 12 meter point for transect C in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =9360
                    LayoutCachedTop =2820
                    LayoutCachedWidth =10200
                    LayoutCachedHeight =3120
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9360
                    Top =3120
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =53
                    BackColor =65535
                    Name ="Litter_C14"
                    ControlSource ="14Litter_C"
                    StatusBarText ="Litter depth at 14 meter point for transect C in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =9360
                    LayoutCachedTop =3120
                    LayoutCachedWidth =10200
                    LayoutCachedHeight =3420
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11040
                    Top =1320
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =55
                    BackColor =65535
                    Name ="Litter_D2"
                    ControlSource ="2Litter_D"
                    StatusBarText ="Litter depth at 2 meter point for transect D in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =11040
                    LayoutCachedTop =1320
                    LayoutCachedWidth =11880
                    LayoutCachedHeight =1620
                    BorderThemeColorIndex =0
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
                            LayoutCachedLeft =11040
                            LayoutCachedTop =1080
                            LayoutCachedWidth =11880
                            LayoutCachedHeight =1320
                            BorderThemeColorIndex =0
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11040
                    Top =1620
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =57
                    BackColor =65535
                    Name ="Litter_D4"
                    ControlSource ="4Litter_D"
                    StatusBarText ="Litter depth at 4 meter point for transect D in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =11040
                    LayoutCachedTop =1620
                    LayoutCachedWidth =11880
                    LayoutCachedHeight =1920
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11040
                    Top =1920
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =59
                    BackColor =65535
                    Name ="Litter_D6"
                    ControlSource ="6Litter_D"
                    StatusBarText ="Litter depth at 6 meter point for transect D in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =11040
                    LayoutCachedTop =1920
                    LayoutCachedWidth =11880
                    LayoutCachedHeight =2220
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11040
                    Top =2220
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =61
                    BackColor =65535
                    Name ="Litter_D8"
                    ControlSource ="8Litter_D"
                    StatusBarText ="Litter depth at 8 meter point for transect D in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =11040
                    LayoutCachedTop =2220
                    LayoutCachedWidth =11880
                    LayoutCachedHeight =2520
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11040
                    Top =2520
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =63
                    BackColor =65535
                    Name ="Litter_D10"
                    ControlSource ="10Litter_D"
                    StatusBarText ="Litter depth at 10 meter point for transect D in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =11040
                    LayoutCachedTop =2520
                    LayoutCachedWidth =11880
                    LayoutCachedHeight =2820
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11040
                    Top =2820
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =65
                    BackColor =65535
                    Name ="Litter_D12"
                    ControlSource ="12Litter_D"
                    StatusBarText ="Litter depth at 12 meter point for transect D in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =11040
                    LayoutCachedTop =2820
                    LayoutCachedWidth =11880
                    LayoutCachedHeight =3120
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11040
                    Top =3120
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =67
                    BackColor =65535
                    Name ="Litter_D14"
                    ControlSource ="14Litter_D"
                    StatusBarText ="Litter depth at 14 meter point for transect D in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =11040
                    LayoutCachedTop =3120
                    LayoutCachedWidth =11880
                    LayoutCachedHeight =3420
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6840
                    Top =1320
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =14
                    BackColor =65535
                    Name ="Duff_A2"
                    ControlSource ="2Duff_A"
                    StatusBarText ="Duff depth at 2 meter point for transect A in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =6840
                    LayoutCachedTop =1320
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =1620
                    BorderThemeColorIndex =0
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =95
                            Left =6840
                            Top =1080
                            Width =840
                            Height =240
                            BackColor =10079487
                            Name ="2Duff_A_Label"
                            Caption ="duff"
                            EventProcPrefix ="Ctl2Duff_A_Label"
                            LayoutCachedLeft =6840
                            LayoutCachedTop =1080
                            LayoutCachedWidth =7680
                            LayoutCachedHeight =1320
                            BorderThemeColorIndex =0
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6840
                    Top =1620
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =16
                    BackColor =65535
                    Name ="Duff_A4"
                    ControlSource ="4Duff_A"
                    StatusBarText ="Duff depth at 4 meter point for transect A in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =6840
                    LayoutCachedTop =1620
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =1920
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6840
                    Top =1920
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =18
                    BackColor =65535
                    Name ="Duff_A6"
                    ControlSource ="6Duff_A"
                    StatusBarText ="Duff depth at 6 meter point for transect A in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =6840
                    LayoutCachedTop =1920
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =2220
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6840
                    Top =2220
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =20
                    BackColor =65535
                    Name ="Duff_A8"
                    ControlSource ="8Duff_A"
                    StatusBarText ="Duff depth at 8 meter point for transect A in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =6840
                    LayoutCachedTop =2220
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =2520
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6840
                    Top =2520
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =22
                    BackColor =65535
                    Name ="Duff_A10"
                    ControlSource ="10Duff_A"
                    StatusBarText ="Duff depth at 10 meter point for transect A in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =6840
                    LayoutCachedTop =2520
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =2820
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6840
                    Top =2820
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =24
                    BackColor =65535
                    Name ="Duff_A12"
                    ControlSource ="12Duff_A"
                    StatusBarText ="Duff depth at 12 meter point for transect A in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =6840
                    LayoutCachedTop =2820
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =3120
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6840
                    Top =3120
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =26
                    BackColor =65535
                    Name ="Duff_A14"
                    ControlSource ="14Duff_A"
                    StatusBarText ="Duff depth at 14 meter point for transect A in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =6840
                    LayoutCachedTop =3120
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =3420
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8520
                    Top =1320
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =28
                    BackColor =65535
                    Name ="Duff_B2"
                    ControlSource ="2Duff_B"
                    StatusBarText ="Duff depth at 2 meter point for transect B in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =8520
                    LayoutCachedTop =1320
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =1620
                    BorderThemeColorIndex =0
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =95
                            Left =8520
                            Top =1080
                            Width =840
                            Height =240
                            BackColor =10079487
                            Name ="2Duff_B_Label"
                            Caption ="duff"
                            EventProcPrefix ="Ctl2Duff_B_Label"
                            LayoutCachedLeft =8520
                            LayoutCachedTop =1080
                            LayoutCachedWidth =9360
                            LayoutCachedHeight =1320
                            BorderThemeColorIndex =0
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8520
                    Top =1620
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =30
                    BackColor =65535
                    Name ="Duff_B4"
                    ControlSource ="4Duff_B"
                    StatusBarText ="Duff depth at 4 meter point for transect B in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =8520
                    LayoutCachedTop =1620
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =1920
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8520
                    Top =1920
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =32
                    BackColor =65535
                    Name ="Duff_B6"
                    ControlSource ="6Duff_B"
                    StatusBarText ="Duff depth at 6 meter point for transect B in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =8520
                    LayoutCachedTop =1920
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =2220
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8520
                    Top =2220
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =34
                    BackColor =65535
                    Name ="Duff_B8"
                    ControlSource ="8Duff_B"
                    StatusBarText ="Duff depth at 8 meter point for transect B in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =8520
                    LayoutCachedTop =2220
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =2520
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8520
                    Top =2520
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =36
                    BackColor =65535
                    Name ="Duff_B10"
                    ControlSource ="10Duff_B"
                    StatusBarText ="Duff depth at 10 meter point for transect B in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =8520
                    LayoutCachedTop =2520
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =2820
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8520
                    Top =2820
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =38
                    BackColor =65535
                    Name ="Duff_B12"
                    ControlSource ="12Duff_B"
                    StatusBarText ="Duff depth at 12 meter point for transect B in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =8520
                    LayoutCachedTop =2820
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =3120
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8520
                    Top =3120
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =40
                    BackColor =65535
                    Name ="Duff_B14"
                    ControlSource ="14Duff_B"
                    StatusBarText ="Duff depth at 14 meter point for transect B in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =8520
                    LayoutCachedTop =3120
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =3420
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10200
                    Top =1320
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =42
                    BackColor =65535
                    Name ="Duff_C2"
                    ControlSource ="2Duff_C"
                    StatusBarText ="Duff depth at 2 meter point for transect C in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =10200
                    LayoutCachedTop =1320
                    LayoutCachedWidth =11040
                    LayoutCachedHeight =1620
                    BorderThemeColorIndex =0
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =95
                            Left =10200
                            Top =1080
                            Width =840
                            Height =240
                            BackColor =10079487
                            Name ="2Duff_C_Label"
                            Caption ="duff"
                            EventProcPrefix ="Ctl2Duff_C_Label"
                            LayoutCachedLeft =10200
                            LayoutCachedTop =1080
                            LayoutCachedWidth =11040
                            LayoutCachedHeight =1320
                            BorderThemeColorIndex =0
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10200
                    Top =1620
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =44
                    BackColor =65535
                    Name ="Duff_C4"
                    ControlSource ="4Duff_C"
                    StatusBarText ="Duff depth at 4 meter point for transect C in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =10200
                    LayoutCachedTop =1620
                    LayoutCachedWidth =11040
                    LayoutCachedHeight =1920
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10200
                    Top =1920
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =46
                    BackColor =65535
                    Name ="Duff_C6"
                    ControlSource ="6Duff_C"
                    StatusBarText ="Duff depth at 6 meter point for transect C in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =10200
                    LayoutCachedTop =1920
                    LayoutCachedWidth =11040
                    LayoutCachedHeight =2220
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10200
                    Top =2220
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =48
                    BackColor =65535
                    Name ="Duff_C8"
                    ControlSource ="8Duff_C"
                    StatusBarText ="Duff depth at 8 meter point for transect C in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =10200
                    LayoutCachedTop =2220
                    LayoutCachedWidth =11040
                    LayoutCachedHeight =2520
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10200
                    Top =2520
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =50
                    BackColor =65535
                    Name ="Duff_C10"
                    ControlSource ="10Duff_C"
                    StatusBarText ="Duff depth at 10 meter point for transect C in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =10200
                    LayoutCachedTop =2520
                    LayoutCachedWidth =11040
                    LayoutCachedHeight =2820
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10200
                    Top =2820
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =52
                    BackColor =65535
                    Name ="Duff_C12"
                    ControlSource ="12Duff_C"
                    StatusBarText ="Duff depth at 12 meter point for transect C in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =10200
                    LayoutCachedTop =2820
                    LayoutCachedWidth =11040
                    LayoutCachedHeight =3120
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10200
                    Top =3120
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =54
                    BackColor =65535
                    Name ="Duff_C14"
                    ControlSource ="14Duff_C"
                    StatusBarText ="Duff depth at 14 meter point for transect C in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =10200
                    LayoutCachedTop =3120
                    LayoutCachedWidth =11040
                    LayoutCachedHeight =3420
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11880
                    Top =1320
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =56
                    BackColor =65535
                    Name ="Duff_D2"
                    ControlSource ="2Duff_D"
                    StatusBarText ="Duff depth at 2 meter point for transect D in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =11880
                    LayoutCachedTop =1320
                    LayoutCachedWidth =12720
                    LayoutCachedHeight =1620
                    BorderThemeColorIndex =0
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =95
                            Left =11880
                            Top =1080
                            Width =840
                            Height =240
                            BackColor =10079487
                            Name ="Duff_D_Label"
                            Caption ="duff"
                            LayoutCachedLeft =11880
                            LayoutCachedTop =1080
                            LayoutCachedWidth =12720
                            LayoutCachedHeight =1320
                            BorderThemeColorIndex =0
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11880
                    Top =1620
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =58
                    BackColor =65535
                    Name ="Duff_D4"
                    ControlSource ="4Duff_D"
                    StatusBarText ="Duff depth at 4 meter point for transect D in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =11880
                    LayoutCachedTop =1620
                    LayoutCachedWidth =12720
                    LayoutCachedHeight =1920
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11880
                    Top =1920
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =60
                    BackColor =65535
                    Name ="Duff_D6"
                    ControlSource ="6Duff_D"
                    StatusBarText ="Duff depth at 6 meter point for transect D in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =11880
                    LayoutCachedTop =1920
                    LayoutCachedWidth =12720
                    LayoutCachedHeight =2220
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11880
                    Top =2220
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =62
                    BackColor =65535
                    Name ="Duff_D8"
                    ControlSource ="8Duff_D"
                    StatusBarText ="Duff depth at 8 meter point for transect D in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =11880
                    LayoutCachedTop =2220
                    LayoutCachedWidth =12720
                    LayoutCachedHeight =2520
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11880
                    Top =2520
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =64
                    BackColor =65535
                    Name ="Duff_D10"
                    ControlSource ="10Duff_D"
                    StatusBarText ="Duff depth at 10 meter point for transect D in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =11880
                    LayoutCachedTop =2520
                    LayoutCachedWidth =12720
                    LayoutCachedHeight =2820
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11880
                    Top =2820
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =66
                    BackColor =65535
                    Name ="Duff_D12"
                    ControlSource ="12Duff_D"
                    StatusBarText ="Duff depth at 12 meter point for transect D in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =11880
                    LayoutCachedTop =2820
                    LayoutCachedWidth =12720
                    LayoutCachedHeight =3120
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11880
                    Top =3120
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    TabIndex =68
                    BackColor =65535
                    Name ="Duff_D14"
                    ControlSource ="14Duff_D"
                    StatusBarText ="Duff depth at 14 meter point for transect D in centimeters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =11880
                    LayoutCachedTop =3120
                    LayoutCachedWidth =12720
                    LayoutCachedHeight =3420
                    BorderThemeColorIndex =0
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
                    LayoutCachedLeft =6000
                    LayoutCachedTop =840
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =1080
                    BorderThemeColorIndex =0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =95
                    Left =7680
                    Top =840
                    Width =1680
                    Height =240
                    FontSize =10
                    BackColor =13434828
                    Name ="Label118"
                    Caption ="B"
                    LayoutCachedLeft =7680
                    LayoutCachedTop =840
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =1080
                    BorderThemeColorIndex =0
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
                    LayoutCachedLeft =9360
                    LayoutCachedTop =840
                    LayoutCachedWidth =11040
                    LayoutCachedHeight =1080
                    BorderThemeColorIndex =0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =95
                    Left =11040
                    Top =840
                    Width =1680
                    Height =240
                    FontSize =10
                    BackColor =13434828
                    Name ="Label120"
                    Caption ="D"
                    LayoutCachedLeft =11040
                    LayoutCachedTop =840
                    LayoutCachedWidth =12720
                    LayoutCachedHeight =1080
                    BorderThemeColorIndex =0
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
                    OverlapFlags =95
                    Left =5100
                    Top =600
                    Width =900
                    Height =720
                    Name ="Label122"
                    Caption ="point on transect (m)"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    Left =5100
                    Top =1320
                    Width =900
                    Height =300
                    Name ="Label123"
                    Caption ="2"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    Left =5100
                    Top =1620
                    Width =900
                    Height =300
                    Name ="Label124"
                    Caption ="4"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    Left =5100
                    Top =1920
                    Width =900
                    Height =300
                    Name ="Label125"
                    Caption ="6"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    Left =5100
                    Top =2220
                    Width =900
                    Height =300
                    Name ="Label126"
                    Caption ="8"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    Left =5100
                    Top =2520
                    Width =900
                    Height =300
                    Name ="Label127"
                    Caption ="10"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    Left =5100
                    Top =2820
                    Width =900
                    Height =300
                    Name ="Label128"
                    Caption ="12"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =87
                    Left =5100
                    Top =3120
                    Width =900
                    Height =300
                    Name ="Label130"
                    Caption ="14"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =0
                    Left =5100
                    Top =300
                    Width =2220
                    Height =240
                    FontSize =10
                    Name ="Label131"
                    Caption ="Litter and Duff Depth"
                    LayoutCachedLeft =5100
                    LayoutCachedTop =300
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =540
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =120
                    Top =120
                    Width =330
                    Height =300
                    Name ="Fuels_ID"
                    ControlSource ="Fuels_ID"
                    StatusBarText ="Unique record identifier - primary key"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =540
                    Top =120
                    Width =330
                    Height =300
                    TabIndex =69
                    Name ="Event_ID"
                    ControlSource ="Event_ID"
                    StatusBarText ="Foreign key to tbl_Events"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1485
                    Top =3660
                    Width =810
                    Height =300
                    TabIndex =75
                    Name ="Bearing_A"
                    StatusBarText ="Bearing of the plot slope + 180 in degrees"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =1485
                    LayoutCachedTop =3660
                    LayoutCachedWidth =2295
                    LayoutCachedHeight =3960
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
                    SpecialEffect =0
                    OverlapFlags =255
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2280
                    Top =3660
                    Width =810
                    Height =300
                    TabIndex =76
                    BackColor =13434828
                    Name ="Bearing_B"
                    StatusBarText ="Bearing of transect 1 in degrees"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =2280
                    LayoutCachedTop =3660
                    LayoutCachedWidth =3090
                    LayoutCachedHeight =3960
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =223
                            Left =2280
                            Top =3420
                            Width =810
                            Height =240
                            FontWeight =400
                            BackColor =13434828
                            Name ="Bearing_B_Label"
                            Caption ="B"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3120
                    Top =3660
                    Width =810
                    Height =300
                    TabIndex =77
                    Name ="Bearing_C"
                    StatusBarText ="Bearing of transect 3 + 180"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =3120
                    LayoutCachedTop =3660
                    LayoutCachedWidth =3930
                    LayoutCachedHeight =3960
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =3120
                            Top =3420
                            Width =810
                            Height =240
                            FontWeight =400
                            Name ="Bearing_C_Label"
                            Caption ="C"
                            LayoutCachedLeft =3120
                            LayoutCachedTop =3420
                            LayoutCachedWidth =3930
                            LayoutCachedHeight =3660
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3945
                    Top =3660
                    Width =810
                    Height =300
                    TabIndex =78
                    BackColor =13434828
                    Name ="Bearing_D"
                    StatusBarText ="Bearing of the plot slope"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =3945
                    LayoutCachedTop =3660
                    LayoutCachedWidth =4755
                    LayoutCachedHeight =3960
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =3945
                            Top =3420
                            Width =810
                            Height =240
                            FontWeight =400
                            BackColor =13434828
                            Name ="Bearing_D_Label"
                            Caption ="D"
                            LayoutCachedLeft =3945
                            LayoutCachedTop =3420
                            LayoutCachedWidth =4755
                            LayoutCachedHeight =3660
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =127
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1485
                    Top =3960
                    Width =810
                    Height =300
                    TabIndex =80
                    Name ="Slope_A"
                    StatusBarText ="Slope of transect A to nearest half percent"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =1485
                    LayoutCachedTop =3960
                    LayoutCachedWidth =2295
                    LayoutCachedHeight =4260
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
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
                    SpecialEffect =0
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2280
                    Top =3960
                    Width =810
                    Height =300
                    TabIndex =81
                    BackColor =13434828
                    Name ="Slope_B"
                    StatusBarText ="Slope of transect B to nearest half percent"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =2280
                    LayoutCachedTop =3960
                    LayoutCachedWidth =3090
                    LayoutCachedHeight =4260
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =87
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
                    SpecialEffect =0
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3120
                    Top =3960
                    Width =810
                    Height =300
                    TabIndex =82
                    Name ="Slope_C"
                    StatusBarText ="Slope of transect C to nearest half percent"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =3120
                    LayoutCachedTop =3960
                    LayoutCachedWidth =3930
                    LayoutCachedHeight =4260
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3945
                    Top =3960
                    Width =810
                    Height =300
                    TabIndex =83
                    BackColor =13434828
                    Name ="Slope_D"
                    StatusBarText ="Slope of transect C to nearest half percent"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"

                    LayoutCachedLeft =3945
                    LayoutCachedTop =3960
                    LayoutCachedWidth =4755
                    LayoutCachedHeight =4260
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1440
                    Top =1020
                    Width =1080
                    Height =300
                    TabIndex =1
                    BackColor =62207
                    Name ="1HR_A"
                    ControlSource ="1HR_A"
                    StatusBarText ="One hour fuel intercept for transect A"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl1HR_A"
                    ConditionalFormat = Begin
                        0x0100000074000000020000000000000001000000000000000200000001000000 ,
                        0x00000000fff20000000000000400000006000000090000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x3000000031003000300000002d00310000000000
                    End

                    LayoutCachedLeft =1440
                    LayoutCachedTop =1020
                    LayoutCachedWidth =2520
                    LayoutCachedHeight =1320
                    ConditionalFormat14 = Begin
                        0x01000200000000000000010000000100000000000000fff20000010000003000 ,
                        0x0300000031003000300000000000000000000000000000000000000000000004 ,
                        0x0000000100000000000000ffffff00020000002d003100000000000000000000 ,
                        0x000000000000000000000000
                    End
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
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1440
                    Top =1320
                    Width =1080
                    Height =300
                    TabIndex =4
                    BackColor =62207
                    Name ="1HR_B"
                    ControlSource ="1HR_B"
                    StatusBarText ="One hour fuel intercept for transect B"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl1HR_B"
                    ConditionalFormat = Begin
                        0x0100000074000000020000000000000001000000000000000200000001000000 ,
                        0x00000000fff20000000000000400000006000000090000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x3000000031003000300000002d00310000000000
                    End

                    LayoutCachedLeft =1440
                    LayoutCachedTop =1320
                    LayoutCachedWidth =2520
                    LayoutCachedHeight =1620
                    ConditionalFormat14 = Begin
                        0x01000200000000000000010000000100000000000000fff20000010000003000 ,
                        0x0300000031003000300000000000000000000000000000000000000000000004 ,
                        0x0000000100000000000000ffffff00020000002d003100000000000000000000 ,
                        0x000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
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
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1440
                    Top =1620
                    Width =1080
                    Height =300
                    TabIndex =7
                    BackColor =62207
                    Name ="1HR_C"
                    ControlSource ="1HR_C"
                    StatusBarText ="One hour fuel intercept for transect C"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl1HR_C"
                    ConditionalFormat = Begin
                        0x0100000074000000020000000000000001000000000000000200000001000000 ,
                        0x00000000fff20000000000000400000006000000090000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x3000000031003000300000002d00310000000000
                    End

                    LayoutCachedLeft =1440
                    LayoutCachedTop =1620
                    LayoutCachedWidth =2520
                    LayoutCachedHeight =1920
                    ConditionalFormat14 = Begin
                        0x01000200000000000000010000000100000000000000fff20000010000003000 ,
                        0x0300000031003000300000000000000000000000000000000000000000000004 ,
                        0x0000000100000000000000ffffff00020000002d003100000000000000000000 ,
                        0x000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
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
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1440
                    Top =1920
                    Width =1080
                    Height =300
                    TabIndex =10
                    BackColor =62207
                    Name ="DI_1HR"
                    ControlSource ="1HR_D"
                    StatusBarText ="One hour fuel intercept for transect D"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000074000000020000000000000001000000000000000200000001000000 ,
                        0x00000000fff20000000000000400000006000000090000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x3000000031003000300000002d00310000000000
                    End

                    LayoutCachedLeft =1440
                    LayoutCachedTop =1920
                    LayoutCachedWidth =2520
                    LayoutCachedHeight =2220
                    ConditionalFormat14 = Begin
                        0x01000200000000000000010000000100000000000000fff20000010000003000 ,
                        0x0300000031003000300000000000000000000000000000000000000000000004 ,
                        0x0000000100000000000000ffffff00020000002d003100000000000000000000 ,
                        0x000000000000000000000000
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2520
                    Top =1020
                    Width =1080
                    Height =300
                    TabIndex =2
                    BackColor =62207
                    Name ="10HR_A"
                    ControlSource ="10HR_A"
                    StatusBarText ="Ten hour fuel intercept for transect A"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl10HR_A"
                    ConditionalFormat = Begin
                        0x0100000074000000020000000000000001000000000000000200000001000000 ,
                        0x00000000fff20000000000000400000006000000090000000100000000000000 ,
                        0xccff990000000000000000000000000000000000000000000000000000000000 ,
                        0x3000000031003000300000002d00310000000000
                    End

                    LayoutCachedLeft =2520
                    LayoutCachedTop =1020
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =1320
                    ConditionalFormat14 = Begin
                        0x01000200000000000000010000000100000000000000fff20000010000003000 ,
                        0x0300000031003000300000000000000000000000000000000000000000000004 ,
                        0x0000000100000000000000ccff9900020000002d003100000000000000000000 ,
                        0x000000000000000000000000
                    End
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =95
                            Left =2520
                            Top =600
                            Width =1080
                            Height =420
                            BackColor =13434828
                            Name ="10HR_A_Label"
                            Caption ="     10-hr     (0.25-1 in)"
                            EventProcPrefix ="Ctl10HR_A_Label"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2520
                    Top =1320
                    Width =1080
                    Height =300
                    TabIndex =5
                    BackColor =62207
                    Name ="10HR_B"
                    ControlSource ="10HR_B"
                    StatusBarText ="Ten hour fuel intercept for transect B"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl10HR_B"
                    ConditionalFormat = Begin
                        0x0100000074000000020000000000000001000000000000000200000001000000 ,
                        0x00000000fff20000000000000400000006000000090000000100000000000000 ,
                        0xccff990000000000000000000000000000000000000000000000000000000000 ,
                        0x3000000031003000300000002d00310000000000
                    End

                    LayoutCachedLeft =2520
                    LayoutCachedTop =1320
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =1620
                    ConditionalFormat14 = Begin
                        0x01000200000000000000010000000100000000000000fff20000010000003000 ,
                        0x0300000031003000300000000000000000000000000000000000000000000004 ,
                        0x0000000100000000000000ccff9900020000002d003100000000000000000000 ,
                        0x000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
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
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2520
                    Top =1620
                    Width =1080
                    Height =300
                    TabIndex =8
                    BackColor =62207
                    Name ="10HR_C"
                    ControlSource ="10HR_C"
                    StatusBarText ="Ten hour fuel intercept for transect C"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl10HR_C"
                    ConditionalFormat = Begin
                        0x0100000074000000020000000000000001000000000000000200000001000000 ,
                        0x00000000fff20000000000000400000006000000090000000100000000000000 ,
                        0xccff990000000000000000000000000000000000000000000000000000000000 ,
                        0x3000000031003000300000002d00310000000000
                    End

                    LayoutCachedLeft =2520
                    LayoutCachedTop =1620
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =1920
                    ConditionalFormat14 = Begin
                        0x01000200000000000000010000000100000000000000fff20000010000003000 ,
                        0x0300000031003000300000000000000000000000000000000000000000000004 ,
                        0x0000000100000000000000ccff9900020000002d003100000000000000000000 ,
                        0x000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =87
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
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2520
                    Top =1920
                    Width =1080
                    Height =300
                    TabIndex =11
                    BackColor =62207
                    Name ="DI_10HR"
                    ControlSource ="10HR_D"
                    StatusBarText ="Ten hour fuel intercept for transect D"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000074000000020000000000000001000000000000000200000001000000 ,
                        0x00000000fff20000000000000400000006000000090000000100000000000000 ,
                        0xccff990000000000000000000000000000000000000000000000000000000000 ,
                        0x3000000031003000300000002d00310000000000
                    End

                    LayoutCachedLeft =2520
                    LayoutCachedTop =1920
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =2220
                    ConditionalFormat14 = Begin
                        0x01000200000000000000010000000100000000000000fff20000010000003000 ,
                        0x0300000031003000300000000000000000000000000000000000000000000004 ,
                        0x0000000100000000000000ccff9900020000002d003100000000000000000000 ,
                        0x000000000000000000000000
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3600
                    Top =1020
                    Width =1080
                    Height =300
                    TabIndex =3
                    BackColor =62207
                    Name ="100HR_A"
                    ControlSource ="100HR_A"
                    StatusBarText ="Hundred hour fuel intercept for transect A"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl100HR_A"
                    ConditionalFormat = Begin
                        0x0100000074000000020000000000000001000000000000000200000001000000 ,
                        0x00000000fff20000000000000400000006000000090000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x3000000031003000300000002d00310000000000
                    End

                    LayoutCachedLeft =3600
                    LayoutCachedTop =1020
                    LayoutCachedWidth =4680
                    LayoutCachedHeight =1320
                    ConditionalFormat14 = Begin
                        0x01000200000000000000010000000100000000000000fff20000010000003000 ,
                        0x0300000031003000300000000000000000000000000000000000000000000004 ,
                        0x0000000100000000000000ffffff00020000002d003100000000000000000000 ,
                        0x000000000000000000000000
                    End
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
                            LayoutCachedLeft =3600
                            LayoutCachedTop =600
                            LayoutCachedWidth =4680
                            LayoutCachedHeight =1020
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3600
                    Top =1320
                    Width =1080
                    Height =300
                    TabIndex =6
                    BackColor =62207
                    Name ="100HR_B"
                    ControlSource ="100HR_B"
                    StatusBarText ="Hundred hour fuel intercept for transect B"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl100HR_B"
                    ConditionalFormat = Begin
                        0x0100000074000000020000000000000001000000000000000200000001000000 ,
                        0x00000000fff20000000000000400000006000000090000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x3000000031003000300000002d00310000000000
                    End

                    LayoutCachedLeft =3600
                    LayoutCachedTop =1320
                    LayoutCachedWidth =4680
                    LayoutCachedHeight =1620
                    ConditionalFormat14 = Begin
                        0x01000200000000000000010000000100000000000000fff20000010000003000 ,
                        0x0300000031003000300000000000000000000000000000000000000000000004 ,
                        0x0000000100000000000000ffffff00020000002d003100000000000000000000 ,
                        0x000000000000000000000000
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3600
                    Top =1620
                    Width =1080
                    Height =300
                    TabIndex =9
                    BackColor =62207
                    Name ="100HR_C"
                    ControlSource ="100HR_C"
                    StatusBarText ="Hundred hour fuel intercept for transect C"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    EventProcPrefix ="Ctl100HR_C"
                    ConditionalFormat = Begin
                        0x0100000074000000020000000000000001000000000000000200000001000000 ,
                        0x00000000fff20000000000000400000006000000090000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x3000000031003000300000002d00310000000000
                    End

                    LayoutCachedLeft =3600
                    LayoutCachedTop =1620
                    LayoutCachedWidth =4680
                    LayoutCachedHeight =1920
                    ConditionalFormat14 = Begin
                        0x01000200000000000000010000000100000000000000fff20000010000003000 ,
                        0x0300000031003000300000000000000000000000000000000000000000000004 ,
                        0x0000000100000000000000ffffff00020000002d003100000000000000000000 ,
                        0x000000000000000000000000
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3600
                    Top =1920
                    Width =1080
                    Height =300
                    TabIndex =12
                    BackColor =62207
                    Name ="DI_100HR"
                    ControlSource ="100HR_D"
                    StatusBarText ="Hundred hour fuel intercept for transect D"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000074000000020000000000000001000000000000000200000001000000 ,
                        0x00000000fff20000000000000400000006000000090000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x3000000031003000300000002d00310000000000
                    End

                    LayoutCachedLeft =3600
                    LayoutCachedTop =1920
                    LayoutCachedWidth =4680
                    LayoutCachedHeight =2220
                    ConditionalFormat14 = Begin
                        0x01000200000000000000010000000100000000000000fff20000010000003000 ,
                        0x0300000031003000300000000000000000000000000000000000000000000004 ,
                        0x0000000100000000000000ffffff00020000002d003100000000000000000000 ,
                        0x000000000000000000000000
                    End
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
                    OverlapFlags =87
                    Left =540
                    Top =1320
                    Width =900
                    Height =300
                    Name ="Label133"
                    Caption ="B"
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =900
                    Top =2280
                    Width =606
                    Height =288
                    TabIndex =70
                    Name ="ButtonA1"
                    Caption ="+ 1"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    OnKeyDown ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =1620
                    Top =2280
                    Width =606
                    Height =288
                    TabIndex =71
                    Name ="ButtonA5"
                    Caption ="+ 5"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    OnKeyDown ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =2340
                    Top =2280
                    Width =606
                    Height =288
                    TabIndex =72
                    Name ="ButtonS1"
                    Caption ="- 1"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    OnKeyDown ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =3060
                    Top =2280
                    Width =606
                    Height =288
                    TabIndex =73
                    Name ="ButtonS5"
                    Caption ="- 5"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    OnKeyDown ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =3780
                    Top =2280
                    Width =606
                    Height =288
                    TabIndex =74
                    Name ="ButtonZero"
                    Caption ="0"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    OnKeyDown ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =4920
                    Top =3840
                    Width =1185
                    Height =300
                    TabIndex =79
                    Name ="ButtonTransect"
                    Caption ="Edit Transect"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
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
'                                        subroutines for handling litter & duff value highlighting
'                                        backcolor = #FFFF00, yellow), fixed litter & duff control naming
'                                        which was 2Litter_D, Duff_D2, etc. to consistent Litter_ or Duff_
'                                        followed by transect letter and point # (Litter_B14, etc.) &
'                                        adjusted all associated control references in this module
'               BLC - 3/24/2016 - 1.02 - added documentation @ control tab order
' =================================

' ---------------------------------
'  CONTROL TAB ORDER
'
'  Tab order MUST be the following (per HT 3/24/2016):
'
'   1 - 12)# of intercepts: 1HR-10HR-100HR  begin tabs L to R (across then down),  begin w/ A the down to D
'   13 - 68) litter & duff depth: start @ transect A, do litter, then duff @ pt 2 for that transect down to 14
'                        then go to B, C, D
'                        (complete all of A, then B, C, D but within do litter then duff for @ pt)
'
'   no other controls within this form are in the tab order
' ---------------------------------

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
'       BLC, 3/16/2016 - added error handling & documentation, handled litter/duff highlighting
'       BLC, 3/23/2017 - revised naming convention for transect D 1, 10, & 100-hr fuel intercepts
'                        from DI_1HR > 1HR_D, etc.
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
    
    'handle litter/duff highlighting from saved data
    Dim ctrl As Control
    
    For Each ctrl In Me.Controls
        
        'handle only visible, enabled textboxes
        If ctrl.ControlType = acTextBox Then
        
            If ctrl.Visible = True And ctrl.Enabled = True Then
        
                ctrl.SetFocus  'Required to avoid Error #2185 control must have focus to reference property or method
            
                'isolate only Litter_ and Duff_ textboxes
                If Len(ctrl.name) > Len(Replace(Replace(ctrl.name, "Litter_", ""), "Duff_", "")) Then
                    SetLitterDuffHighlight ctrl
                End If
                
            End If
            
        End If
        
    Next

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
' NOTE: There are 56 afterupdate events for litter & duff values.
'
'       To avoid requiring 56 separate changes, the SetLitterDuffHighlight()
'       subroutine is provided for handling all 56 actions provided ALL litter/duff
'       controls should behave similarly (including having the same value requirements).
'
'       If controls have different requirements, they will need to be handled
'       and changed separately. This means if the action is changed for one,
'       ALL 56 should be changed appropriately.
'
'       These 56 events mirror the 56 fields in the denormalized fuels table
'       previously setup (pre-2014 by RDB/HT) within the backend & frontend
'       application to mirror the uplands field datasheet.
'       [Normalizing would reduce code proliferation & performance!]
'-------------------------

' ---------------------------------
' SUB:          SetLitterDuffHighlight
' Description:  handles litter/duff highlight actions
' Parameters:   ctrl - litter/duff textbox control (textbox)
' Returns:      -
' Assumptions:  highlighting will be consistent across all textboxes
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub SetLitterDuffHighlight(ctrl As TextBox)
On Error GoTo Err_Handler

    'set the backcolor to white when the value reaches a threshold >= 0, checking for NULL and empty values
    SetControlBackcolor ctrl, RGB(255, 255, 255), True, True, 0, "gteq"
   
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetLitterDuffHighlight[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

'==================================
'              POINT 2
'==================================

' ---------------------------------
' SUB:          Litter_A2_AfterUpdate
' Description:  handles litter actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Litter_A2_AfterUpdate()
On Error GoTo Err_Handler

    SetLitterDuffHighlight Litter_A2
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Litter_A2_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Litter_B2_AfterUpdate
' Description:  handles litter actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Litter_B2_AfterUpdate()
On Error GoTo Err_Handler

    SetLitterDuffHighlight Litter_B2
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Litter_B2_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Litter_C2_AfterUpdate
' Description:  handles litter actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Litter_C2_AfterUpdate()
On Error GoTo Err_Handler

    SetLitterDuffHighlight Litter_C2
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Litter_C2_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Litter_D2_AfterUpdate
' Description:  handles litter actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Litter_D2_AfterUpdate()
On Error GoTo Err_Handler

    SetLitterDuffHighlight Litter_D2
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Litter_D2_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Duff_A2_AfterUpdate
' Description:  handles duff actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Duff_A2_AfterUpdate()
On Error GoTo Err_Handler

    'clear highlight if not null
    SetLitterDuffHighlight Duff_A2

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Duff_A2_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Duff_B2_AfterUpdate
' Description:  handles duff actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Duff_B2_AfterUpdate()
On Error GoTo Err_Handler

    'clear highlight if not null
    SetLitterDuffHighlight Duff_B2

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Duff_B2_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Duff_C2_AfterUpdate
' Description:  handles duff actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Duff_C2_AfterUpdate()
On Error GoTo Err_Handler

    'clear highlight if not null
    SetLitterDuffHighlight Duff_C2

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Duff_C2_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Duff_D2_AfterUpdate
' Description:  handles duff actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Duff_D2_AfterUpdate()
On Error GoTo Err_Handler

    'clear highlight if not null
    SetLitterDuffHighlight Duff_D2

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Duff_D2_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

'==================================
'              POINT 4
'==================================

' ---------------------------------
' SUB:          Litter_A4_AfterUpdate
' Description:  handles litter actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Litter_A4_AfterUpdate()
On Error GoTo Err_Handler

    SetLitterDuffHighlight Litter_A4
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Litter_A4_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Litter_B4_AfterUpdate
' Description:  handles litter actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Litter_B4_AfterUpdate()
On Error GoTo Err_Handler

    SetLitterDuffHighlight Litter_B4
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Litter_B4_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Litter_C4_AfterUpdate
' Description:  handles litter actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Litter_C4_AfterUpdate()
On Error GoTo Err_Handler

    SetLitterDuffHighlight Litter_C4
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Litter_C4_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Litter_D4_AfterUpdate
' Description:  handles litter actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Litter_D4_AfterUpdate()
On Error GoTo Err_Handler

    SetLitterDuffHighlight Litter_D4
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Litter_D4_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Duff_A4_AfterUpdate
' Description:  handles duff actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Duff_A4_AfterUpdate()
On Error GoTo Err_Handler

    'clear highlight if not null
    SetLitterDuffHighlight Duff_A4

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Duff_A4_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Duff_B4_AfterUpdate
' Description:  handles duff actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Duff_B4_AfterUpdate()
On Error GoTo Err_Handler

    'clear highlight if not null
    SetLitterDuffHighlight Duff_B4

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Duff_B4_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Duff_C4_AfterUpdate
' Description:  handles duff actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Duff_C4_AfterUpdate()
On Error GoTo Err_Handler

    'clear highlight if not null
    SetLitterDuffHighlight Duff_C4

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Duff_C4_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Duff_D4_AfterUpdate
' Description:  handles duff actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Duff_D4_AfterUpdate()
On Error GoTo Err_Handler

    'clear highlight if not null
    SetLitterDuffHighlight Duff_D4

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Duff_D4_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

'==================================
'              POINT 6
'==================================

' ---------------------------------
' SUB:          Litter_A6_AfterUpdate
' Description:  handles litter actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Litter_A6_AfterUpdate()
On Error GoTo Err_Handler

    SetLitterDuffHighlight Litter_A6
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Litter_A6_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Litter_B6_AfterUpdate
' Description:  handles litter actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Litter_B6_AfterUpdate()
On Error GoTo Err_Handler

    SetLitterDuffHighlight Litter_B6
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Litter_B6_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Litter_C6_AfterUpdate
' Description:  handles litter actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Litter_C6_AfterUpdate()
On Error GoTo Err_Handler

    SetLitterDuffHighlight Litter_C6
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Litter_C6_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Litter_D6_AfterUpdate
' Description:  handles litter actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Litter_D6_AfterUpdate()
On Error GoTo Err_Handler

    SetLitterDuffHighlight Litter_D6
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Litter_D6_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Duff_A6_AfterUpdate
' Description:  handles duff actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Duff_A6_AfterUpdate()
On Error GoTo Err_Handler

    'clear highlight if not null
    SetLitterDuffHighlight Duff_A6

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Duff_A6_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Duff_B6_AfterUpdate
' Description:  handles duff actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Duff_B6_AfterUpdate()
On Error GoTo Err_Handler

    'clear highlight if not null
    SetLitterDuffHighlight Duff_B6

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Duff_B6_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Duff_C6_AfterUpdate
' Description:  handles duff actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Duff_C6_AfterUpdate()
On Error GoTo Err_Handler

    'clear highlight if not null
    SetLitterDuffHighlight Duff_C6

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Duff_C6_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Duff_D6_AfterUpdate
' Description:  handles duff actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Duff_D6_AfterUpdate()
On Error GoTo Err_Handler

    'clear highlight if not null
    SetLitterDuffHighlight Duff_D6

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Duff_D6_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

'==================================
'              POINT 8
'==================================

' ---------------------------------
' SUB:          Litter_A8_AfterUpdate
' Description:  handles litter actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Litter_A8_AfterUpdate()
On Error GoTo Err_Handler

    SetLitterDuffHighlight Litter_A8
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Litter_A8_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Litter_B8_AfterUpdate
' Description:  handles litter actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Litter_B8_AfterUpdate()
On Error GoTo Err_Handler

    SetLitterDuffHighlight Litter_B8
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Litter_B8_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Litter_C8_AfterUpdate
' Description:  handles litter actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Litter_C8_AfterUpdate()
On Error GoTo Err_Handler

    SetLitterDuffHighlight Litter_C8
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Litter_C8_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Litter_D8_AfterUpdate
' Description:  handles litter actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Litter_D8_AfterUpdate()
On Error GoTo Err_Handler

    SetLitterDuffHighlight Litter_D8
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Litter_D8_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Duff_A8_AfterUpdate
' Description:  handles duff actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Duff_A8_AfterUpdate()
On Error GoTo Err_Handler

    'clear highlight if not null
    SetLitterDuffHighlight Duff_A8

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Duff_A8_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Duff_B8_AfterUpdate
' Description:  handles duff actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Duff_B8_AfterUpdate()
On Error GoTo Err_Handler

    'clear highlight if not null
    SetLitterDuffHighlight Duff_B8

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Duff_B8_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Duff_C8_AfterUpdate
' Description:  handles duff actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Duff_C8_AfterUpdate()
On Error GoTo Err_Handler

    'clear highlight if not null
    SetLitterDuffHighlight Duff_C8

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Duff_C8_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Duff_D8_AfterUpdate
' Description:  handles duff actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Duff_D8_AfterUpdate()
On Error GoTo Err_Handler

    'clear highlight if not null
    SetLitterDuffHighlight Duff_D8

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Duff_D8_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

'==================================
'              POINT 10
'==================================

' ---------------------------------
' SUB:          Litter_A10_AfterUpdate
' Description:  handles litter actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Litter_A10_AfterUpdate()
On Error GoTo Err_Handler

    SetLitterDuffHighlight Litter_A10
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Litter_A10_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Litter_B10_AfterUpdate
' Description:  handles litter actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Litter_B10_AfterUpdate()
On Error GoTo Err_Handler

    SetLitterDuffHighlight Litter_B10
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Litter_B10_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Litter_C10_AfterUpdate
' Description:  handles litter actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Litter_C10_AfterUpdate()
On Error GoTo Err_Handler

    SetLitterDuffHighlight Litter_C10
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Litter_C10_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Litter_D10_AfterUpdate
' Description:  handles litter actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Litter_D10_AfterUpdate()
On Error GoTo Err_Handler

    SetLitterDuffHighlight Litter_D10
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Litter_D10_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Duff_A10_AfterUpdate
' Description:  handles duff actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Duff_A10_AfterUpdate()
On Error GoTo Err_Handler

    'clear highlight if not null
    SetLitterDuffHighlight Duff_A10

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Duff_A10_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Duff_B10_AfterUpdate
' Description:  handles duff actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Duff_B10_AfterUpdate()
On Error GoTo Err_Handler

    'clear highlight if not null
    SetLitterDuffHighlight Duff_B10

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Duff_B10_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Duff_C10_AfterUpdate
' Description:  handles duff actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Duff_C10_AfterUpdate()
On Error GoTo Err_Handler

    'clear highlight if not null
    SetLitterDuffHighlight Duff_C10

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Duff_C10_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Duff_D10_AfterUpdate
' Description:  handles duff actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Duff_D10_AfterUpdate()
On Error GoTo Err_Handler

    'clear highlight if not null
    SetLitterDuffHighlight Duff_D10

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Duff_D10_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

'==================================
'              POINT 12
'==================================

' ---------------------------------
' SUB:          Litter_A12_AfterUpdate
' Description:  handles litter actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Litter_A12_AfterUpdate()
On Error GoTo Err_Handler

    SetLitterDuffHighlight Litter_A12
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Litter_A12_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Litter_B12_AfterUpdate
' Description:  handles litter actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Litter_B12_AfterUpdate()
On Error GoTo Err_Handler

    SetLitterDuffHighlight Litter_B12
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Litter_B12_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Litter_C12_AfterUpdate
' Description:  handles litter actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Litter_C12_AfterUpdate()
On Error GoTo Err_Handler

    SetLitterDuffHighlight Litter_C12
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Litter_C12_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Litter_D12_AfterUpdate
' Description:  handles litter actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Litter_D12_AfterUpdate()
On Error GoTo Err_Handler

    SetLitterDuffHighlight Litter_D12
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Litter_D12_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Duff_A12_AfterUpdate
' Description:  handles duff actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Duff_A12_AfterUpdate()
On Error GoTo Err_Handler

    'clear highlight if not null
    SetLitterDuffHighlight Duff_A12

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Duff_A12_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Duff_B12_AfterUpdate
' Description:  handles duff actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Duff_B12_AfterUpdate()
On Error GoTo Err_Handler

    'clear highlight if not null
    SetLitterDuffHighlight Duff_B12

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Duff_B12_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Duff_C12_AfterUpdate
' Description:  handles duff actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Duff_C12_AfterUpdate()
On Error GoTo Err_Handler

    'clear highlight if not null
    SetLitterDuffHighlight Duff_C12

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Duff_C12_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Duff_D12_AfterUpdate
' Description:  handles duff actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Duff_D12_AfterUpdate()
On Error GoTo Err_Handler

    'clear highlight if not null
    SetLitterDuffHighlight Duff_D12

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Duff_D12_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

'==================================
'              POINT 14
'==================================

' ---------------------------------
' SUB:          Litter_A14_AfterUpdate
' Description:  handles litter actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Litter_A14_AfterUpdate()
On Error GoTo Err_Handler

    SetLitterDuffHighlight Litter_A14
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Litter_A14_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Litter_B14_AfterUpdate
' Description:  handles litter actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Litter_B14_AfterUpdate()
On Error GoTo Err_Handler

    SetLitterDuffHighlight Litter_B14
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Litter_B14_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Litter_C14_AfterUpdate
' Description:  handles litter actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Litter_C14_AfterUpdate()
On Error GoTo Err_Handler

    SetLitterDuffHighlight Litter_C14
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Litter_C14_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Litter_D14_AfterUpdate
' Description:  handles litter actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Litter_D14_AfterUpdate()
On Error GoTo Err_Handler

    SetLitterDuffHighlight Litter_D14
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Litter_D14_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Duff_A14_AfterUpdate
' Description:  handles duff actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Duff_A14_AfterUpdate()
On Error GoTo Err_Handler

    'clear highlight if not null
    SetLitterDuffHighlight Duff_A14

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Duff_A14_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Duff_B14_AfterUpdate
' Description:  handles duff actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Duff_B14_AfterUpdate()
On Error GoTo Err_Handler

    'clear highlight if not null
    SetLitterDuffHighlight Duff_B14

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Duff_B14_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Duff_C14_AfterUpdate
' Description:  handles duff actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Duff_C14_AfterUpdate()
On Error GoTo Err_Handler

    'clear highlight if not null
    SetLitterDuffHighlight Duff_C14

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Duff_C14_AfterUpdate[Form_fsub_Fuels_LD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Duff_D14_AfterUpdate
' Description:  handles duff actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/16/2016 - initial version
' ---------------------------------
Private Sub Duff_D14_AfterUpdate()
On Error GoTo Err_Handler

    'clear highlight if not null
    SetLitterDuffHighlight Duff_D14

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Duff_D14_AfterUpdate[Form_fsub_Fuels_LD])"
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

Private Sub Duff_A10_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Duff_B10_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Duff_C10_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub Litter_A10_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Litter_B10_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Litter_C10_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Duff_A12_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Duff_B12_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Duff_C12_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Litter_A12_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Litter_B12_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Litter_C12_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Duff_A14_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Duff_B14_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Duff_C14_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Litter_A14_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Litter_B14_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Litter_C14_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub Duff_A2_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Duff_B2_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Duff_C2_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Litter_A2_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Litter_B2_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Litter_C2_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Duff_A4_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Duff_B4_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Duff_C4_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Litter_A4_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Litter_B4_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Litter_C4_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Duff_A6_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Duff_B6_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Duff_C6_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Litter_A6_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Litter_B6_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Litter_C6_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Duff_A8_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Duff_B8_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Duff_C8_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Litter_A8_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Litter_B8_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Litter_C8_KeyDown(KeyCode As Integer, Shift As Integer)
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
