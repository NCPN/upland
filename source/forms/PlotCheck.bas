Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DataEntry = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7560
    DatasheetFontHeight =11
    ItemSuffix =53
    Left =-1155
    Top =120
    Right =6660
    Bottom =14160
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0xce54d73357e8e440
    End
    Caption ="_List"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    OrderByOnLoad =0
    FilterOnLoad =255
    OrderByOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderColor =16777215
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ComboBox
            AddColon = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            BackColor =4144959
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =93
                    Width =3480
                    Height =300
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblTitle"
                    Caption ="title"
                    GridlineColor =10921638
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =300
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =223
                    Left =120
                    Top =120
                    Width =7260
                    Height =540
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblDirections"
                    Caption ="directions"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =120
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =660
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =1200
                    Top =1080
                    Width =3840
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblTemplate"
                    Caption ="Check"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =1200
                    LayoutCachedTop =1080
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =1395
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =223
                    Left =1440
                    Top =600
                    Width =720
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblPlot"
                    Caption ="Plot #"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =1440
                    LayoutCachedTop =600
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =915
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =6780
                    Top =180
                    Width =720
                    ForeColor =4210752
                    Name ="btnAddTemplate"
                    Caption ="Add Record"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add new template"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b09880ff201010ff201010ff201010ff201010ff201010ff ,
                        0x201010ff201010ff201010ff201010ff201010ff201010ff201010ff00000000 ,
                        0x0000000000000000c0a090fffff8f0fffff8f0fffff0f0fffff0e0fff0e8e0ff ,
                        0xf0e8d0fff0e0d0fff0e0d0fff0e0d0fff0d8d0fff0d8d0ff201810ff00000000 ,
                        0x0000000000000000c0a090ffffffffffd07850ffd07840ffd07040ffc07040ff ,
                        0xc06840ffc06840ffc06840ffc07040ffa06040fff0e0d0ff403830ff00000000 ,
                        0x0000000000000000c0a890ffffffffffd07850fff0b8a0fff0b090fff0a880ff ,
                        0xf0a080fff09870fff09870fff0a880ffc09880fffff0f0ff909090ff00000000 ,
                        0x0000000000000000c0a890ffffffffffd07850ffd07850ffd07840ffd07040ff ,
                        0xc07040ffc07050ffd09070ff70b8c0ff90d8f0ff90f0ffff40c0e0ffa0f0ffff ,
                        0xa0e8ffff90d8f0ffc0a8a0fffffffffffffffffffffffffffffffffffff8f0ff ,
                        0xfff8f0fffff8f0fffff8f0ffb0e8ffff30b8e0ff80e8ffff60c8e0ff90f0ffff ,
                        0x30b8e0ffa0e8ffffc0a8a0ffc0a8a0ffc0a890ffc0a090ffc0a090ffc0a090ff ,
                        0xc09880ffc0a090ffd0c0b0ffa0e8ffff90f0ffffc0f8ffffb0e8f0ffc0f8ffff ,
                        0x90f0ffffa0f0ffff000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000020a8e0ff50c0e0ffb0e8f0fff0ffffffb0e8f0ff ,
                        0x50c0e0ff30b8e0ff000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000080e8ffc090f0ffffc0f8ffffb0e8f0ffc0f8ffff ,
                        0x90f0ffff90d8e0ff000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000050d8ff8030b8e0ff90f0ffff60c0e0ff90f0ffff ,
                        0x30b8e0ff50d0f080000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000030b0e0a040c8f09080e8ffc020b0e0ff70e8ffc0 ,
                        0x50d8f08030b0e080000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =6780
                    LayoutCachedTop =180
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =540
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =15060409
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =6120
                    Top =180
                    Width =504
                    TabIndex =1
                    ForeColor =4210752
                    Name ="btnRefresh"
                    Caption ="Refresh List"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Refresh template list"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000a08070ff604830ff ,
                        0x604830ff604830ff604830ff604830ff604830ff604830ff604830ff604830ff ,
                        0x604830ff000000000000000000000000a08070ff604830ffa08070ffffffffff ,
                        0xb0a090ffb0a090ffb0a090ffb0a090ffb0a090ffb0a090ffb0a090ffb0a090ff ,
                        0x604830ff00000000a08070ff604830ffa08070ffffffffffa08070ffffffffff ,
                        0xfffffffffff8fffff0f0f0fff0e8e0fff0e0d0ffe0d0d0ffe0c8c0ffb0a090ff ,
                        0x604830ff00000000a08070ffffffffffa08070ffffffffffa08070ffffffffff ,
                        0xffffffffd0f0e0ff106850fff0f0f0fff0e0e0fff0d8d0ffe0d0c0ffb0a090ff ,
                        0x604830ff00000000a08070ffffffffffa08070ffffffffffa08070ffffffffff ,
                        0xffffffff209870ff209870ff209870ff209870ffc0c8c0ffe0d8d0ffb0a090ff ,
                        0x604830ff00000000a08070ffffffffffa08070ffffffffffa08870ffffffffff ,
                        0xffffffffe0f0f0ff209870fffff8f0ffc0e0d0ff209870fff0d8d0ffb0a090ff ,
                        0x604830ff00000000a08070ffffffffffa08870ffffffffffa08880ffffffffff ,
                        0xfffffffffffffffffffffffffffffffffff8f0ff209870fff0e0e0ffb0a090ff ,
                        0x604830ff00000000a08870ffffffffffa08880ffffffffffb09080ffffffffff ,
                        0xffffffff209870fffffffffffffffffffff8fffff0f0f0fff0e8e0ffb0a090ff ,
                        0x604830ff00000000a08880ffffffffffb09080ffffffffffb09080ffffffffff ,
                        0xffffffff209870ffb0d8c0ffffffffff107850ffd0e0e0fff0f0f0ffb0a090ff ,
                        0x604830ff00000000b09080ffffffffffb09080ffffffffffb09880ffffffffff ,
                        0xffffffffd0e8e0ff209870ff209870ff209870ff107850ffd0b8b0ffb0a090ff ,
                        0x604830ff00000000b09080ffffffffffb09880ffffffffffb09880ffffffffff ,
                        0xffffffffffffffffffffffffffffffff209870ffd0d8d0ffa09080ff605040ff ,
                        0x604830ff00000000b09880ffffffffffb09880ffffffffffb0a090ffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffd0b8b0ffd0c8c0ff604830ff ,
                        0xd0b0a09000000000b09880ffffffffffb0a090ffffffffffc0a090ffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffc0a8a0ff604830ffd0b0a090 ,
                        0x0000000000000000b0a090ffffffffffc0a090ffffffffffc0a090ffd0c0b0ff ,
                        0xd0c0b0ffd0c0b0ffd0b8b0ffd0b8a0ffc0b0a0ffc0a090ffd0b0a09000000000 ,
                        0x0000000000000000c0a090ffffffffffc0a090ffe0c8b0ffe0c8c0ffe0d0c0ff ,
                        0xe0d0c0ffe0d0c0ffe0d0c0ffd0b8b0ffd0b0a090000000000000000000000000 ,
                        0x0000000000000000b09890ffd0c0b0ffd0c0b0ffd0c0b0ffd0c0b0ffd0c0b0ff ,
                        0xd0b8b0ffc0b0a0ffd0b0a0900000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =6120
                    LayoutCachedTop =180
                    LayoutCachedWidth =6624
                    LayoutCachedHeight =540
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =15060409
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =5460
                    Top =180
                    Width =504
                    TabIndex =2
                    ForeColor =4210752
                    Name ="btnRunChecks"
                    Caption ="Run All Checks"
                    ControlTipText ="Run all checks"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000a08070ff604830ff ,
                        0x604830ff604830ff604830ff604830ff604830ff604830ff604830ff604830ff ,
                        0x604830ff000000000000000000000000a08070ff604830ffa08070ffffffffff ,
                        0xb0a090ffb0a090ffb0a090ffb0a090ffb0a090ffb0a090ffb0a090ffb0a090ff ,
                        0x604830ff00000000a08070ff604830ffa08070ffffffffffa08070ffffffffff ,
                        0xfffffffffff8fffff0f0f0fff0e8e0fff0e0d0ffe0d0d0ffe0c8c0ffb0a090ff ,
                        0x604830ff00000000a08070ffffffffffa08070ffffffffffa08070ffffffffff ,
                        0xffffffffd0f0e0ff106850fff0f0f0fff0e0e0fff0d8d0ffe0d0c0ffb0a090ff ,
                        0x604830ff00000000a08070ffffffffffa08070ffffffffffa08070ffffffffff ,
                        0xffffffff209870ff209870ff209870ff209870ffc0c8c0ffe0d8d0ffb0a090ff ,
                        0x604830ff00000000a08070ffffffffffa08070ffffffffffa08870ffffffffff ,
                        0xffffffffe0f0f0ff209870fffff8f0ffc0e0d0ff209870fff0d8d0ffb0a090ff ,
                        0x604830ff00000000a08070ffffffffffa08870ffffffffffa08880ffffffffff ,
                        0xfffffffffffffffffffffffffffffffffff8f0ff209870fff0e0e0ffb0a090ff ,
                        0x604830ff00000000a08870ffffffffffa08880ffffffffffb09080ffffffffff ,
                        0xffffffff209870fffffffffffffffffffff8fffff0f0f0fff0e8e0ffb0a090ff ,
                        0x604830ff00000000a08880ffffffffffb09080ffffffffffb09080ffffffffff ,
                        0xffffffff209870ffb0d8c0ffffffffff107850ffd0e0e0fff0f0f0ffb0a090ff ,
                        0x604830ff00000000b09080ffffffffffb09080ffffffffffb09880ffffffffff ,
                        0xffffffffd0e8e0ff209870ff209870ff209870ff107850ffd0b8b0ffb0a090ff ,
                        0x604830ff00000000b09080ffffffffffb09880ffffffffffb09880ffffffffff ,
                        0xffffffffffffffffffffffffffffffff209870ffd0d8d0ffa09080ff605040ff ,
                        0x604830ff00000000b09880ffffffffffb09880ffffffffffb0a090ffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffd0b8b0ffd0c8c0ff604830ff ,
                        0xd0b0a09000000000b09880ffffffffffb0a090ffffffffffc0a090ffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffc0a8a0ff604830ffd0b0a090 ,
                        0x0000000000000000b0a090ffffffffffc0a090ffffffffffc0a090ffd0c0b0ff ,
                        0xd0c0b0ffd0c0b0ffd0b8b0ffd0b8a0ffc0b0a0ffc0a090ffd0b0a09000000000 ,
                        0x0000000000000000c0a090ffffffffffc0a090ffe0c8b0ffe0c8c0ffe0d0c0ff ,
                        0xe0d0c0ffe0d0c0ffe0d0c0ffd0b8b0ffd0b0a090000000000000000000000000 ,
                        0x0000000000000000b09890ffd0c0b0ffd0c0b0ffd0c0b0ffd0c0b0ffd0c0b0ff ,
                        0xd0b8b0ffc0b0a0ffd0b0a0900000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =5460
                    LayoutCachedTop =180
                    LayoutCachedWidth =5964
                    LayoutCachedHeight =540
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =15060409
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =223
                    Left =120
                    Top =600
                    Width =720
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblPark"
                    Caption ="Park"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =600
                    LayoutCachedWidth =840
                    LayoutCachedHeight =915
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =2040
                    Top =630
                    Width =360
                    Height =315
                    FontSize =9
                    BorderColor =10921638
                    ForeColor =6750207
                    Name ="lblPlotID"
                    GridlineColor =10921638
                    LayoutCachedLeft =2040
                    LayoutCachedTop =630
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =945
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    BorderShade =65.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =600
                    Top =630
                    Width =660
                    Height =315
                    FontSize =9
                    BorderColor =10921638
                    ForeColor =6750207
                    Name ="lblParkCode"
                    GridlineColor =10921638
                    LayoutCachedLeft =600
                    LayoutCachedTop =630
                    LayoutCachedWidth =1260
                    LayoutCachedHeight =945
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    BorderShade =65.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =840
            Name ="Detail"
            OnMouseMove ="[Event Procedure]"
            AutoHeight =255
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Top =45
                    Width =360
                    Height =315
                    FontSize =9
                    BorderColor =8355711
                    ForeColor =690698
                    Name ="tbxIcon"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =45
                    LayoutCachedWidth =480
                    LayoutCachedHeight =360
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =50.0
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1020
                    Top =45
                    Width =4020
                    Height =315
                    FontSize =9
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxTemplate"
                    ControlSource ="TemplateName"
                    OnMouseMove ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =1020
                    LayoutCachedTop =45
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6780
                    Width =720
                    TabIndex =1
                    ForeColor =4210752
                    Name ="btnViewSQL"
                    Caption ="Edit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="View template SQL"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000060000000a0000000d0000000600000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000060000000f00000007000000010000000000000000000000000 ,
                        0x000000000000000000000090000000c000000080000000b00000004000000000 ,
                        0x000000a0000000e0000000e0000000b00000001000000070000000ff000000d0 ,
                        0x000000f0000000d0000000e00000003000000020000000a0000000e000000080 ,
                        0x000000f00000003000000000000000c00000009000000000000000ff00000090 ,
                        0x00000010000000b00000005000000000000000a0000000ff000000e0000000e0 ,
                        0x000000d0000000000000000000000090000000b000000000000000ff00000070 ,
                        0x000000000000002000000020000000b0000000ff000000e000000040000000f0 ,
                        0x000000b0000000000000000000000070000000ff00000000000000ff00000070 ,
                        0x000000000000000000000090000000ff000000b00000002000000020000000d0 ,
                        0x000000c0000000000000000000000090000000c000000000000000ff00000070 ,
                        0x0000000000000000000000e0000000900000000000000020000000b000000060 ,
                        0x000000f00000004000000020000000e0000000a000000000000000ff00000070 ,
                        0x000000000000000000000050000000b000000070000000d00000009000000000 ,
                        0x00000070000000a0000000c0000000800000001000000070000000ff000000c0 ,
                        0x0000002000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =6780
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =360
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =15060409
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5880
                    Width =720
                    TabIndex =4
                    ForeColor =4210752
                    Name ="btnOpenTable"
                    Caption ="Add Record"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Open tsys_Db_Templates table"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000b0a090ff604830ff604830ff604830ff604830ff ,
                        0x604830ff604830ff604830ff604830ff604830ff604830ff604830ff604830ff ,
                        0x604830ff0000000000000000b0a090ffe0c8c0ffd0c0b0ffd0b8b0ffd0b8b0ff ,
                        0xc0b0a0ffc0b0a0ffc0b0a0ffc0a8a0ffc0a890ffc0a890ffb0a090ffb0a090ff ,
                        0x604830ff0000000000000000b0a090fffffffffffffffffffff8ffffd0b8b0ff ,
                        0xfff0f0fffff0e0ffffe8e0ffc0a8a0fff0d8d0fff0d8c0fff0d0b0ffb0a090ff ,
                        0x604830ff0000000000000000b0a090ffffffffffffffffffffffffffd0c0b0ff ,
                        0xfff8f0fffff0f0fffff0e0ffc0b0a0ffffe0d0fff0d8d0fff0d8c0ffc0a890ff ,
                        0x604830ff0000000000000000b0a090ffe0d0d0ffd0c8c0ffd0c0c0ffd0c0b0ff ,
                        0xd0c0b0ffd0b8b0ffd0b8b0ffc0b0a0ffc0b0a0ffc0b0a0ffc0a8a0ffc0a890ff ,
                        0x604830ff0000000000000000c0a890ffffffffffffffffffffffffffd0c8c0ff ,
                        0xfffffffffff8fffffff8f0ffd0b8b0fffff0e0ffffe8e0ffffe0d0ffc0a8a0ff ,
                        0x604830ff0000000000000000c0a8a0ffffffffffffffffffffffffffd0c8c0ff ,
                        0xfffffffffffffffffff8ffffd0b8b0fffff0f0fffff0e0ffffe8e0ffc0a8a0ff ,
                        0x604830ff0000000000000000c0b0a0ffe0d8d0ffe0d0c0ffe0d0c0ffe0c8c0ff ,
                        0xd0c8c0ffd0c8c0ffd0c0b0ffd0c0b0ffd0b8b0ffd0b8b0ffc0b0a0ffc0b0a0ff ,
                        0x604830ff0000000000000000d0b0a0ffffffffffffffffffffffffffe0d0c0ff ,
                        0xffffffffffffffffffffffffd0c0b0fffff8fffffff8f0fffff0f0ffc0b0a0ff ,
                        0x604830ff0000000000000000d0b8a0ffffffffffffffffffffffffffe0d0c0ff ,
                        0xffffffffffffffffffffffffd0c8c0fffffffffffff8fffffff8f0ffd0b8b0ff ,
                        0x604830ff0000000000000000f0a890fff0a890fff0a890fff0a880fff0a080ff ,
                        0xe09870ffe09060ffe08850ffe08050ffe07840ffe07040ffe07040ffe07040ff ,
                        0xd06030ff0000000000000000f0a890ffffc0a0ffffc0a0ffffc0a0ffffb890ff ,
                        0xffb890ffffb090ffffa880ffffa880fff0a070fff0a070fff09870fff09860ff ,
                        0xd06830ff0000000000000000f0a890fff0a890fff0a890fff0a890fff0a880ff ,
                        0xf0a080fff09870ffe09870ffe09060ffe08860ffe08050ffe07840ffe07840ff ,
                        0xe07040ff00000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =5880
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =360
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =15060409
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =540
                    Top =60
                    Width =360
                    Height =300
                    FontSize =9
                    TabIndex =3
                    BorderColor =8355711
                    ForeColor =690698
                    Name ="tbxNumRecords"
                    ControlSource ="NumRecords"
                    GridlineColor =10921638

                    LayoutCachedLeft =540
                    LayoutCachedTop =60
                    LayoutCachedWidth =900
                    LayoutCachedHeight =360
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =50.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5220
                    Top =45
                    Width =360
                    Height =315
                    FontSize =9
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =11916796
                    Name ="tbxID"
                    ControlSource ="ID"
                    GridlineColor =10921638

                    LayoutCachedLeft =5220
                    LayoutCachedTop =45
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =9
                    ForeTint =40.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5880
                    Top =480
                    Width =720
                    TabIndex =6
                    ForeColor =4210752
                    Name ="btnRunCheck"
                    Caption ="Run"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Open tsys_Db_Templates table"
                    GridlineColor =10921638

                    LayoutCachedLeft =5880
                    LayoutCachedTop =480
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =840
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =15060409
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4080
                    Top =30
                    Width =1080
                    Height =315
                    FontSize =9
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxSQL"
                    ControlSource ="Template"
                    GridlineColor =10921638

                    LayoutCachedLeft =4080
                    LayoutCachedTop =30
                    LayoutCachedWidth =5160
                    LayoutCachedHeight =345
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =540
                    Top =480
                    Width =360
                    Height =300
                    FontSize =9
                    TabIndex =8
                    BorderColor =8355711
                    ForeColor =690698
                    Name ="tbxCheckOK"
                    GridlineColor =10921638

                    LayoutCachedLeft =540
                    LayoutCachedTop =480
                    LayoutCachedWidth =900
                    LayoutCachedHeight =780
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =50.0
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
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
' Form:         PlotCheck
' Level:        Application form
' Version:      1.01
' Basis:        Dropdown form
'
' Description:  Plot field check form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, March 22, 2017
' References:   -
' Revisions:    BLC - 3/22/2017 - 1.00 - initial version
'               BLC - 3/24/2017 - 1.01 - added CallingForm, CallingRecordID properties
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Private m_Title As String
Private m_Directions As String
Private m_CallingForm As String
Private m_CallingRecordID As Integer

'---------------------
' Event Declarations
'---------------------
Public Event InvalidTitle(Value As String)
Public Event InvalidDirections(Value As String)
Public Event InvalidCallingForm(Value As String)
Public Event InvalidCallingRecordID(Value As Integer)

'---------------------
' Properties
'---------------------
Public Property Let Title(Value As String)
    If Len(Value) > 0 Then
        m_Title = Value

        'set the form title & caption
        Me.lblTitle.Caption = m_Title
        Me.Caption = m_Title
    Else
        RaiseEvent InvalidTitle(Value)
    End If
End Property

Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Let Directions(Value As String)
    If Len(Value) > 0 Then
        m_Directions = Value

        'set the form directions
        Me.lblDirections.Caption = m_Directions
    Else
        RaiseEvent InvalidDirections(Value)
    End If
End Property

Public Property Get Directions() As String
    Directions = m_Directions
End Property

Public Property Let CallingForm(Value As String)
        m_CallingForm = Value
End Property

Public Property Get CallingForm() As String
    CallingForm = m_CallingForm
End Property

Public Property Let CallingRecordID(Value As Integer)
        m_CallingRecordID = Value
End Property

Public Property Get CallingRecordID() As Integer
    CallingRecordID = m_CallingRecordID
End Property

'---------------------
' Methods
'---------------------

' ---------------------------------
' Sub:          Form_Open
' Description:  form opening actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   JeffK, March 26, 2009
'   http://www.utteraccess.com/forum/set-height-continuous-fo-t1804798.html
' Source/date:  Bonnie Campbell, March 22, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/22/2017 - initial version
'   BLC - 3/24/2017 - set & minimize CallingForm
'   BLC - 3/27/2017 - added tbxCheckOK
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    'default
    Me.CallingForm = "frm_Data_Entry"
    Me.CallingRecordID = -1
        
    'If Len(Nz(Me.OpenArgs, "")) > 0 Then Me.CallingForm = Me.OpenArgs

    'minimize calling form
    ToggleForm Me.CallingForm, -1

    'set record
    If Len(Nz(Me.OpenArgs, "")) > 0 Then
        If InStr(Me.OpenArgs, "|") Then
            Dim ary() As String
            ary = Split(Me.OpenArgs, "|")
            Me.CallingForm = ary(0)
            Me.CallingRecordID = ary(1)
        End If
    End If

    'set park & record
    Me.lblParkCode.Caption = Nz(TempVars("ParkCode"), "")
    Me.lblPlotID.Caption = Me.CallingRecordID
    
    SetTempVar "plotID", Me.CallingRecordID
        
    Me.Caption = "Plot Check"
    lblTitle.Caption = ""
    Me.Directions = "Before you leave the plot, run these checks!"
    lblDirections.Caption = Me.Directions
    
    tbxIcon.Value = StringFromCodepoint(uLocked)
    tbxIcon.ForeColor = lngDkGreen
    lblDirections.ForeColor = lngLtBlue
    
    'set hover
'    btnDelete.HoverColor = lngGreen
'    btnViewSQL.HoverColor = lngGreen
    btnOpenTable.HoverColor = lngGreen
    btnRunCheck.HoverColor = lngGreen
'
'    btnDelete.Caption = StringFromCodepoint(uDelete)
'    btnDelete.ForeColor = lngRed

    'enable textbox to ensure scrollbar is available for longer text
    tbxTemplate.Enabled = True
    
    'cover to avoid data entry
    
    'set underlying data
    Set Me.Recordset = GetRecords("s_template_num_records")
    
    'set form height <- must be set or detail height = 1 record
    '                   due to setting recordset programmatically
    Me.InsideHeight = Me.FormHeader.Height + Me.FormFooter.Height + _
                        (Me.Detail.Height * 15)
    
    'defaults
    Me.Filter = "[FieldCheck]=" & 1
    Me.FilterOnLoad = True
    Me.AllowEdits = True
    Me.AllowFilters = True
    
    'clear num records & run queries
    RunPlotCheck
    
    Dim chk As String
    chk = StringFromCodepoint(uCheck)
    
    Me.tbxCheckOK = IIf(Me.tbxNumRecords > 0, chk, "")
    
    Me.Requery

    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[PlotCheck form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Load
' Description:  form loading actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 22, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/22/2017 - initial version
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler

    'eliminate NULLs
    If IsNull(Me.OpenArgs) Then GoTo Exit_Handler

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[PlotCheck form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Current
' Description:  form current actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/1/2016 - initial version
'   BLC - 2/1/2017 - handles giving focus to new template after TemplateAdd
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler

    If FormIsOpen("TemplateAdd") Then
    
        'go to the new record (assumes new record is last)
        'must sort & give form focus first --> sort is a toggle so do version then ID
        '                                      to give low to high ID
        Call lblVersion_Click
        Call lblHdrID_Click
        Me.SetFocus
        DoCmd.GoToRecord acDataForm, Me.Name, acLast
'        DoCmd.GoToRecord acDataForm, Me.Name, acPrevious, 30 << causes endless loop
        
        'close form
        DoCmd.Close acForm, "TemplateAdd"
    Else
'        Debug.Print "not open"
    End If
       
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[PlotCheck form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnRefresh_Click
' Description:  Open table button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, February 1, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 2/1/2017 - initial version
' ---------------------------------
Private Sub btnRefresh_Click()
On Error GoTo Err_Handler
    
    'refresh templates
    GetTemplates
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnRefresh_Click[PlotCheck form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnAddTemplate_Click
' Description:  Add template button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 4, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/4/2016 - initial version
' ---------------------------------
Private Sub btnAddTemplate_Click()
On Error GoTo Err_Handler
    
    'minimize PlotCheck
    ToggleForm "PlotCheck", -1
    
    'DoCmd.OpenTable "tsys_Db_Templates", acViewNormal, acAdd

    DoCmd.OpenForm "TemplateAdd", acNormal

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnAddTemplate_Click[PlotCheck form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnOpenTable_Click
' Description:  Open table button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, January 10, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/10/2017 - initial version
' ---------------------------------
Private Sub btnOpenTable_Click()
On Error GoTo Err_Handler
    
    'minimize PlotCheck
    ToggleForm "PlotCheck", -1
    
    DoCmd.OpenTable "tsys_Db_Templates", acViewNormal ',acAdd

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnOpenTable_Click[PlotCheck form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnViewSQL_Click
' Description:  Delete button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, September 14, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/14/2016 - initial version
' ---------------------------------
Private Sub btnViewSQL_Click()
On Error GoTo Err_Handler
    
'    Dim strOA As String
'
'    'prepare open args
'    strOA = Me.ID.Value & "|" _
'            & Me.Version.Value & "|" _
'            & Me.TemplateName.Value & "|" _
'            & Me.Template.Value & "|" _
'            & Me.EffectiveDate.Value & "|" _
'            & Me.Syntax.Value
'
'    DoCmd.OpenForm "TemplateSQL", acNormal, , , , , strOA

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnViewSQL_Click[PlotCheck form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnRunCheck_Click
' Description:  Run check button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 24, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/24/2017 - initial version
' ---------------------------------
Private Sub btnRunCheck_Click()
On Error GoTo Err_Handler
    
    'retrieve the SQL for the ID
    'Me.Template
    
    'ensure temp query is closed
    DoCmd.Close acQuery, "usys_temp_display", acSaveNo
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef, qdf2 As DAO.QueryDef
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    
    With db
        Set qdf = .QueryDefs("usys_temp_qdf")
        
        With qdf

            Dim strSQL As String
            
            'replace park code & plotID parameters
            'strSQL = Replace(Replace(Me.tbxSQL, "[pkid]", "'" & TempVars("ParkCode") & "'"), _
                        "[pid]", Me.lblPlotID.Caption)
            
            .SQL = Me.tbxSQL 'strSQL 'Me.Template 'Me.tbxSQL
            
            'set park code & plotID parameters
            .Parameters("pkid") = TempVars("ParkCode")
            .Parameters("pid") = Me.lblPlotID.Caption
            
            Debug.Print .SQL
            
            'open query window
'            DoCmd.OpenQuery qdf.Name, acViewNormal, acReadOnly '<< asks for params
'            DoCmd.RunSQL (qdf.SQL) '<< requires query statement
            
            With db
                'replace park code & plotID parameters
                strSQL = Replace(Replace(Me.tbxSQL, "[pkid]", "'" & TempVars("ParkCode") & "'"), _
                            "[pid]", Me.lblPlotID.Caption)
                
                'remove parameter clause (values already replaced)
                strSQL = right(strSQL, Len(strSQL) - InStr(strSQL, ";"))
                
 Debug.Print strSQL
                Set qdf2 = .CreateQueryDef("usys_temp_display", strSQL)
                
                DoCmd.OpenQuery "usys_temp_display", acViewNormal, acReadOnly
                
                '.QueryDefs.Delete "usys_temp_display"
            End With
            
            
            'don't .OpenRecordset here --> causes missing param errors
            'OK for SELECT/QA as long as all params are accommodated
            Set rs = .OpenRecordset()
            
            'update the number
            Dim Template As String
    
            Template = "u_num_records" '"i_num_records"
    
            Dim Params(0 To 2) As Variant
        
            With Me
                Params(0) = Template
                Params(1) = Me.tbxID
                Params(2) = rs.RecordCount
            
                SetRecord Template, Params
            End With
    
'    'after template is saved, refresh global Template dictionary
'    GetTemplates
'
'            SetRecord "i_num_records", params
'            Me.tbxNumRecords = rs.RecordCount
            
            'refresh form
            Me.Requery
            
            'run query
'            .Execute
            
            'minimize plotcheck so user can see query result
            ToggleForm "PlotCheck", -1
            
        End With
        
        'open & run query to provide parameter prompts
'        DoCmd.OpenQuery "usys_temp_qdf", acViewNormal
        
        'identify how many records are returned
'        qdf.ReturnsRecords
        
'        'minimize TemplateSQL
'        ToggleForm "TemplateSQL", -1
        
'        'close form
'        DoCmd.Close acForm, "TemplateSQL"
        
    End With

    
Exit_Handler:
    'cleanup
    Set rs = Nothing
    db.Close
    qdf.Close
    qdf2.Close
    
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnRunCheck_Click[PlotCheck form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnEdit_Click
' Description:  Enter button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 22, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/22/2017 - initial version
' ---------------------------------
Private Sub btnEdit_Click()
On Error GoTo Err_Handler
    
    'populate the parent form
'    PopulateForm Me.Parent, ID

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnEdit_Click[PlotCheck form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnDelete_Click
' Description:  Delete button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/1/2016 - initial version
' ---------------------------------
Private Sub btnDelete_Click()
On Error GoTo Err_Handler
    
    Dim result As Integer
    
    'identify the record ID
     result = MsgBox("Delete Record this record: #" & tbxID & " ?" _
                        & vbCrLf & "This action cannot be undone.", vbYesNo, "Delete Record?")

'    If result = vbYes Then DeleteRecord "Event", ID
    
    'clear the deleted record
    Me.Requery

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnDelete_Click[PlotCheck form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblHdrID_Click
' Description:  lbl click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   pere_de_chipstic, August 5, 2012
'   http://www.utteraccess.com/forum/Sort-Continuous-Form-Hea-t1991553.html
' Source/date:  Bonnie Campbell, September 13, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/13/2016 - initial version
' ---------------------------------
Private Sub lblHdrID_Click()
On Error GoTo Err_Handler

    'set the sort
'    SortListForm Me, Me.lblHdrID
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblHdrID_Click[PlotCheck form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblVersion_Click
' Description:  lbl click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   pere_de_chipstic, August 5, 2012
'   http://www.utteraccess.com/forum/Sort-Continuous-Form-Hea-t1991553.html
' Source/date:  Bonnie Campbell, September 13, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/13/2016 - initial version
' ---------------------------------
Private Sub lblVersion_Click()
On Error GoTo Err_Handler

    'set the sort
'    SortListForm Me, Me.lblVersion

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblVersion_Click[PlotCheck form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblTemplate_Click
' Description:  lbl click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   pere_de_chipstic, August 5, 2012
'   http://www.utteraccess.com/forum/Sort-Continuous-Form-Hea-t1991553.html
' Source/date:  Bonnie Campbell, September 13, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/13/2016 - initial version
' ---------------------------------
Private Sub lblTemplate_Click()
On Error GoTo Err_Handler

    'set the sort
    SortListForm Me, Me.lblTemplate

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblTemplate_Click[PlotCheck form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblSyntax_Click
' Description:  lbl click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   pere_de_chipstic, August 5, 2012
'   http://www.utteraccess.com/forum/Sort-Continuous-Form-Hea-t1991553.html
' Source/date:  Bonnie Campbell, September 13, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/13/2016 - initial version
' ---------------------------------
Private Sub lblSyntax_Click()
On Error GoTo Err_Handler

    'set the sort
'    SortListForm Me, Me.lblSyntax

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblSyntax_Click[PlotCheck form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblEffectiveDate_Click
' Description:  lbl click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   pere_de_chipstic, August 5, 2012
'   http://www.utteraccess.com/forum/Sort-Continuous-Form-Hea-t1991553.html
' Source/date:  Bonnie Campbell, September 13, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/13/2016 - initial version
' ---------------------------------
Private Sub lblEffectiveDate_Click()
On Error GoTo Err_Handler

    'set the sort
'    SortListForm Me, Me.lblEffectiveDate

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblEffectiveDate_Click[PlotCheck form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxTemplate_MouseMove
' Description:  mouse move (hover) actions
' Assumptions:  -
'               Template Name textbox is disabled, so control tips won't display
'               Otherwise this would be tbxTemplateName_MouseMove instead & tbxTemplate would
'               not be necessary
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   hnaser, March 17, 2013
'   https://www.experts-exchange.com/questions/28067200/MS-Access-tooltip-on-a-disabled-control.html
' Source/date:  Bonnie Campbell, September 13, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/13/2016 - initial version
' ---------------------------------
Private Sub tbxTemplate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

    Me.tbxTemplate.ControlTipText = Nz(FetchAddlData("tsys_Db_Templates", "Remarks", Me.tbxID)(0), "")
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxTemplate_MouseMove[PlotCheck form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Detail_MouseMove
' Description:  mouse move (hover) actions
' Assumptions:  -
'               Template Name textbox is disabled, so control tips won't display
'               Otherwise this would be tbxTemplateName_MouseMove instead & tbxControlTip would
'               not be necessary
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   hnaser, March 17, 2013
'   https://www.experts-exchange.com/questions/28067200/MS-Access-tooltip-on-a-disabled-control.html
' Source/date:  Bonnie Campbell, September 13, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/13/2016 - initial version
' ---------------------------------
Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

    Me.tbxTemplate.ControlTipText = ""
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Detail_MouseMove[PlotCheck form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Close
' Description:  form closing actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 22, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/22/2017 - initial version
'   BLC - 3/24/2017 - revise to restore calling form
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

    'close the temp query if open
    

    'restore calling form
    ToggleForm Me.CallingForm, 0
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[PlotCheck form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          RunPlotCheck
' Description:  Run plot check queries
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 27, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/27/2017 - initial version
' ---------------------------------
'Public Sub RunPlotCheck()
'On Error GoTo Err_Handler
'
'    Dim Template As String
'    Dim rs As DAO.Recordset, rs2 As DAO.Recordset
'
'    'clear num records
'    ClearTable "NumRecords"
'
'    'fetch queries
'    Template = "s_template_num_records"
'
'    Set rs = GetRecords(Template)
'
'    'iterate through records
'    If Not rs.EOF And rs.BOF Then
'        rs.MoveFirst
'        Do Until rs.EOF
'
'            'run query & retrieve record #s
'            Set rs2 = GetRecords(rs("Template"))
'
'            'add values to numrecords
'            Dim Params(0 To 2) As Variant
'
'            Params(0) = "i_num_records"
'            Params(1) = rs("ID")
'            Params(2) = rs2.RecordCount
'
'            SetRecord "i_num_records", Params
'
'            Debug.Print Params(1) & " " & rs("Template") & " " & Params(2)
'        Loop
'    End If
'
'Exit_Handler:
'    Exit Sub
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - RunPlotCheck[PlotCheck form])"
'    End Select
'    Resume Exit_Handler
'End Sub
'
