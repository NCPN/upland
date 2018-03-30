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
    AllowEdits = NotDefault
    ScrollBars =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7560
    DatasheetFontHeight =11
    ItemSuffix =17
    Left =5370
    Top =3840
    Right =13185
    Bottom =8100
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x786bd5b5d4e8e440
    End
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
    OrderByOnLoad =0
    SplitFormDatasheet =1
    FilterOnLoad =255
    OrderByOnLoad =0
    SplitFormDatasheet =1
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
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
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            BorderColor =16777215
            GridlineColor =16777215
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
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
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
        Begin ListBox
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
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =4260
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
                    OverlapFlags =93
                    Left =780
                    Top =450
                    Width =5760
                    Height =540
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblDirections"
                    Caption ="directions"
                    GridlineColor =10921638
                    LayoutCachedLeft =780
                    LayoutCachedTop =450
                    LayoutCachedWidth =6540
                    LayoutCachedHeight =990
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    Left =1200
                    Top =1080
                    Width =3960
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblTemplate"
                    Caption ="Check"
                    GridlineColor =10921638
                    LayoutCachedLeft =1200
                    LayoutCachedTop =1080
                    LayoutCachedWidth =5160
                    LayoutCachedHeight =1395
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =223
                    Left =1260
                    Top =60
                    Width =720
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblPlot"
                    Caption ="Plot #"
                    GridlineColor =10921638
                    LayoutCachedLeft =1260
                    LayoutCachedTop =60
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =375
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =223
                    Left =6360
                    Top =540
                    Width =720
                    ForeColor =4210752
                    Name ="btnAddTemplate"
                    Caption ="Add Record"
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

                    LayoutCachedLeft =6360
                    LayoutCachedTop =540
                    LayoutCachedWidth =7080
                    LayoutCachedHeight =900
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
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =5640
                    Top =540
                    Width =504
                    TabIndex =1
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

                    LayoutCachedLeft =5640
                    LayoutCachedTop =540
                    LayoutCachedWidth =6144
                    LayoutCachedHeight =900
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
                    Left =60
                    Top =60
                    Width =720
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblPark"
                    Caption ="Park"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =780
                    LayoutCachedHeight =375
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =1860
                    Top =90
                    Width =480
                    Height =315
                    FontSize =9
                    BorderColor =10921638
                    ForeColor =6750207
                    Name ="lblPlotID"
                    GridlineColor =10921638
                    LayoutCachedLeft =1860
                    LayoutCachedTop =90
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =405
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    BorderShade =65.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =540
                    Top =90
                    Width =660
                    Height =315
                    FontSize =9
                    BorderColor =10921638
                    ForeColor =6750207
                    Name ="lblParkCode"
                    GridlineColor =10921638
                    LayoutCachedLeft =540
                    LayoutCachedTop =90
                    LayoutCachedWidth =1200
                    LayoutCachedHeight =405
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    BorderShade =65.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    Left =6420
                    Top =540
                    Width =960
                    Height =345
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblRecordsReturned"
                    Caption ="# Records"
                    GridlineColor =10921638
                    LayoutCachedLeft =6420
                    LayoutCachedTop =540
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =885
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =420
                    Top =1080
                    Width =360
                    Height =285
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblID"
                    Caption ="ID"
                    GridlineColor =10921638
                    LayoutCachedLeft =420
                    LayoutCachedTop =1080
                    LayoutCachedWidth =780
                    LayoutCachedHeight =1365
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    Left =5580
                    Top =60
                    Width =1080
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblVIsitDate"
                    Caption ="Visit Date"
                    GridlineColor =10921638
                    LayoutCachedLeft =5580
                    LayoutCachedTop =60
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =375
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    Left =6660
                    Top =90
                    Width =840
                    Height =315
                    FontSize =9
                    BorderColor =10921638
                    ForeColor =6750207
                    Name ="lblSampleDate"
                    GridlineColor =10921638
                    LayoutCachedLeft =6660
                    LayoutCachedTop =90
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =405
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    BorderShade =65.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3840
                    Top =60
                    Width =900
                    Height =255
                    ColumnOrder =2
                    FontSize =8
                    TabIndex =2
                    BackColor =-2147483643
                    ForeColor =12566463
                    Name ="tbxDevMode"
                    FontName ="MS Sans Serif"
                    ConditionalFormat = Begin
                        0x010000006e000000010000000000000002000000000000000600000001000000 ,
                        0x3f3f3f00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x460061006c007300650000000000
                    End

                    LayoutCachedLeft =3840
                    LayoutCachedTop =60
                    LayoutCachedWidth =4740
                    LayoutCachedHeight =315
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x0100010000000000000002000000010000003f3f3f00ffffff00050000004600 ,
                        0x61006c0073006500000000000000000000000000000000000000000000
                    End
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =1065
                    Width =360
                    Height =288
                    ColumnOrder =1
                    FontWeight =500
                    TabIndex =3
                    BackColor =-2147483643
                    ForeColor =14277081
                    Name ="tbxCurrentRecord"
                    ConditionalFormat = Begin
                        0x0100000094000000010000000100000000000000000000001900000001000000 ,
                        0x00000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00740062007800430075007200720065006e0074005200650063006f007200 ,
                        0x64005d003d00460061006c007300650000000000
                    End

                    LayoutCachedLeft =60
                    LayoutCachedTop =1065
                    LayoutCachedWidth =420
                    LayoutCachedHeight =1353
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =85.0
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ffffff00180000005b00 ,
                        0x740062007800430075007200720065006e0074005200650063006f0072006400 ,
                        0x5d003d00460061006c0073006500000000000000000000000000000000000000 ,
                        0x000000
                    End
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =215
                    MultiSelect =2
                    IMESentenceMode =3
                    ColumnCount =7
                    Left =480
                    Top =1380
                    Width =6420
                    Height =2040
                    ColumnOrder =0
                    TabIndex =4
                    ForeColor =4210752
                    BorderColor =10921638
                    Name ="lbxTemplates"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="[Event Procedure]"
                    LeftPadding =720
                    GridlineColor =10921638

                    LayoutCachedLeft =480
                    LayoutCachedTop =1380
                    LayoutCachedWidth =6900
                    LayoutCachedHeight =3420
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6420
                    Top =3540
                    Width =960
                    Height =600
                    TabIndex =5
                    Name ="btnRunPlotChecks"
                    Caption ="Plot Check!"
                    StatusBarText ="Run plot checks!"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00ddddddddddddddddddd0000000000ddddd0ffffffffff0dd ,
                        0xdd0fff88fffff0dddd0ff8188ffff0dddd0f811188fff0dddd0f11f118fff0dd ,
                        0xdd0fffff178ff0dddd0ffffff188f0dddd0fffffff18f0dddd0ffffffff1f0dd ,
                        0xdd0ffffffffff0dddd0ff000000ff0ddddd000f888000ddddddddd0000dddddd ,
                        0xdddddddddddddddd
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Run plot checks!"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638

                    CursorOnHover =1
                    LayoutCachedLeft =6420
                    LayoutCachedTop =3540
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =4140
                    ForeTint =100.0
                    Shape =2
                    BackColor =5066944
                    BackThemeColorIndex =5
                    BackTint =100.0
                    BorderColor =5066944
                    BorderThemeColorIndex =5
                    BorderTint =100.0
                    HoverColor =15709952
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =15709952
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =24
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
            End
        End
        Begin Section
            Height =0
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
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
' Form:         PlotCheckSelect
' Level:        Application form
' Version:      1.01
' Basis:        Dropdown form
'
' Description:  Plot field check selection form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, January 30, 2018
' References:   -
' Revisions:    BLC - 1/30/2018 - 1.00 - initial version
'               BLC - 3/30/2018 - 1.01 - remove form detail references
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
Private m_CallingSampleDate As Date

Private m_SelectedChecks As Collection
Private m_SelectedCheck As Long

'---------------------
' Event Declarations
'---------------------
Public Event InvalidTitle(Value As String)
Public Event InvalidDirections(Value As String)
Public Event InvalidCallingForm(Value As String)
Public Event InvalidCallingRecordID(Value As Integer)
Public Event InvalidCallingSampleDate(Value As Date)

Public Event InvalidSelCheck(Value As Long)

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

Public Property Let CallingSampleDate(Value As Date)
        m_CallingSampleDate = Value
End Property

Public Property Get CallingSampleDate() As Date
    CallingSampleDate = m_CallingSampleDate
End Property

Public Property Let SelectedChecks(Value As Collection)
        Set m_SelectedChecks = Value
End Property

Public Property Get SelectedChecks() As Collection
    Set SelectedChecks = m_SelectedChecks
End Property

Public Property Let SelectedCheck(Value As Long)
    If IsNumeric(Value) Then
        m_SelectedCheck = Value
    Else
        RaiseEvent InvalidSelCheck(Value)
    End If
    
    'check if value is already present
    Dim InCollection As Boolean
    InCollection = False
    Dim i As Long
    
    For i = 1 To Me.SelectedChecks.Count
        If SelectedChecks.Item(i) = Value Then
            InCollection = True
            Exit For
        End If
    Next
    
    If InCollection = False Then
        'add to the collection
        Me.SelectedChecks.Add Value
    End If
    
End Property

Public Property Get SelectedCheck() As Long
    SelectedCheck = m_SelectedCheck
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
'   BLC - 3/30/2017 - hid unfiltered query num records
'   BLC - 3/31/2017 - added CallingSampleDate property
'   BLC - 8/11/2017 - removed btnRunCheck & revised directions to use
'                     double click on check name
'   BLC - 3/30/2018 - remove form detail references
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    'default
    Me.CallingForm = "frm_Data_Entry"
    Me.CallingRecordID = -1
    Me.CallingSampleDate = Date
        
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
            Me.CallingSampleDate = ary(2)
        End If
    End If

    'set park & record
    Me.lblParkCode.Caption = Nz(TempVars("ParkCode"), "")
    Me.lblPlotID.Caption = Me.CallingRecordID
    Me.lblSampleDate.Caption = Me.CallingSampleDate
    
    SetTempVar "plotID", Me.CallingRecordID
    SetTempVar "SampleDate", Me.CallingSampleDate
    
    Me.Caption = "Select Plot Check(s)"
    lblTitle.Caption = ""
    Me.Directions = "Choose the plot checks to run and click Run." _
        & vbCrLf & "Use SHFT or CTRL to select more than one."
    lblDirections.Caption = Me.Directions
    
    'tbxIcon.Value = StringFromCodepoint(uLocked)
    'tbxIcon.ForeColor = lngDkGreen
    lblDirections.ForeColor = lngLtBlue
    
    'set hover
    btnRunPlotChecks.HoverColor = lngGreen

    'default
    btnRunPlotChecks.Enabled = False

    'enable textbox to ensure scrollbar is available for longer text
    'tbxTemplate.Enabled = True
        
    'set underlying data
    Set Me.Recordset = GetRecords("s_template_num_records")
    
    'prep listbox
    'relevant columns: ID, TemplateName
    Set lbxTemplates.Recordset = GetRecords("s_template_num_records")
    With lbxTemplates
        .ColumnHeads = False 'hide & use labels instead (easier formatting)
        '.MultiSelect = 2 '0-None, 1-Simple, 2-Extended
        .ColumnCount = 14
        .ColumnWidths = ".5in;0;0;0;0;0;1.5in;0;0;0;0;0;0;0"
        .GridlineStyleBottom = 1 '1-solid
        .GridlineWidthBottom = 1 '1 pt
        .LeftPadding = 2 'pt scale
        .FontSize = 9
    End With
    
    'set form height <- must be set or detail height = 1 record
    '                   due to setting recordset programmatically
    Me.InsideHeight = Me.FormHeader.Height + Me.FormFooter.Height + _
                        (Me.Detail.Height * 10)
    
    'defaults
    Me.Filter = "[FieldCheck]=" & 1
    Me.FilterOnLoad = True
    Me.AllowEdits = True
    Me.AllowFilters = True
    
    
'    'hide dev mode so it doesn't flash w/ @ transect
'    If Not DEV_MODE Then Me.tbxDevMode.Visible = False
    
    'set dev mode
    Me.tbxDevMode = DEV_MODE
    
    'clear num records & run queries
    RunPlotCheck
    
    Dim chk As String
    chk = StringFromCodepoint(uCheck)
    
    'Me.tbxCheckOK = IIf(Me.tbxNumRecords > 0, chk, "")
    
    'hide initial unfiltered query record #s
    'Me.tbxNumRecords.Visible = True 'False
    
    Me.Requery

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Form_Open[Form_PlotCheckSelect])"
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
            "Error encountered (#" & Err.Number & " - Form_Load[PlotCheckSelect form])"
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
' References:
'   Michael S. Meyers-Jouan, January 27, 2012
'   http://database.ittoolbox.com/groups/technical-functional/access-l/highlighted-field-on-open-form-4618567
' Source/date:  Bonnie Campbell, June 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/1/2016 - initial version
'   BLC - 2/1/2017 - handles giving focus to new template after TemplateAdd
'   BLC - 3/28/2017 - clear unused code for uplands
'   BLC - 8/9/2017 - prevent focus & selection of textbox
'   BLC - 8/10/2017 - added current record highlight (via conditional format & tbxCurrentRecord)
'   BLC - 8/14/2017 - removed btnRunCheck, tbxNoRunCheck
'   BLC - 3/30/2018 - remove form detail references
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler

    'set current record
    'Me.tbxCurrentRecord = Me.tbxID

    'prevent focus/select on query name (n.b. cannot focus on btnRunCheck > Error #2110)
    'Me.tbxNumRecords.SetFocus
    'Me.tbxNumRecords.SelStart = 0
    'Me.tbxNumRecords.SelLength = 0

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[PlotCheckSelect form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lbxTemplates_AfterUpdate
' Description:  Templates listbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, February 1, 2018 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 2/1/2018 - initial version
' ---------------------------------
Private Sub lbxTemplates_AfterUpdate()
On Error GoTo Err_Handler
    
    Dim check As Variant
    
    'clear list & repopulate
    Me.SelectedChecks = New Collection
    
    'check for selections
    If lbxTemplates.ItemsSelected.Count > 0 Then
        btnRunPlotChecks.Enabled = True
    Else
        btnRunPlotChecks.Enabled = False
        GoTo Exit_Handler
    End If
        
    'repopulate list
    For Each check In lbxTemplates.ItemsSelected
    
'        Debug.Print lbxTemplates.ItemData(check)
    
        Me.SelectedCheck = lbxTemplates.ItemData(check)
    
    Next

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxTemplates_AfterUpdate[PlotCheckSelect form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnRunPlotChecks_Click
' Description:  Run plot checks button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, February 1, 2018 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 2/1/2018 - initial version
' ---------------------------------
Private Sub btnRunPlotChecks_Click()
On Error GoTo Err_Handler
    
    'convert object to pointer
    Dim chks As Long
    
    chks = GetPointer(Me.SelectedChecks)
    
    DoCmd.OpenForm "PlotCheck", acNormal, , , , acWindowNormal, Me.Name & _
                                                            "|" & Me.lblPlotID.Caption & _
                                                            "|" & Me.lblSampleDate.Caption & _
                                                            "|" & chks

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnRunPlotChecks_Click[PlotCheckSelect form])"
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
    CloseObject "usys_temp_display", "qry"

'    'remove template queries
'    RemoveTemplateQueries

    'restore calling form
    ToggleForm Me.CallingForm, 0
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[PlotCheckSelect form])"
    End Select
    Resume Exit_Handler
End Sub
