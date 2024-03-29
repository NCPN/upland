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
    AllowEdits = NotDefault
    ScrollBars =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7560
    DatasheetFontHeight =11
    ItemSuffix =13
    Left =5370
    Top =4740
    Right =13185
    Bottom =10380
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
                    OverlapFlags =85
                    Left =1080
                    Top =1080
                    Width =3960
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblTemplate"
                    Caption ="Check"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =1080
                    LayoutCachedTop =1080
                    LayoutCachedWidth =5040
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
                    Top =840
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
                    LayoutCachedTop =840
                    LayoutCachedWidth =7080
                    LayoutCachedHeight =1200
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
                    Top =840
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
                    LayoutCachedTop =840
                    LayoutCachedWidth =6144
                    LayoutCachedHeight =1200
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
                    OverlapFlags =215
                    TextAlign =2
                    Left =6240
                    Top =1080
                    Width =960
                    Height =345
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblRecordsReturned"
                    Caption ="# Records"
                    GridlineColor =10921638
                    LayoutCachedLeft =6240
                    LayoutCachedTop =1080
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =1425
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
                    OnClick ="[Event Procedure]"
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
                    ColumnOrder =1
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
                    ColumnOrder =0
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
            End
        End
        Begin Section
            Height =420
            Name ="Detail"
            OnMouseMove ="[Event Procedure]"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    BorderWidth =1
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Width =7560
                    Height =420
                    FontSize =9
                    TabIndex =7
                    BorderColor =2366701
                    ForeColor =4138256
                    Name ="tbxRecord"
                    ConditionalFormat = Begin
                        0x01000000b4000000010000000100000000000000000000002900000001000000 ,
                        0x00000000f1fa7600000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x28005b00490044005d003d005b00740062007800430075007200720065006e00 ,
                        0x74005200650063006f00720064005d00290020004f00720020005b004e006500 ,
                        0x77005200650063006f00720064005d0000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedWidth =7560
                    LayoutCachedHeight =420
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000f1fa7600280000002800 ,
                        0x5b00490044005d003d005b00740062007800430075007200720065006e007400 ,
                        0x5200650063006f00720064005d00290020004f00720020005b004e0065007700 ,
                        0x5200650063006f00720064005d00000000000000000000000000000000000000 ,
                        0x000000
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    BackStyle =0
                    IMESentenceMode =3
                    Top =30
                    Width =360
                    Height =315
                    FontSize =9
                    BorderColor =8355711
                    ForeColor =690698
                    Name ="tbxIcon"
                    GridlineColor =10921638

                    LayoutCachedTop =30
                    LayoutCachedWidth =360
                    LayoutCachedHeight =345
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =50.0
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1140
                    Top =30
                    Width =4020
                    Height =315
                    FontSize =9
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxTemplate"
                    ControlSource ="TemplateName"
                    OnDblClick ="[Event Procedure]"
                    OnMouseMove ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000e0000000030000000100000000000000000000001b00000001010000 ,
                        0x00000000ffffff0001000000000000001c0000002d00000001000000ed1c2400 ,
                        0xffffff0001000000000000002e0000003f0000000100000022b14c00ffffff00 ,
                        0x5b00740062007800490044005d003d005b007400620078004300750072007200 ,
                        0x65006e0074005200650063006f00720064005d00000000005b00460069006500 ,
                        0x6c00640043006800650063006b004f004b005d003d003000000000005b004600 ,
                        0x690065006c00640043006800650063006b004f004b005d003d00310000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1140
                    LayoutCachedTop =30
                    LayoutCachedWidth =5160
                    LayoutCachedHeight =345
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x01000300000001000000000000000101000000000000ffffff001a0000005b00 ,
                        0x740062007800490044005d003d005b0074006200780043007500720072006500 ,
                        0x6e0074005200650063006f00720064005d000000000000000000000000000000 ,
                        0x00000000000000010000000000000001000000ed1c2400ffffff00100000005b ,
                        0x004600690065006c00640043006800650063006b004f004b005d003d00300000 ,
                        0x0000000000000000000000000000000000000000010000000000000001000000 ,
                        0x22b14c00ffffff00100000005b004600690065006c0064004300680065006300 ,
                        0x6b004f004b005d003d0031000000000000000000000000000000000000000000 ,
                        0x00
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6360
                    Top =45
                    Width =600
                    Height =300
                    FontSize =9
                    TabIndex =2
                    BorderColor =8355711
                    ForeColor =690698
                    Name ="tbxNumRecords"
                    ControlSource ="NumRecords"
                    ConditionalFormat = Begin
                        0x01000000e0000000030000000100000000000000000000001b00000001010000 ,
                        0x00000000ffffff0001000000000000001c0000002d00000001000000ff000000 ,
                        0xffffff0001000000000000002e0000003f0000000100000022b14c00ffffff00 ,
                        0x5b00740062007800490044005d003d005b007400620078004300750072007200 ,
                        0x65006e0074005200650063006f00720064005d00000000005b00460069006500 ,
                        0x6c00640043006800650063006b004f004b005d003d003000000000005b004600 ,
                        0x690065006c00640043006800650063006b004f004b005d003d00310000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =6360
                    LayoutCachedTop =45
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =345
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =50.0
                    ConditionalFormat14 = Begin
                        0x01000300000001000000000000000101000000000000ffffff001a0000005b00 ,
                        0x740062007800490044005d003d005b0074006200780043007500720072006500 ,
                        0x6e0074005200650063006f00720064005d000000000000000000000000000000 ,
                        0x00000000000000010000000000000001000000ff000000ffffff00100000005b ,
                        0x004600690065006c00640043006800650063006b004f004b005d003d00300000 ,
                        0x0000000000000000000000000000000000000000010000000000000001000000 ,
                        0x22b14c00ffffff00100000005b004600690065006c0064004300680065006300 ,
                        0x6b004f004b005d003d0031000000000000000000000000000000000000000000 ,
                        0x00
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =420
                    Top =30
                    Width =360
                    Height =315
                    FontSize =9
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =15921906
                    Name ="tbxID"
                    ControlSource ="ID"
                    ConditionalFormat = Begin
                        0x01000000e0000000030000000100000000000000000000001b00000001010000 ,
                        0x00000000ffffff0001000000000000001c0000002d00000001000000ff000000 ,
                        0xffffff0001000000000000002e0000003f0000000100000022b14c00ffffff00 ,
                        0x5b00740062007800490044005d003d005b007400620078004300750072007200 ,
                        0x65006e0074005200650063006f00720064005d00000000005b00460069006500 ,
                        0x6c00640043006800650063006b004f004b005d003d003000000000005b004600 ,
                        0x690065006c00640043006800650063006b004f004b005d003d00310000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =420
                    LayoutCachedTop =30
                    LayoutCachedWidth =780
                    LayoutCachedHeight =345
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =95.0
                    ConditionalFormat14 = Begin
                        0x01000300000001000000000000000101000000000000ffffff001a0000005b00 ,
                        0x740062007800490044005d003d005b0074006200780043007500720072006500 ,
                        0x6e0074005200650063006f00720064005d000000000000000000000000000000 ,
                        0x00000000000000010000000000000001000000ff000000ffffff00100000005b ,
                        0x004600690065006c00640043006800650063006b004f004b005d003d00300000 ,
                        0x0000000000000000000000000000000000000000010000000000000001000000 ,
                        0x22b14c00ffffff00100000005b004600690065006c0064004300680065006300 ,
                        0x6b004f004b005d003d0031000000000000000000000000000000000000000000 ,
                        0x00
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5340
                    Top =15
                    Width =1080
                    Height =315
                    FontSize =9
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="tbxSQL"
                    ControlSource ="Template"
                    GridlineColor =10921638

                    LayoutCachedLeft =5340
                    LayoutCachedTop =15
                    LayoutCachedWidth =6420
                    LayoutCachedHeight =330
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7020
                    Top =45
                    Width =360
                    Height =300
                    FontSize =9
                    TabIndex =5
                    BorderColor =8355711
                    ForeColor =690698
                    Name ="tbxCheckOK"
                    GridlineColor =10921638

                    LayoutCachedLeft =7020
                    LayoutCachedTop =45
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =345
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =50.0
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6840
                    Top =15
                    Width =360
                    Height =315
                    FontSize =9
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =15921906
                    Name ="tbxFieldCheckOK"
                    ControlSource ="FieldCheckOK"
                    GridlineColor =10921638

                    LayoutCachedLeft =6840
                    LayoutCachedTop =15
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =330
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =95.0
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
' Version:      1.09
' Basis:        Dropdown form
'
' Description:  Plot field check form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, March 22, 2017
' References:   -
' Revisions:    BLC - 3/22/2017 - 1.00 - initial version
'               BLC - 3/24/2017 - 1.01 - added CallingForm, CallingRecordID properties
'               BLC - 3/28/2017 - 1.02 - removed unused click events (btnAdd,
'                                        btnDelete, btnEdit, lblHdr, lblVersion)
'               BLC - 3/30/2017 - 1.03 - added lblID_Click, revised RunCheck(),
'                                        updated checks
'               BLC - 3/31/2017 - 1.04 - added CallingSampleDate property
'               BLC - 4/3/2017  - 1.05 - code cleanup
'               BLC - 8/10/2017 - 1.06 - change from CurrentDb to CurrDb property to reduce pointers
'                                        revise to minimize vs. open query
'                                        added current record highlighting & tbxTemplate double click
'                                        to run query
'               BLC - 8/11/2017 - 1.07 - removed btnRunCheck & revised directions to use
'                                        double click on check name
'               BLC - 8/14/2017 - 1.08 - redo error handling for Check double click to address error 3048
'                                        remove btnRunCheck, tbxRunCheck methods
'               BLC - 2/1/2018  - 1.09 - revised to display only selected checks from PlotCheckSelect
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

'---------------------
' Event Declarations
'---------------------
Public Event InvalidTitle(Value As String)
Public Event InvalidDirections(Value As String)
Public Event InvalidCallingForm(Value As String)
Public Event InvalidCallingRecordID(Value As Integer)
Public Event InvalidCallingSampleDate(Value As Date)

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
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    'default
    Me.CallingForm = "PlotCheckSelect" '"frm_Data_Entry"
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
            
            Dim ptr As Long
            ptr = ary(3)
            
            Me.SelectedChecks = GetObject(ptr)
        End If
    Else
        GoTo Exit_Handler
    End If

    'set park & record
    Me.lblParkCode.Caption = Nz(TempVars("ParkCode"), "")
    Me.lblPlotID.Caption = Me.CallingRecordID
    Me.lblSampleDate.Caption = Me.CallingSampleDate
    
    SetTempVar "plotID", Me.CallingRecordID
    SetTempVar "SampleDate", Me.CallingSampleDate
    
    Me.Caption = "Plot Check"
    lblTitle.Caption = ""
    Me.Directions = "The following plot checks have been run." _
        & vbCrLf & "To re-run && view results double click the check name."
    lblDirections.Caption = Me.Directions
    
    tbxIcon.Value = StringFromCodepoint(uLocked)
    tbxIcon.ForeColor = lngDkGreen
    lblDirections.ForeColor = lngLtBlue
    
    'set hover
    'btnRunCheck.HoverColor = lngGreen

    'enable textbox to ensure scrollbar is available for longer text
    tbxTemplate.Enabled = True
        
    'set underlying data
    Set Me.Recordset = GetRecords("s_template_num_records")
    
    'set form height <- must be set or detail height = 1 record
    '                   due to setting recordset programmatically
    Me.InsideHeight = Me.FormHeader.Height + Me.FormFooter.Height + _
                        (Me.Detail.Height * 10)
    
    'prepare checks for filter
    Dim schk As Variant
    Dim chks As String
    
    chks = ""
    
    For Each schk In Me.SelectedChecks
        chks = chks & IIf(Len(chks) > 0, "," & schk, schk)
    Next
    
    'defaults
'    Me.Filter = "[FieldCheck]=" & 1
    Me.Filter = "[FieldCheck]=" & 1 & " AND [ID] IN (" & chks & ")"
    Me.FilterOn = True
    Me.FilterOnLoad = True
    Me.AllowEdits = True
    Me.AllowFilters = True
    
'    'hide dev mode so it doesn't flash w/ @ transect
'    If Not DEV_MODE Then Me.tbxDevMode.Visible = False
    
    'set dev mode
    Me.tbxDevMode = DEV_MODE
    
    'clear num records & run queries
    RunPlotCheck Me.SelectedChecks
    
    Dim chk As String
    chk = StringFromCodepoint(uCheck)
    
    Me.tbxCheckOK = IIf(Me.tbxNumRecords > 0, chk, "")
    
    'hide initial unfiltered query record #s
    Me.tbxNumRecords.Visible = True 'False
    
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
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler

    'set current record
    Me.tbxCurrentRecord = Me.tbxID

    'prevent focus/select on query name (n.b. cannot focus on btnRunCheck > Error #2110)
    Me.tbxNumRecords.SetFocus
    Me.tbxNumRecords.SelStart = 0
    Me.tbxNumRecords.SelLength = 0

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
' Sub:          tbxTemplate_DblClick
' Description:  template double click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Steve Schapel, September 15, 2008
'   https://www.pcreview.co.uk/threads/switch-focus-to-query-through-vba.3622059/
' Source/date:  Bonnie Campbell, March 24, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 8/10/2017 - initial version
'   BLC - 8/14/2017 - redo error handling to address error 3048, add status messaging
' ---------------------------------
Private Sub tbxTemplate_DblClick(Cancel As Integer)
On Error GoTo Err_Handler
    
    Dim db As dao.Database
    Dim qdf As dao.QueryDef, qdf2 As dao.QueryDef
    Dim rs As dao.Recordset
    Dim PlotID As Integer
    Dim ParkCode As String, fltr As String
    
    'clear status messages
    Application.SysCmd acSysCmdClearStatus
    
    Set db = CurrDb
    
    With db
        Set qdf = .QueryDefs("usys_temp_qdf")
        
        With qdf

            Dim strSQL As String
            Dim IsParameterized As Boolean
            
            'default
            IsParameterized = False

            'set values
'            ParkCode = TempVars("ParkCode")
            PlotID = Me.lblPlotID.Caption
'            SampleDate = Me.lblSampleDate.Caption
            
            .SQL = Me.tbxSQL
            strSQL = .SQL
            
            'open query window
            With db
                
                If QueryExists("usys_temp_display") Then
                    'ensure temp query is closed & removed
                    DoCmd.Close acQuery, "usys_temp_display", acSaveNo
                    
                    'remove usys_temp_display if it already exists
                    If Not db.QueryDefs("usys_temp_display") Is Nothing Then _
                        DoCmd.DeleteObject acQuery, "usys_temp_display"
                End If
                 
                'limit query by park & plot
                If Len(strSQL) > Len(Replace(strSQL, "PARAMETERS", "")) Then

                    'replace park code & plotID parameters
                    strSQL = Replace( _
                             Replace( _
                             Replace(strSQL, "[pkcode]", "'" & TempVars("ParkCode") & "'"), _
                                "[pid]", PlotID), _
                                "[vdate]", "#" & TempVars("SampleDate") & "#")

                    'remove parameter clause (values already replaced)
                    strSQL = Right(strSQL, Len(strSQL) - InStr(strSQL, ";"))
                                    
                    Set qdf2 = .CreateQueryDef("usys_temp_display", strSQL)

                End If
                                                                
                'display results
'                DoCmd.OpenForm "PlotCheckResults", acNormal << empty results
                                
                'minimize form
                DoCmd.Minimize
                                
                DoCmd.OpenQuery "usys_temp_display", acViewNormal, acReadOnly
'                 DoCmd.OpenForm "QueryView", acFormDS, , , acFormReadOnly, acWindowNormal << bumps into #3048-too many dbs open (> 2048 IDs)
            End With

        End With
                
    End With

    
Exit_Handler:
    'cleanup
    Set rs = Nothing
    'db.Close
    qdf.Close
    qdf2.Close
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 3048
        Debug.Print "Error #" & Err.Number & " " & Err.Description & " " & Err.Source
        'need to save the data first, so close the forms & re-open
'        DoCmd.Close acForm, Me.Name, acSaveNo
'        DoCmd.Close acForm, "frm_Data_Entry", acSaveYes
'        're-open forms
'        DoCmd.OpenForm "frm_Data_Entry", acNormal

'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - tbxTemplate_DblClick[PlotCheck form])"
'        DoCmd.SelectObject acForm, "PlotCheck"
'        DoCmd.Restore
        Application.SysCmd acSysCmdSetStatus, "Missing species check info, please close plot check & data entry forms & re-check..."
        Resume Next
      Case 3270
        'Error 3270 Property not found < caused by attempting to run query &
        '                                getting Error 3048
        Debug.Print "Error #" & Err.Number & " " & Err.Description & " " & Err.Source
        'Err.Raise 3048
        Application.SysCmd acSysCmdSetStatus, "Missing species check info, please close plot check & data entry forms & re-check..."
'        DoCmd.SelectObject acForm, "PlotCheck"
'        DoCmd.Restore
        Resume Next
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxTemplate_Click[PlotCheck form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblID_Click
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
Private Sub lblID_Click()
On Error GoTo Err_Handler

    'set the sort
    SortListForm Me, Me.lblID

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblID_Click[PlotCheck form])"
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
Private Sub tbxTemplate_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
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
Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
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
            "Error encountered (#" & Err.Number & " - Form_Close[PlotCheck form])"
    End Select
    Resume Exit_Handler
End Sub
