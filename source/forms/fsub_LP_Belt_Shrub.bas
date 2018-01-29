Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11880
    DatasheetFontHeight =9
    ItemSuffix =43
    Left =-705
    Top =945
    Right =11415
    Bottom =4485
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x9aa5143d6c56e340
    End
    RecordSource ="tbl_LP_Shrub"
    Caption ="fsub_LP_Belt_Shrub"
    BeforeInsert ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    FetchDefaults =0
    FetchDefaults =0
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
            Height =1200
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Rectangle
                    BackStyle =1
                    OverlapFlags =93
                    Left =6300
                    Top =660
                    Width =1008
                    Height =540
                    BackColor =13434828
                    Name ="rct4"
                    LayoutCachedLeft =6300
                    LayoutCachedTop =660
                    LayoutCachedWidth =7308
                    LayoutCachedHeight =1200
                End
                Begin Rectangle
                    BackStyle =1
                    OverlapFlags =93
                    Left =8580
                    Top =660
                    Width =1008
                    Height =540
                    BackColor =13434828
                    Name ="rct6"
                    LayoutCachedLeft =8580
                    LayoutCachedTop =660
                    LayoutCachedWidth =9588
                    LayoutCachedHeight =1200
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =6735
                    Top =735
                    Width =195
                    Height =240
                    FontSize =5
                    FontWeight =700
                    BackColor =13434828
                    Name ="lbl4"
                    Caption ="4"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =6735
                    LayoutCachedTop =735
                    LayoutCachedWidth =6930
                    LayoutCachedHeight =975
                End
                Begin Rectangle
                    BackStyle =1
                    OverlapFlags =93
                    Left =4200
                    Top =660
                    Width =1008
                    Height =540
                    BackColor =13434828
                    Name ="rct2"
                    LayoutCachedLeft =4200
                    LayoutCachedTop =660
                    LayoutCachedWidth =5208
                    LayoutCachedHeight =1200
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =240
                    Top =720
                    Width =1320
                    Height =240
                    FontWeight =700
                    Name ="Species_Label"
                    Caption ="Shrub Species"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    Left =2520
                    Top =720
                    Width =720
                    Height =240
                    FontWeight =700
                    Name ="Alive_Label"
                    Caption ="Alive?"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =3300
                    Top =960
                    Width =705
                    Height =240
                    FontSize =5
                    FontWeight =700
                    Name ="HC10_Label"
                    Caption ="0-10cm"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =3300
                    LayoutCachedTop =960
                    LayoutCachedWidth =4005
                    LayoutCachedHeight =1200
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =4200
                    Top =960
                    Width =975
                    Height =240
                    FontSize =5
                    FontWeight =700
                    BackColor =13434828
                    Name ="HC25_Label"
                    Caption ="10.1-25cm"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =4200
                    LayoutCachedTop =960
                    LayoutCachedWidth =5175
                    LayoutCachedHeight =1200
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =5265
                    Top =960
                    Width =975
                    Height =240
                    FontSize =5
                    FontWeight =700
                    Name ="HC50_Label"
                    Caption ="25.1-50cm"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =5265
                    LayoutCachedTop =960
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =1200
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =6292
                    Top =960
                    Width =1080
                    Height =240
                    FontSize =5
                    FontWeight =700
                    BackColor =13434828
                    Name ="HC100_Label"
                    Caption ="50.1-100cm"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =6292
                    LayoutCachedTop =960
                    LayoutCachedWidth =7372
                    LayoutCachedHeight =1200
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =7530
                    Top =960
                    Width =765
                    Height =240
                    FontSize =5
                    FontWeight =700
                    Name ="HC2m_Label"
                    Caption ="1.01-2m"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =7530
                    LayoutCachedTop =960
                    LayoutCachedWidth =8295
                    LayoutCachedHeight =1200
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =8640
                    Top =960
                    Width =705
                    Height =240
                    FontSize =5
                    FontWeight =700
                    BackColor =13434828
                    Name ="HCGT2_Label"
                    Caption =">2.01m"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =8640
                    LayoutCachedTop =960
                    LayoutCachedWidth =9345
                    LayoutCachedHeight =1200
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =3060
                    Top =480
                    Width =6480
                    Height =240
                    FontWeight =700
                    BackColor =14277081
                    Name ="lblHeightClassTotals"
                    Caption ="Height Class Totals"
                    LayoutCachedLeft =3060
                    LayoutCachedTop =480
                    LayoutCachedWidth =9540
                    LayoutCachedHeight =720
                    BackThemeColorIndex =1
                    BackShade =85.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2460
                    Top =60
                    Width =6480
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label23"
                    Caption ="Number of Live Shrubs Rooted in 1 Meter Belt Transect"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =3495
                    Top =735
                    Width =195
                    Height =240
                    FontSize =5
                    FontWeight =700
                    Name ="lbl1"
                    Caption ="1"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =3495
                    LayoutCachedTop =735
                    LayoutCachedWidth =3690
                    LayoutCachedHeight =975
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =4575
                    Top =735
                    Width =195
                    Height =240
                    FontSize =5
                    FontWeight =700
                    BackColor =13434828
                    Name ="lbl2"
                    Caption ="2"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =4575
                    LayoutCachedTop =735
                    LayoutCachedWidth =4770
                    LayoutCachedHeight =975
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =5655
                    Top =735
                    Width =195
                    Height =240
                    FontSize =5
                    FontWeight =700
                    Name ="lbl3"
                    Caption ="3"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =5655
                    LayoutCachedTop =735
                    LayoutCachedWidth =5850
                    LayoutCachedHeight =975
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =7815
                    Top =735
                    Width =195
                    Height =240
                    FontSize =5
                    FontWeight =700
                    Name ="lbl5"
                    Caption ="5"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =7815
                    LayoutCachedTop =735
                    LayoutCachedWidth =8010
                    LayoutCachedHeight =975
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =8895
                    Top =735
                    Width =195
                    Height =240
                    FontSize =5
                    FontWeight =700
                    BackColor =13434828
                    Name ="lbl6"
                    Caption ="6"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =8895
                    LayoutCachedTop =735
                    LayoutCachedWidth =9090
                    LayoutCachedHeight =975
                End
            End
        End
        Begin Section
            Height =420
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Rectangle
                    BackStyle =1
                    OverlapFlags =93
                    Left =8580
                    Width =1008
                    Height =420
                    BackColor =13434828
                    Name ="rct6data"
                    LayoutCachedLeft =8580
                    LayoutCachedWidth =9588
                    LayoutCachedHeight =420
                End
                Begin Rectangle
                    BackStyle =1
                    OverlapFlags =93
                    Left =6300
                    Width =1008
                    Height =420
                    BackColor =13434828
                    Name ="rct4data"
                    LayoutCachedLeft =6300
                    LayoutCachedWidth =7308
                    LayoutCachedHeight =420
                End
                Begin Rectangle
                    BackStyle =1
                    OverlapFlags =93
                    Left =4200
                    Width =1008
                    Height =420
                    BackColor =13434828
                    Name ="rct2data"
                    LayoutCachedLeft =4200
                    LayoutCachedWidth =5208
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =360
                    Height =255
                    ColumnWidth =2310
                    Name ="Shrub_ID"
                    ControlSource ="Shrub_ID"
                    StatusBarText ="Unique record identifier - primary key"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =300
                    Top =60
                    Width =300
                    Height =255
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Transect_ID"
                    ControlSource ="Transect_ID"
                    StatusBarText ="Foreign key to tbl_Canopy_Transect"

                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3360
                    Top =60
                    Width =540
                    Height =255
                    ColumnWidth =600
                    TabIndex =4
                    BackColor =65535
                    Name ="HC10"
                    ControlSource ="HC10"
                    StatusBarText ="0-10cm height class total"
                    ConditionalFormat = Begin
                        0x01000000b6000000030000000000000006000000000000000200000001000000 ,
                        0x00000000ffffff00000000000500000003000000050000000100000000000000 ,
                        0xffff00000100000000000000060000002a0000000100000000000000ffffff00 ,
                        0x3000000000003000000000005b0050006100720065006e0074005d002e005b00 ,
                        0x6300620078004e006f005300680072007500620073005d002e005b0056006100 ,
                        0x6c00750065005d003d00540072007500650000000000
                    End

                    LayoutCachedLeft =3360
                    LayoutCachedTop =60
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =315
                    ConditionalFormat14 = Begin
                        0x01000400000000000000060000000100000000000000ffffff00010000003000 ,
                        0x0000000000000000000000000000000000000000000000000005000000010000 ,
                        0x0000000000ffff00000100000030000000000000000000000000000000000000 ,
                        0x0000000001000000000000000100000000000000ffffff00230000005b005000 ,
                        0x6100720065006e0074005d002e005b006300620078004e006f00530068007200 ,
                        0x7500620073005d002e005b00560061006c00750065005d003d00540072007500 ,
                        0x6500000000000000000000000000000000000000000000010000000000000001 ,
                        0x00000000000000ffff66002d0000005b0048004300310030005d002b005b0048 ,
                        0x004300320035005d002b005b0048004300350030005d002b005b004800430031 ,
                        0x00300030005d002b005b004800430032006d005d002b005b0048004300470054 ,
                        0x0032005d003d003000000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4380
                    Top =60
                    Width =540
                    Height =255
                    ColumnWidth =600
                    TabIndex =5
                    BackColor =65535
                    Name ="HC25"
                    ControlSource ="HC25"
                    StatusBarText ="10.1-25cm height class total"
                    ConditionalFormat = Begin
                        0x01000000b6000000030000000000000006000000000000000200000001000000 ,
                        0x00000000ffffff00000000000500000003000000050000000100000000000000 ,
                        0xffff00000100000000000000060000002a0000000100000000000000ffffff00 ,
                        0x3000000000003000000000005b0050006100720065006e0074005d002e005b00 ,
                        0x6300620078004e006f005300680072007500620073005d002e005b0056006100 ,
                        0x6c00750065005d003d00540072007500650000000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000400000000000000060000000100000000000000ffffff00010000003000 ,
                        0x0000000000000000000000000000000000000000000000000005000000010000 ,
                        0x0000000000ffff00000100000030000000000000000000000000000000000000 ,
                        0x0000000001000000000000000100000000000000ffffff00230000005b005000 ,
                        0x6100720065006e0074005d002e005b006300620078004e006f00530068007200 ,
                        0x7500620073005d002e005b00560061006c00750065005d003d00540072007500 ,
                        0x6500000000000000000000000000000000000000000000010000000000000001 ,
                        0x00000000000000ffff66002d0000005b0048004300310030005d002b005b0048 ,
                        0x004300320035005d002b005b0048004300350030005d002b005b004800430031 ,
                        0x00300030005d002b005b004800430032006d005d002b005b0048004300470054 ,
                        0x0032005d003d003000000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5520
                    Top =60
                    Width =540
                    Height =255
                    ColumnWidth =600
                    TabIndex =6
                    BackColor =65535
                    Name ="HC50"
                    ControlSource ="HC50"
                    StatusBarText ="25.1-50cm height class total"
                    ConditionalFormat = Begin
                        0x01000000b6000000030000000000000006000000000000000200000001000000 ,
                        0x00000000ffffff00000000000500000003000000050000000100000000000000 ,
                        0xffff00000100000000000000060000002a0000000100000000000000ffffff00 ,
                        0x3000000000003000000000005b0050006100720065006e0074005d002e005b00 ,
                        0x6300620078004e006f005300680072007500620073005d002e005b0056006100 ,
                        0x6c00750065005d003d00540072007500650000000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000400000000000000060000000100000000000000ffffff00010000003000 ,
                        0x0000000000000000000000000000000000000000000000000005000000010000 ,
                        0x0000000000ffff00000100000030000000000000000000000000000000000000 ,
                        0x0000000001000000000000000100000000000000ffffff00230000005b005000 ,
                        0x6100720065006e0074005d002e005b006300620078004e006f00530068007200 ,
                        0x7500620073005d002e005b00560061006c00750065005d003d00540072007500 ,
                        0x6500000000000000000000000000000000000000000000010000000000000001 ,
                        0x00000000000000ffff66002d0000005b0048004300310030005d002b005b0048 ,
                        0x004300320035005d002b005b0048004300350030005d002b005b004800430031 ,
                        0x00300030005d002b005b004800430032006d005d002b005b0048004300470054 ,
                        0x0032005d003d003000000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6540
                    Top =60
                    Width =540
                    Height =255
                    ColumnWidth =600
                    TabIndex =7
                    BackColor =65535
                    Name ="HC100"
                    ControlSource ="HC100"
                    StatusBarText ="50.1-100cm height class total"
                    ConditionalFormat = Begin
                        0x01000000b6000000030000000000000006000000000000000200000001000000 ,
                        0x00000000ffffff00000000000500000003000000050000000100000000000000 ,
                        0xffff00000100000000000000060000002a0000000100000000000000ffffff00 ,
                        0x3000000000003000000000005b0050006100720065006e0074005d002e005b00 ,
                        0x6300620078004e006f005300680072007500620073005d002e005b0056006100 ,
                        0x6c00750065005d003d00540072007500650000000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000400000000000000060000000100000000000000ffffff00010000003000 ,
                        0x0000000000000000000000000000000000000000000000000005000000010000 ,
                        0x0000000000ffff00000100000030000000000000000000000000000000000000 ,
                        0x0000000001000000000000000100000000000000ffffff00230000005b005000 ,
                        0x6100720065006e0074005d002e005b006300620078004e006f00530068007200 ,
                        0x7500620073005d002e005b00560061006c00750065005d003d00540072007500 ,
                        0x6500000000000000000000000000000000000000000000010000000000000001 ,
                        0x00000000000000ffff66002d0000005b0048004300310030005d002b005b0048 ,
                        0x004300320035005d002b005b0048004300350030005d002b005b004800430031 ,
                        0x00300030005d002b005b004800430032006d005d002b005b0048004300470054 ,
                        0x0032005d003d003000000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7680
                    Top =60
                    Width =540
                    Height =255
                    ColumnWidth =600
                    TabIndex =8
                    BackColor =65535
                    Name ="HC2m"
                    ControlSource ="HC2m"
                    StatusBarText ="1.01-2m height class total"
                    ConditionalFormat = Begin
                        0x01000000b6000000030000000000000006000000000000000200000001000000 ,
                        0x00000000ffffff00000000000500000003000000050000000100000000000000 ,
                        0xffff00000100000000000000060000002a0000000100000000000000ffffff00 ,
                        0x3000000000003000000000005b0050006100720065006e0074005d002e005b00 ,
                        0x6300620078004e006f005300680072007500620073005d002e005b0056006100 ,
                        0x6c00750065005d003d00540072007500650000000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000400000000000000060000000100000000000000ffffff00010000003000 ,
                        0x0000000000000000000000000000000000000000000000000005000000010000 ,
                        0x0000000000ffff00000100000030000000000000000000000000000000000000 ,
                        0x0000000001000000000000000100000000000000ffffff00230000005b005000 ,
                        0x6100720065006e0074005d002e005b006300620078004e006f00530068007200 ,
                        0x7500620073005d002e005b00560061006c00750065005d003d00540072007500 ,
                        0x6500000000000000000000000000000000000000000000010000000000000001 ,
                        0x00000000000000ffff66002d0000005b0048004300310030005d002b005b0048 ,
                        0x004300320035005d002b005b0048004300350030005d002b005b004800430031 ,
                        0x00300030005d002b005b004800430032006d005d002b005b0048004300470054 ,
                        0x0032005d003d003000000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8760
                    Top =60
                    Width =540
                    Height =255
                    ColumnWidth =600
                    TabIndex =9
                    BackColor =65535
                    Name ="HCGT2"
                    ControlSource ="HCGT2"
                    StatusBarText =">2.01m height class total"
                    ConditionalFormat = Begin
                        0x01000000b6000000030000000000000006000000000000000200000001000000 ,
                        0x00000000ffffff00000000000500000003000000050000000100000000000000 ,
                        0xffff00000100000000000000060000002a0000000100000000000000ffffff00 ,
                        0x3000000000003000000000005b0050006100720065006e0074005d002e005b00 ,
                        0x6300620078004e006f005300680072007500620073005d002e005b0056006100 ,
                        0x6c00750065005d003d00540072007500650000000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000400000000000000060000000100000000000000ffffff00010000003000 ,
                        0x0000000000000000000000000000000000000000000000000005000000010000 ,
                        0x0000000000ffff00000100000030000000000000000000000000000000000000 ,
                        0x0000000001000000000000000100000000000000ffffff00230000005b005000 ,
                        0x6100720065006e0074005d002e005b006300620078004e006f00530068007200 ,
                        0x7500620073005d002e005b00560061006c00750065005d003d00540072007500 ,
                        0x6500000000000000000000000000000000000000000000010000000000000001 ,
                        0x00000000000000ffff66002d0000005b0048004300310030005d002b005b0048 ,
                        0x004300320035005d002b005b0048004300350030005d002b005b004800430031 ,
                        0x00300030005d002b005b004800430032006d005d002b005b0048004300470054 ,
                        0x0032005d003d003000000000000000000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =375
                    Left =2520
                    Top =60
                    Width =780
                    TabIndex =3
                    Name ="Alive"
                    ControlSource ="Alive"
                    RowSourceType ="Value List"
                    RowSource ="-1;\"Yes\";0;\"No\""
                    ColumnWidths ="0;375"
                    DefaultValue ="-1"

                End
                Begin ComboBox
                    OverlapFlags =247
                    TextFontFamily =0
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =6480
                    Left =120
                    Top =60
                    Width =2304
                    TabIndex =2
                    BackColor =65535
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    ConditionalFormat = Begin
                        0x0100000046010000030000000100000000000000000000001100000001000000 ,
                        0x00000000ffffff000100000000000000120000004f0000000100000000000000 ,
                        0xffff0000010000000000000050000000720000000100000000000000ffffff00 ,
                        0x4c0065006e0028005b0053007000650063006900650073005d0029003e003000 ,
                        0x000000004900490066002800490073004e0075006c006c0028005b0048004300 ,
                        0x310030005d002b005b0048004300320035005d002b005b004800430035003000 ,
                        0x5d002b005b00480043003100300030005d002b005b004800430032006d005d00 ,
                        0x2b005b00480043004700540032005d0029002c0030002c003100290000000000 ,
                        0x5b0050006100720065006e0074005d002e005b006300620078004e006f005300 ,
                        0x680072007500620073005d002e00560061006c00750065003d00540072007500 ,
                        0x650000000000
                    End
                    Name ="Species"
                    ControlSource ="Species"
                    RowSourceType ="Table/Query"
                    RowSource ="(SELECT DISTINCT qryU_Top_Canopy.Master_PLANT_Code, qryU_Top_Canopy.LU_Code, qry"
                        "U_Top_Canopy.Utah_Species, qryU_Top_Canopy.Lifeform FROM qryU_Top_Canopy WHERE ("
                        "((qryU_Top_Canopy.Utah_Species) Is Not Null) AND ((qryU_Top_Canopy.Lifeform) In "
                        "('Shrub','DwarfShrub')) AND ((qryU_Top_Canopy.tlu_NCPN_Plants.Master_PLANT_Code)"
                        " Not In (SELECT Master_PLANT_Code FROM ShrubExclusionList))))    UNION  (SELECT "
                        "DISTINCT tbl_Unknown_Species.Unknown_Code,  tbl_Unknown_Species.Unknown_Code,   "
                        " tbl_Unknown_Species.Plant_Type + \" - \" + tbl_Unknown_Species.Plant_Descriptio"
                        "n,  tbl_Unknown_Species.Plant_Type AS Lifeform  FROM tbl_Unknown_Species  WHERE "
                        "tbl_Unknown_Species.Plant_Type IN ('Shrub','Other') OR tbl_Unknown_Species.Plant"
                        "_Type IS NULL ORDER BY tbl_Unknown_Species.Unknown_Code)  ORDER BY qryU_Top_Cano"
                        "py.LU_Code;"
                    ColumnWidths ="0;2160;4320"
                    BeforeUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"
                    ConditionalFormat14 = Begin
                        0x01000300000001000000000000000100000000000000ffffff00100000004c00 ,
                        0x65006e0028005b0053007000650063006900650073005d0029003e0030000000 ,
                        0x0000000000000000000000000000000000000001000000000000000100000000 ,
                        0x000000ffff00003c0000004900490066002800490073004e0075006c006c0028 ,
                        0x005b0048004300310030005d002b005b0048004300320035005d002b005b0048 ,
                        0x004300350030005d002b005b00480043003100300030005d002b005b00480043 ,
                        0x0032006d005d002b005b00480043004700540032005d0029002c0030002c0031 ,
                        0x0029000000000000000000000000000000000000000000000100000000000000 ,
                        0x0100000000000000ffffff00210000005b0050006100720065006e0074005d00 ,
                        0x2e005b006300620078004e006f005300680072007500620073005d002e005600 ,
                        0x61006c00750065003d0054007200750065000000000000000000000000000000 ,
                        0x00000000000000
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =10080
                    Top =60
                    Width =1275
                    Height =300
                    TabIndex =10
                    ForeColor =255
                    Name ="ButtonDelete"
                    Caption ="Delete Record"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
        Begin FormFooter
            Height =420
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =4560
                    Top =60
                    Width =606
                    Height =288
                    Name ="ButtonA1"
                    Caption ="+ 1"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5280
                    Top =60
                    Width =606
                    Height =288
                    TabIndex =1
                    Name ="ButtonA5"
                    Caption ="+ 5"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6000
                    Top =60
                    Width =606
                    Height =288
                    TabIndex =2
                    Name ="ButtonS1"
                    Caption ="- 1"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6720
                    Top =60
                    Width =606
                    Height =288
                    TabIndex =3
                    Name ="ButtonS5"
                    Caption ="- 5"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7440
                    Top =60
                    Width =606
                    Height =288
                    TabIndex =4
                    Name ="ButtonZero"
                    Caption ="0"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
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
' MODULE:       Form_fsub_LP_Belt_Shrub
' Level:        Form module
' Version:      1.04
' Description:  data functions & procedures specific to LP belt shrub monitoring
'
' Source/date:  Bonnie Campbell, 2/2/2016
' Revisions:    RDB - unknown  - 1.00 - initial version
'               BLC - 2/2/2016 - 1.01 - added documentation, checkbox for no species found
'               BLC - 3/7/2016 - 1.02 - fixed Error #94 Invalid use of NULL issue in Form_Load() on new visits
'               BLC - 4/13/2016 - 1.03 - adjusted species dropdown SQL to include JUNOST & PINEDU (LU_Code)
'                                        tree species & changed the sort to be across the entire
'                                        UNION query versus within its component SELECTS
'                                        to accommodate "EPHVIR" being selected vs. "EPHVIRV"
'                                        when users type in "EPHVIR" (see also 2016 Upland In-Season Updates documentation)
'                                        added refresh for underlying subforms for conditional formatting
'               BLC - 11/29/2016 - 1.04 - adjusted species dropdown SQL to *EXCLUDE* JUNOST & PINEDU (LU_Code)
'                                        tree species w/ the decision to no longer treat these species as
'                                        shrubs (see also 2017 Upland Pre-Season Updates documentation)
' =================================

' ---------------------------------
' SUB:          Form_Load
' Description:  Handles form loading actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 2, 2016 - for NCPN tools
' Revisions:
'   BLC, 2/2/2016  - initial version
'   BLC, 3/7/2016  - addressed Me.Transect_ID = NULL issue causing Error #94 Invalid use of NULL
'                    when new visit is created (IsNull check)
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler

    Dim NoData As Scripting.Dictionary
    
    If Not IsNull(Me.Transect_ID) Then
        'set no data checkboxes/rectangles
        Set NoData = GetNoDataCollected(Me.Transect_ID, "T")
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[Form_fsub_LP_Belt_Shrub])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_BeforeInsert
' Description:  Handles form pre-insert actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Russ DenBleyker, unknown
' Adapted:      Bonnie Campbell, February 11, 2016 - for NCPN tools
' Revisions:
'   RDB, unknown    - initial version
'   BLC, 2/11/2016  - added no data collected info updates
' ---------------------------------
Private Sub Form_BeforeInsert(Cancel As Integer)
On Error GoTo Err_Handler

  '  If IsNull(Me.Parent!Observer) And IsNull(Me.Parent!Recorder) Then
  '    MsgBox "You must enter an observer or recorder first."
  '    DoCmd.CancelEvent
  '    SendKeys "{ESC}"
  '    GoTo Exit_Procedure
  '  End If
  
    ' Create the GUID primary key value
    If IsNull(Me!Shrub_ID) Then
        If GetDataType("tbl_LP_Shrub", "Shrub_ID") = dbText Then
            Me.Shrub_ID = fxnGUIDGen
        End If
    End If
    
    '-----------------------------------
    ' update the NoDataCollected info
    '-----------------------------------
    Dim NoData As Scripting.Dictionary
    
    'remove the no data collected record
    Set NoData = SetNoDataCollected(Me.Parent!Transect_ID, "T", "1mBelt-Shrub", 0)
        
    'update checkbox/rectangle
    Me.Parent.Form.Controls("cbxNoShrubs") = 0
    Me.Parent.Form.Controls("cbxNoShrubs").Enabled = False
    Me.Parent.Form.Controls("rctNoShrubs").Visible = False

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeInsert[Form_fsub_LP_Belt_Shrub])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Species_GotFocus
' Description:  Handles species actions when control has focus
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Russ DenBleyker, unknown
' Adapted:      Bonnie Campbell, February 9, 2016 - for NCPN tools
' Revisions:
'   RDB, unknown  - initial version
'   BLC, 2/9/2016 - added error handling, documentation, refresh list to catch unknowns
' ---------------------------------
Private Sub Species_GotFocus()
On Error GoTo Err_Handler

    If IsNull(Me.Parent!Visit_Date) Then    ' If they didn't bother to enter a date, default to event date.
      Me.Parent!Visit_Date = Me.Parent.Parent!Start_Date
      Me.Parent.Refresh   ' Force save of transect record
    End If

    'update the data to ensure new unknowns are added
    Me.ActiveControl.Requery

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Species_GotFocus[Form_fsub_LP_Belt_Shrub])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Species_BeforeUpdate
' Description:  Handles species pre-update actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 2, 2016 - for NCPN tools
' Revisions:
'   RDB, unknown  - initial version
'   BLC, 2/2/2016  - added documentation, disable checkbox if species exist
' ---------------------------------
Private Sub Species_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

    If Not IsNull(DLookup("[Shrub_ID]", "tbl_LP_Shrub", "[Transect_ID] = '" & Me!Transect_ID & "' AND [Species] = '" & Me!Species & "'")) Then
      MsgBox "This species is already recorded for this transect."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
    End If
    
    'if species is added disable checkbox & change color of rectangle background
    If Not IsNull(Me.Species) Then
        Me.Parent.Form.Controls("cbxNoShrubs").Enabled = False
        Me.Parent.Form.Controls("rctNoShrubs").Visible = False
    End If
    
    'capture the CTRL+Z keystroke
    
    'note: watch out for SendKeys   http://access.mvps.org/access/api/api0046.htm
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Species_BeforeUpdate[Form_fsub_LP_Belt_Shrub])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub Button_Master_Species_Click()
On Error GoTo Err_Button_Master_Species_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim strOpenArg As String

    strOpenArg = "fsub_LP_Belt_Shrub"
    stDocName = "frm_Master_Species"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , , strOpenArg

Exit_Button_Master_Species_Click:
    Exit Sub

Err_Button_Master_Species_Click:
    MsgBox Err.Description
    Resume Exit_Button_Master_Species_Click
 
End Sub

Private Sub ButtonUnknown_Click()
    Dim stDocName As String
    Dim stLinkCriteria As String

    DoCmd.OpenForm "frm_List_Unknown", , , , , acDialog
    Me.Refresh
    
 '   stLinkCriteria = "[Species_ID]=" & "'" & Me![Shrub_ID] & "'"
 '   DoCmd.OpenForm stDocName, , , stLinkCriteria, , , Me![Shrub_ID]
End Sub

Private Sub ButtonA1_Click()

  If Screen.PreviousControl.Name <> "Species" And Not IsNull(Me!Species) Then
    If IsNull(Screen.PreviousControl.value) Then
      Screen.PreviousControl.value = 1
    Else
      Screen.PreviousControl.value = Screen.PreviousControl.value + 1
    End If
  End If
  Screen.PreviousControl.SetFocus
End Sub

Private Sub ButtonA5_Click()
  If Screen.PreviousControl.Name <> "Species" And Not IsNull(Me!Species) Then
    If IsNull(Screen.PreviousControl.value) Then
      Screen.PreviousControl.value = 5
    Else
      Screen.PreviousControl.value = Screen.PreviousControl.value + 5
    End If
  End If
  Screen.PreviousControl.SetFocus
End Sub

Private Sub ButtonS1_Click()
  If Screen.PreviousControl.Name <> "Species" And Not IsNull(Me!Species) Then
    If IsNull(Screen.PreviousControl.value) Then
      Screen.PreviousControl.value = 0
    ElseIf Screen.PreviousControl.value - 1 < 0 Then
      MsgBox "Total cannot be negative.", , "Belt Shrubs"
      Exit Sub
    Else
      Screen.PreviousControl.value = Screen.PreviousControl.value - 1
    End If
  End If
  Screen.PreviousControl.SetFocus
End Sub

Private Sub ButtonS5_Click()
  If Screen.PreviousControl.Name <> "Species" And Not IsNull(Me!Species) Then
    If IsNull(Screen.PreviousControl.value) Then
      Screen.PreviousControl.value = 0
    ElseIf Screen.PreviousControl.value - 5 < 0 Then
      MsgBox "Total cannot be negative.", , "Belt Shrubs"
      Exit Sub
    Else
      Screen.PreviousControl.value = Screen.PreviousControl.value - 5
    End If
  End If
  Screen.PreviousControl.SetFocus
End Sub

Private Sub ButtonZero_Click()
  If Screen.PreviousControl.Name <> "Species" And Not IsNull(Me!Species) Then
      Screen.PreviousControl.value = 0
  End If
  Screen.PreviousControl.SetFocus
End Sub

' ---------------------------------
' SUB:          ButtonDelete_Click
' Description:  Handles delete button actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Russ DenBleyker, unknown
' Adapted:      Bonnie Campbell, February 11, 2016 - for NCPN tools
' Revisions:
'   RDB, unknown  - initial version
'   BLC, 2/11/2016 - added error handling, documentation, refresh checkbox/no data collected
' ---------------------------------
Private Sub ButtonDelete_Click()
On Error GoTo Err_Handler
  Dim intReply As Integer
  
  intReply = MsgBox("Are you sure you want to delete this record?", vbYesNo, "Delete Record")
    If intReply = vbYes Then
      DoCmd.SetWarnings False
      DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
      DoCmd.DoMenuItem acFormBar, acEditMenu, 6, , acMenuVer70
      DoCmd.SetWarnings True
      Me.Requery
    End If

    '-----------------------------------
    ' update the NoDataCollected info IF no records now exist
    '-----------------------------------
    If Me.RecordsetClone.RecordCount = 0 Then
    
        Dim NoData As Scripting.Dictionary
        
        'remove the no data collected record
        Set NoData = SetNoDataCollected(Me.Parent.Form.Controls("Transect_ID"), "T", "1mBelt-Shrub", 1)
    
        'update checkbox/rectangle
        Me.Parent.Form.Controls("cbxNoShrubs") = 1
        Me.Parent.Form.Controls("cbxNoShrubs").Enabled = True
        Me.Parent.Form.Controls("rctNoShrubs").Visible = True
        
        'refresh the form for conditional formatting
        Me.Requery
        
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ButtonDelete_Click[Form_fsub_LP_Belt_Shrub])"
    End Select
    Resume Exit_Handler
End Sub
