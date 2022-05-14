Version =21
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    OrderByOn = NotDefault
    DefaultView =0
    TabularFamily =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9300
    DatasheetFontHeight =9
    ItemSuffix =44
    Left =1950
    Top =60
    Right =14310
    Bottom =9720
    DatasheetGridlinesColor =12632256
    Filter ="Unit_Code = 'DINO' AND Plot_ID = 116"
    OrderBy ="[temp_Sp_Rpt_by_Park_Rollup].[ParkPlotSpecies]"
    RecSrcDt = Begin
        0xfe3370dd0da1e440
    End
    RecordSource ="temp_Sp_Rpt_by_Park_Rollup"
    Caption ="rpt_Species_by_Park"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000542400004a01000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            TextAlign =1
            TextFontFamily =18
            FontSize =9
            FontWeight =700
            FontName ="Times New Roman"
        End
        Begin Rectangle
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Line
            BorderLineStyle =0
        End
        Begin Image
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =18
            BorderLineStyle =0
            BackStyle =0
            FontName ="Times New Roman"
            AsianLineBreak =255
        End
        Begin ListBox
            TextFontFamily =18
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Times New Roman"
        End
        Begin ComboBox
            OldBorderStyle =0
            TextFontFamily =18
            BorderLineStyle =0
            BackStyle =0
            FontName ="Times New Roman"
        End
        Begin Subform
            OldBorderStyle =0
            BorderLineStyle =0
        End
        Begin BreakLevel
            KeepTogether =2
            ControlSource ="Unit_Code"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            KeepTogether =2
            ControlSource ="Plot_ID"
        End
        Begin BreakLevel
            ControlSource ="Plot_ID"
        End
        Begin BreakLevel
            ControlSource ="Master_Family"
        End
        Begin BreakLevel
            ControlSource ="Utah_Species"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =495
            BackColor =11525325
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =1080
                    Width =2685
                    Height =480
                    FontSize =18
                    FontWeight =400
                    BackColor =11525325
                    Name ="Label10"
                    Caption ="Species by Park:"
                    FontName ="Calibri"
                    LayoutCachedLeft =1080
                    LayoutCachedWidth =3765
                    LayoutCachedHeight =480
                    ThemeFontIndex =1
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =5640
                    Height =480
                    FontSize =18
                    Name ="Text27"
                    ControlSource ="Unit_Code"
                    FontName ="Calibri"

                    LayoutCachedLeft =5640
                    LayoutCachedWidth =7080
                    LayoutCachedHeight =480
                End
                Begin TextBox
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =7080
                    Width =840
                    Height =480
                    FontSize =18
                    TabIndex =1
                    Name ="Text29"
                    ControlSource ="Plot_ID"
                    FontName ="Calibri"

                    LayoutCachedLeft =7080
                    LayoutCachedWidth =7920
                    LayoutCachedHeight =480
                End
                Begin Line
                    BorderWidth =2
                    Top =480
                    Width =9300
                    Name ="Line_Report_Banner_Bottom"
                    LayoutCachedTop =480
                    LayoutCachedWidth =9300
                    LayoutCachedHeight =480
                End
                Begin Line
                    BorderWidth =2
                    Width =9300
                    Name ="Line_Report_Banner_Top"
                    LayoutCachedWidth =9300
                End
            End
        End
        Begin PageHeader
            Height =915
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextFontFamily =34
                    Left =3480
                    Top =600
                    Width =960
                    Height =270
                    FontSize =10
                    Name ="Utah_Species_Label"
                    Caption ="Species"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =3480
                    LayoutCachedTop =600
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =870
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =7500
                    Top =600
                    Width =600
                    Height =270
                    FontSize =10
                    Name ="lblYears"
                    Caption ="Years"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =7500
                    LayoutCachedTop =600
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =870
                End
                Begin Label
                    TextFontFamily =34
                    Left =660
                    Top =600
                    Width =840
                    Height =270
                    FontSize =10
                    Name ="Master_Family_Label"
                    Caption ="Family"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =660
                    LayoutCachedTop =600
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =870
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =1
                    IMESentenceMode =3
                    Width =5220
                    Height =360
                    FontSize =14
                    FontWeight =500
                    BackColor =11525325
                    ForeColor =4210752
                    Name ="tbxPageHeader"
                    ControlSource ="=IIf([Page]>1,\"Species by Park\",\"\")"
                    FontName ="Calibri"

                    LayoutCachedWidth =5220
                    LayoutCachedHeight =360
                    ThemeFontIndex =1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =900
                    Width =9240
                    Name ="Line_Below_Column_Headers"
                    LayoutCachedLeft =60
                    LayoutCachedTop =900
                    LayoutCachedWidth =9300
                    LayoutCachedHeight =900
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4980
                    Top =600
                    Width =2160
                    Height =270
                    FontSize =9
                    FontWeight =500
                    TabIndex =1
                    Name ="tbxFilter"
                    ControlSource ="=IIf(Len([OpenArgs])>0,\"Filter:  \" & [OpenArgs],\"\")"
                    FontName ="Calibri"

                    LayoutCachedLeft =4980
                    LayoutCachedTop =600
                    LayoutCachedWidth =7140
                    LayoutCachedHeight =870
                    ThemeFontIndex =1
                End
                Begin TextBox
                    HideDuplicates = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =1
                    IMESentenceMode =3
                    Left =5220
                    Width =1620
                    Height =360
                    FontSize =14
                    TabIndex =2
                    BackColor =11525325
                    ForeColor =4210752
                    Name ="Text25"
                    ControlSource ="=IIf([Page]>1,[Unit_Code],\"\")"
                    FontName ="Calibri"

                    LayoutCachedLeft =5220
                    LayoutCachedWidth =6840
                    LayoutCachedHeight =360
                    ThemeFontIndex =1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =1
                    IMESentenceMode =3
                    Left =6840
                    Width =1620
                    Height =360
                    FontSize =14
                    TabIndex =3
                    BackColor =11525325
                    ForeColor =4210752
                    Name ="Text26"
                    ControlSource ="=IIf([Page]>1,[Plot_ID],\"\")"
                    FontName ="Calibri"

                    LayoutCachedLeft =6840
                    LayoutCachedWidth =8460
                    LayoutCachedHeight =360
                    ThemeFontIndex =1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                End
                Begin TextBox
                    CanShrink = NotDefault
                    BackStyle =1
                    IMESentenceMode =3
                    Left =8460
                    Width =840
                    Height =120
                    FontSize =9
                    TabIndex =4
                    BackColor =11525325
                    ForeColor =11525325
                    Name ="Text37"
                    ControlSource ="=IIf([Page]>1,[Pages],\"\")"

                    LayoutCachedLeft =8460
                    LayoutCachedWidth =9300
                    LayoutCachedHeight =120
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =1
                    IMESentenceMode =3
                    Left =8460
                    Top =120
                    Width =840
                    FontSize =9
                    TabIndex =5
                    BackColor =11525325
                    ForeColor =4210752
                    Name ="Text36"
                    ControlSource ="=IIf([Page]>1,[Page] & \" of \" & [Pages],\"\")"
                    FontName ="Calibri"

                    LayoutCachedLeft =8460
                    LayoutCachedTop =120
                    LayoutCachedWidth =9300
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                End
                Begin Line
                    BorderWidth =1
                    Top =360
                    Width =9300
                    Name ="Line40"
                    LayoutCachedTop =360
                    LayoutCachedWidth =9300
                    LayoutCachedHeight =360
                End
                Begin Line
                    BorderWidth =1
                    Width =9300
                    Name ="Line_Page_Banner_Top"
                    LayoutCachedWidth =9300
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            RepeatSection = NotDefault
            Height =300
            BreakLevel =1
            BackColor =11525325
            Name ="GroupHeader1"
            AlternateBackColor =8965045
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =2880
                    Width =3600
                    Height =270
                    FontSize =10
                    FontWeight =500
                    Name ="tbxNoData"
                    ControlSource ="=IIf(IsNull([SpeciesYears]) Or IsNull([Plot_ID]),\"- No Data Found -\",\"\")"
                    FontName ="Calibri"

                    LayoutCachedLeft =2880
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =270
                    ThemeFontIndex =1
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    BackStyle =1
                    IMESentenceMode =3
                    Left =1020
                    TabIndex =1
                    BackColor =11525325
                    ForeColor =11525325
                    Name ="Plot_ID"
                    ControlSource ="Plot_ID"

                    LayoutCachedLeft =1020
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =240
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =330
            Name ="Detail"
            AlternateBackColor =12566463
            Begin
                Begin TextBox
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =2940
                    Top =60
                    Width =2940
                    Height =270
                    ColumnWidth =2520
                    Name ="Utah_Species"
                    ControlSource ="Utah_Species"
                    FontName ="Calibri"

                    LayoutCachedLeft =2940
                    LayoutCachedTop =60
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =5940
                    Top =60
                    Width =3312
                    Height =270
                    TabIndex =1
                    Name ="tbxYear"
                    ControlSource ="SpeciesYears"
                    FontName ="Calibri"

                    LayoutCachedLeft =5940
                    LayoutCachedTop =60
                    LayoutCachedWidth =9252
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =600
                    Top =60
                    Width =2280
                    Height =270
                    ColumnWidth =1395
                    TabIndex =2
                    Name ="Master_Family"
                    ControlSource ="Master_Family"
                    FontName ="Calibri"

                    LayoutCachedLeft =600
                    LayoutCachedTop =60
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =330
                End
            End
        End
        Begin PageFooter
            Height =390
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Top =120
                    Width =4560
                    Height =270
                    Name ="Text11"
                    ControlSource ="=Now()"
                    Format ="Long Date"

                    LayoutCachedLeft =60
                    LayoutCachedTop =120
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4740
                    Top =120
                    Width =4560
                    Height =270
                    TabIndex =1
                    Name ="Text12"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"

                    LayoutCachedLeft =4740
                    LayoutCachedTop =120
                    LayoutCachedWidth =9300
                    LayoutCachedHeight =390
                End
                Begin Line
                    Left =60
                    Top =120
                    Width =9240
                    Name ="Line_Bottom_of_Page"
                    LayoutCachedLeft =60
                    LayoutCachedTop =120
                    LayoutCachedWidth =9300
                    LayoutCachedHeight =120
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
        End
    End
End
