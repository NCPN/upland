Version =21
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PageHeader =1
    TabularFamily =126
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10830
    DatasheetFontHeight =9
    ItemSuffix =65
    Left =2610
    Top =2280
    Right =14970
    Bottom =11985
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x04ae7a726ca8e340
    End
    RecordSource ="qry_sel_OT_Census_Report"
    Caption ="rpt_OT_Census"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x40020000d002000040020000d0020000000000004e2a00005802000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =255
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
            GroupHeader = NotDefault
            ControlSource ="Unit_Code"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="Plot_ID"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="Quad"
        End
        Begin BreakLevel
            ControlSource ="Tag_No"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =480
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =1740
                    Width =7380
                    Height =420
                    FontSize =18
                    FontWeight =400
                    Name ="Label20"
                    Caption ="Overstory Census Revisit"
                    LayoutCachedLeft =1740
                    LayoutCachedWidth =9120
                    LayoutCachedHeight =420
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =420
                    Width =10617
                    Name ="Line34"
                    LayoutCachedLeft =60
                    LayoutCachedTop =420
                    LayoutCachedWidth =10677
                    LayoutCachedHeight =420
                End
            End
        End
        Begin PageHeader
            Height =375
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =2
                    Left =60
                    Width =540
                    Height =270
                    FontSize =10
                    Name ="Label48"
                    Caption ="Tag #"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =600
                    LayoutCachedHeight =270
                End
                Begin Label
                    TextAlign =2
                    Left =660
                    Width =1620
                    Height =270
                    FontSize =10
                    Name ="Label49"
                    Caption ="Species"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =660
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =270
                End
                Begin Label
                    TextAlign =2
                    Left =2340
                    Width =1005
                    Height =270
                    FontSize =10
                    Name ="Label50"
                    Caption ="DBH/DRC"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =2340
                    LayoutCachedWidth =3345
                    LayoutCachedHeight =270
                End
                Begin Label
                    TextAlign =2
                    Left =3420
                    Width =1155
                    Height =270
                    FontSize =10
                    Name ="Label51"
                    Caption ="Crown Class"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =3420
                    LayoutCachedWidth =4575
                    LayoutCachedHeight =270
                End
                Begin Label
                    TextAlign =2
                    Left =4620
                    Width =1320
                    Height =270
                    FontSize =10
                    Name ="Label52"
                    Caption ="Crown Health"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =4620
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =270
                End
                Begin Label
                    TextAlign =2
                    Left =6360
                    Width =2700
                    Height =270
                    FontSize =10
                    Name ="Label53"
                    Caption ="Notes"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =6360
                    LayoutCachedWidth =9060
                    LayoutCachedHeight =270
                End
                Begin Line
                    Top =300
                    Width =10740
                    Name ="Line56"
                    LayoutCachedTop =300
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9420
                    Width =660
                    ForeColor =3422101
                    Name ="Text57"
                    ControlSource ="Unit_Code"
                    StatusBarText ="Park Code."

                    LayoutCachedLeft =9420
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10080
                    Width =480
                    TabIndex =1
                    ForeColor =3422101
                    Name ="Text61"
                    ControlSource ="Plot_ID"
                    StatusBarText ="Plot identifier"

                    LayoutCachedLeft =10080
                    LayoutCachedWidth =10560
                    LayoutCachedHeight =240
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =0
            Name ="GroupHeader0"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =480
            BreakLevel =1
            Name ="GroupHeader1"
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5640
                    Top =120
                    Width =1080
                    Height =360
                    FontSize =12
                    FontWeight =700
                    ForeColor =3422101
                    Name ="Plot_ID"
                    ControlSource ="Plot_ID"
                    StatusBarText ="Plot identifier"

                    LayoutCachedLeft =5640
                    LayoutCachedTop =120
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =480
                    Begin
                        Begin Label
                            Left =4680
                            Top =120
                            Width =660
                            Height =360
                            FontSize =12
                            ForeColor =3422101
                            Name ="Plot_ID_Label"
                            Caption ="Plot:"
                            LayoutCachedLeft =4680
                            LayoutCachedTop =120
                            LayoutCachedWidth =5340
                            LayoutCachedHeight =480
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1500
                    Top =120
                    Height =360
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    ForeColor =3422101
                    Name ="Unit_Code"
                    ControlSource ="Unit_Code"
                    StatusBarText ="Park Code."

                    LayoutCachedLeft =1500
                    LayoutCachedTop =120
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =480
                    Begin
                        Begin Label
                            Left =120
                            Top =120
                            Width =1320
                            Height =360
                            FontSize =12
                            ForeColor =3422101
                            Name ="Unit_Code_Label"
                            Caption ="Park Code:"
                            LayoutCachedLeft =120
                            LayoutCachedTop =120
                            LayoutCachedWidth =1440
                            LayoutCachedHeight =480
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    IMESentenceMode =3
                    Left =9900
                    Top =120
                    Width =720
                    Height =360
                    FontSize =12
                    FontWeight =700
                    TabIndex =2
                    ForeColor =3422101
                    Name ="Text46"
                    ControlSource ="Visit_Year"
                    StatusBarText ="Park Code."

                    LayoutCachedLeft =9900
                    LayoutCachedTop =120
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =480
                    Begin
                        Begin Label
                            Left =8640
                            Top =120
                            Width =1200
                            Height =360
                            FontSize =12
                            ForeColor =3422101
                            Name ="Label47"
                            Caption ="Visit Year:"
                            LayoutCachedLeft =8640
                            LayoutCachedTop =120
                            LayoutCachedWidth =9840
                            LayoutCachedHeight =480
                        End
                    End
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =720
            BreakLevel =2
            Name ="GroupHeader2"
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    BackStyle =1
                    IMESentenceMode =3
                    Left =645
                    Top =60
                    Width =495
                    Height =300
                    FontSize =10
                    Name ="Quad"
                    ControlSource ="=IIf(IsNull([tbxQuadCheck]),\"-\",[tbxQuadCheck])"
                    StatusBarText ="Quadrat number"

                    LayoutCachedLeft =645
                    LayoutCachedTop =60
                    LayoutCachedWidth =1140
                    LayoutCachedHeight =360
                    BackThemeColorIndex =1
                    Begin
                        Begin Label
                            BackStyle =1
                            Left =60
                            Top =60
                            Width =585
                            Height =300
                            FontSize =10
                            Name ="Quad_Label"
                            Caption ="Quad:"
                            LayoutCachedLeft =60
                            LayoutCachedTop =60
                            LayoutCachedWidth =645
                            LayoutCachedHeight =360
                            BackThemeColorIndex =1
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Top =420
                    Width =600
                    Height =300
                    FontSize =10
                    Name ="Tag_No_Label"
                    Caption ="Tag #"
                    Tag ="DetachedLabel"
                    LayoutCachedTop =420
                    LayoutCachedWidth =600
                    LayoutCachedHeight =720
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =600
                    Top =420
                    Width =1680
                    Height =300
                    FontSize =10
                    Name ="Species_Label"
                    Caption ="Species"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =600
                    LayoutCachedTop =420
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =720
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =2280
                    Top =420
                    Width =1200
                    Height =300
                    FontSize =10
                    Name ="DBH_Label"
                    Caption ="DBH/DRC"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =2280
                    LayoutCachedTop =420
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =720
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =3420
                    Top =420
                    Width =1200
                    Height =300
                    FontSize =10
                    Name ="Crown_Class_Label"
                    Caption ="Crown Class"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =3420
                    LayoutCachedTop =420
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =720
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =4620
                    Top =420
                    Width =1380
                    Height =300
                    FontSize =10
                    Name ="Class_Description_Label"
                    Caption ="Crown Health"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =4620
                    LayoutCachedTop =420
                    LayoutCachedWidth =6000
                    LayoutCachedHeight =720
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =6000
                    Top =420
                    Width =4830
                    Height =300
                    FontSize =10
                    Name ="Notes_Label"
                    Caption ="Notes"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =6000
                    LayoutCachedTop =420
                    LayoutCachedWidth =10830
                    LayoutCachedHeight =720
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    BackStyle =1
                    IMESentenceMode =3
                    Left =1155
                    Top =60
                    Width =810
                    Height =300
                    FontSize =10
                    TabIndex =1
                    Name ="tbxQuadCheck"
                    ControlSource ="Quad"
                    StatusBarText ="Quadrat number"

                    LayoutCachedLeft =1155
                    LayoutCachedTop =60
                    LayoutCachedWidth =1965
                    LayoutCachedHeight =360
                    BackThemeColorIndex =1
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =60
                    Width =10620
                    Name ="Line40"
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =60
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =360
                    Width =10620
                    Name ="Line41"
                    LayoutCachedLeft =60
                    LayoutCachedTop =360
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =360
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =600
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =540
                    Height =270
                    Name ="Tag_No"
                    ControlSource ="Tag_No"
                    StatusBarText ="Tag number"

                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =660
                    Top =60
                    Width =1620
                    Height =270
                    ColumnWidth =3345
                    TabIndex =1
                    Name ="Utah_Species"
                    ControlSource ="Utah_Species"
                    StatusBarText ="Utah Species (Welsh et al 2003)"

                    LayoutCachedLeft =660
                    LayoutCachedTop =60
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    DecimalPlaces =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2340
                    Top =60
                    Width =480
                    Height =270
                    TabIndex =2
                    Name ="DBH"
                    ControlSource ="DBH"
                    StatusBarText ="Diameter at breast height in centimeters"

                    LayoutCachedLeft =2340
                    LayoutCachedTop =60
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3420
                    Top =60
                    Width =1140
                    Height =270
                    TabIndex =3
                    Name ="Crown_Class"
                    ControlSource ="Crown_Class"
                    StatusBarText ="Crown class"

                    LayoutCachedLeft =3420
                    LayoutCachedTop =60
                    LayoutCachedWidth =4560
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6000
                    Top =60
                    Width =4740
                    Height =540
                    TabIndex =4
                    Name ="Notes"
                    ControlSource ="Notes"
                    StatusBarText ="Notes about any significant damage to a living tree"

                    LayoutCachedLeft =6000
                    LayoutCachedTop =60
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =600
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =4620
                    Top =60
                    Width =1320
                    Height =270
                    ColumnWidth =1845
                    TabIndex =5
                    Name ="Class_Description"
                    ControlSource ="Class_Description"
                    StatusBarText ="Health class description"

                    LayoutCachedLeft =4620
                    LayoutCachedTop =60
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2880
                    Top =60
                    Width =480
                    Height =270
                    TabIndex =6
                    Name ="DType"
                    ControlSource ="DType"
                    StatusBarText ="Diameter type indicator - dbh or DRC"

                    LayoutCachedLeft =2880
                    LayoutCachedTop =60
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =330
                End
            End
        End
        Begin PageFooter
            Height =270
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Width =4560
                    Height =270
                    Name ="Text21"
                    ControlSource ="=Now()"
                    Format ="Long Date"

                    LayoutCachedLeft =60
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4740
                    Width =5820
                    Height =270
                    TabIndex =1
                    Name ="Text22"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"

                    LayoutCachedLeft =4740
                    LayoutCachedWidth =10560
                    LayoutCachedHeight =270
                End
                Begin Line
                    Left =60
                    Width =10620
                    Name ="Line42"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =10680
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
