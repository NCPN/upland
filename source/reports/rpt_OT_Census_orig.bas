Version =21
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =126
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10800
    DatasheetFontHeight =9
    ItemSuffix =44
    Left =195
    Top =360
    Right =9465
    Bottom =7260
    DatasheetGridlinesColor =12632256
    Filter ="[Unit_Code] = 'BLCA' AND [Plot_Id] = 203AND [Visit_Year] = '2018'"
    RecSrcDt = Begin
        0x04ae7a726ca8e340
    End
    RecordSource ="qry_sel_OT_Census_Report"
    Caption ="rpt_OT_Census"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x40020000d002000040020000d002000000000000602700002805000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
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
                    Left =120
                    Top =420
                    Width =10512
                    Name ="Line34"
                    LayoutCachedLeft =120
                    LayoutCachedTop =420
                    LayoutCachedWidth =10632
                    LayoutCachedHeight =420
                End
            End
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =360
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    IMESentenceMode =3
                    Left =1620
                    Height =360
                    FontSize =12
                    FontWeight =700
                    Name ="Unit_Code"
                    ControlSource ="Unit_Code"
                    StatusBarText ="Park Code."

                    LayoutCachedLeft =1620
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =360
                    Begin
                        Begin Label
                            Left =60
                            Width =1500
                            Height =360
                            FontSize =12
                            Name ="Unit_Code_Label"
                            Caption ="Park Code"
                            LayoutCachedLeft =60
                            LayoutCachedWidth =1560
                            LayoutCachedHeight =360
                        End
                    End
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =360
            BreakLevel =1
            Name ="GroupHeader1"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    IMESentenceMode =3
                    Left =1620
                    Height =360
                    FontSize =12
                    Name ="Plot_ID"
                    ControlSource ="Plot_ID"
                    StatusBarText ="Plot identifier"

                    LayoutCachedLeft =1620
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =360
                    Begin
                        Begin Label
                            Left =60
                            Width =1500
                            Height =360
                            FontSize =12
                            FontWeight =400
                            Name ="Plot_ID_Label"
                            Caption ="Plot"
                            LayoutCachedLeft =60
                            LayoutCachedWidth =1560
                            LayoutCachedHeight =360
                        End
                    End
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =780
            BreakLevel =2
            Name ="GroupHeader2"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    IMESentenceMode =3
                    Left =660
                    Top =60
                    Width =420
                    Height =300
                    FontSize =10
                    Name ="Quad"
                    ControlSource ="=IIf(IsNull([tbxQuadCheck]),\"-\",[tbxQuadCheck])"
                    StatusBarText ="Quadrat number"

                    Begin
                        Begin Label
                            Left =60
                            Top =60
                            Width =540
                            Height =300
                            FontSize =10
                            Name ="Quad_Label"
                            Caption ="Quad"
                        End
                    End
                End
                Begin Label
                    TextAlign =2
                    Left =60
                    Top =420
                    Width =540
                    Height =270
                    FontSize =10
                    Name ="Tag_No_Label"
                    Caption ="Tag #"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =660
                    Top =420
                    Width =1620
                    Height =270
                    FontSize =10
                    Name ="Species_Label"
                    Caption ="Species"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =660
                    LayoutCachedTop =420
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =690
                End
                Begin Label
                    TextAlign =2
                    Left =2340
                    Top =420
                    Width =1005
                    Height =270
                    FontSize =10
                    Name ="DBH_Label"
                    Caption ="DBH/DRC"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =2340
                    LayoutCachedTop =420
                    LayoutCachedWidth =3345
                    LayoutCachedHeight =690
                End
                Begin Label
                    TextAlign =2
                    Left =3420
                    Top =420
                    Width =1155
                    Height =270
                    FontSize =10
                    Name ="Crown_Class_Label"
                    Caption ="Crown Class"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =3420
                    LayoutCachedTop =420
                    LayoutCachedWidth =4575
                    LayoutCachedHeight =690
                End
                Begin Label
                    TextAlign =2
                    Left =4620
                    Top =420
                    Width =1320
                    Height =270
                    FontSize =10
                    Name ="Class_Description_Label"
                    Caption ="Crown Health"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =4620
                    LayoutCachedTop =420
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =690
                End
                Begin Label
                    TextAlign =2
                    Left =6000
                    Top =420
                    Width =4800
                    Height =270
                    FontSize =10
                    Name ="Notes_Label"
                    Caption ="Notes"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =6000
                    LayoutCachedTop =420
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =690
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    IMESentenceMode =3
                    Left =1200
                    Top =60
                    Width =420
                    Height =300
                    FontSize =10
                    TabIndex =1
                    Name ="tbxQuadCheck"
                    ControlSource ="Quad"
                    StatusBarText ="Quadrat number"

                    LayoutCachedLeft =1200
                    LayoutCachedTop =60
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =360
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
            Height =840
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
                    Left =4620
                    Top =60
                    Width =1320
                    Height =270
                    ColumnWidth =1845
                    TabIndex =4
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
                    Left =6000
                    Top =60
                    Width =4800
                    Height =780
                    TabIndex =5
                    Name ="Notes"
                    ControlSource ="Notes"
                    StatusBarText ="Notes about any significant damage to a living tree"

                    LayoutCachedLeft =6000
                    LayoutCachedTop =60
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =840
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
