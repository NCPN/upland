Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =127
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =12186
    DatasheetFontHeight =9
    ItemSuffix =51
    Left =3840
    Top =4350
    Right =15465
    Bottom =9705
    DatasheetGridlinesColor =12632256
    Filter ="[Location_ID]='20100715151932-871445834.636688'"
    RecSrcDt = Begin
        0x9becc7edac0fe340
    End
    RecordSource ="tbl_Locations"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            SpecialEffect =2
            FontName ="Tahoma"
        End
        Begin Subform
            SpecialEffect =2
        End
        Begin Section
            CanGrow = NotDefault
            Height =8640
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =3240
                    Top =120
                    Width =3780
                    Height =480
                    FontSize =18
                    FontWeight =700
                    Name ="Label0"
                    Caption ="Plot Establishment"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =720
                    Top =780
                    Width =540
                    Name ="Unit_Code"
                    ControlSource ="Unit_Code"
                    StatusBarText ="Park Code."
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =180
                            Top =780
                            Width =480
                            Height =240
                            FontWeight =700
                            Name ="Label1"
                            Caption ="Park"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2280
                    Top =780
                    Width =600
                    TabIndex =1
                    Name ="Plot_ID"
                    ControlSource ="Plot_ID"
                    StatusBarText ="Plot identifier"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =1560
                            Top =780
                            Width =660
                            Height =240
                            FontWeight =700
                            Name ="Label2"
                            Caption ="Plot ID"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3720
                    Top =780
                    Width =960
                    TabIndex =2
                    Name ="Plot_Date"
                    ControlSource ="Plot_Date"
                    Format ="Short Date"
                    StatusBarText ="Date plot established if different from site characterization date"
                    InputMask ="99/99/0000;0;_"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =3180
                            Top =780
                            Width =480
                            Height =240
                            FontWeight =700
                            Name ="Label3"
                            Caption ="Date"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1560
                    Top =1200
                    Width =1020
                    TabIndex =4
                    Name ="Plot_E_Coord"
                    ControlSource ="E_Coord"
                    StatusBarText ="UTM East of Centroid if plot established on different visit from site characteri"
                        "zation"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =180
                            Top =1200
                            Width =1380
                            Height =240
                            FontWeight =700
                            Name ="Label4"
                            Caption ="Centroid UTM E"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3540
                    Top =1200
                    Width =1019
                    TabIndex =5
                    Name ="Plot_N_Coord"
                    ControlSource ="N_Coord"
                    StatusBarText ="UTM North of Centroid (Y_Coord) if plot established on different visit from site"
                        " characterization"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =2880
                            Top =1200
                            Width =600
                            Height =240
                            FontWeight =700
                            Name ="Label5"
                            Caption ="UTM N"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =840
                    Top =1620
                    Width =480
                    TabIndex =6
                    Name ="Plot_Slope"
                    ControlSource ="Plot_Slope"
                    StatusBarText ="Plot slope in percent - 1 decimal"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =180
                            Top =1620
                            Width =600
                            Height =240
                            FontWeight =700
                            Name ="Label7"
                            Caption ="Slope"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2340
                    Top =1620
                    Width =480
                    TabIndex =7
                    Name ="Plot_Aspect"
                    ControlSource ="Plot_Aspect"
                    StatusBarText ="Plot aspect in degrees - 1 decimal"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =1620
                            Top =1620
                            Width =660
                            Height =240
                            FontWeight =700
                            Name ="Label8"
                            Caption ="Aspect"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4020
                    Top =1620
                    Width =480
                    TabIndex =8
                    Name ="Azimuth"
                    ControlSource ="Azimuth"
                    StatusBarText ="Direction from origin to end of center transect in degrees - 1 decimal"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =3120
                            Top =1620
                            Width =840
                            Height =240
                            FontWeight =700
                            Name ="Label9"
                            Caption ="Azimuth"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2100
                    Top =2700
                    Width =1080
                    TabIndex =9
                    Name ="T1O_UTME"
                    ControlSource ="T1O_UTME"
                    StatusBarText ="UTM East of Transect 1 origin"
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =1200
                            Top =2700
                            Width =900
                            Height =240
                            FontWeight =700
                            Name ="Label10"
                            Caption ="1"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3180
                    Top =2700
                    Width =1080
                    TabIndex =10
                    Name ="T1O_UTMN"
                    ControlSource ="T1O_UTMN"
                    StatusBarText ="UTM North of Transect 1 origin"
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =2100
                            Top =2280
                            Width =1080
                            Height =420
                            FontWeight =700
                            Name ="Label11"
                            Caption ="UTM E"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4260
                    Top =2700
                    Width =960
                    TabIndex =11
                    Name ="T1O_Rebar"
                    ControlSource ="T1O_Rebar"
                    StatusBarText ="Distance from origin of rebar in meters - 1 decimal"
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =4260
                            Top =2040
                            Width =960
                            Height =660
                            FontWeight =700
                            Name ="Label12"
                            Caption ="Rebar location (m)"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5220
                    Top =2700
                    Width =1080
                    TabIndex =12
                    Name ="T1E_UTME"
                    ControlSource ="T1E_UTME"
                    StatusBarText ="UTM East of Transect 1 end"
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =5220
                            Top =2280
                            Width =1080
                            Height =420
                            FontWeight =700
                            Name ="Label13"
                            Caption ="UTM E"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6300
                    Top =2700
                    Width =1080
                    TabIndex =13
                    Name ="T1E_UTMN"
                    ControlSource ="T1E_UTMN"
                    StatusBarText ="UTM North of Transect 1 end"
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =6300
                            Top =2280
                            Width =1080
                            Height =420
                            FontWeight =700
                            Name ="Label14"
                            Caption ="UTM N"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7380
                    Top =2700
                    Width =959
                    TabIndex =14
                    Name ="T1E_Rebar"
                    ControlSource ="T1E_Rebar"
                    StatusBarText ="Distance from end of rebar in meters - 1 decimal"
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =7380
                            Top =2040
                            Width =959
                            Height =660
                            FontWeight =700
                            Name ="Label15"
                            Caption ="Rebar location (m)"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2100
                    Top =2940
                    Width =1080
                    TabIndex =16
                    Name ="T2O_UTME"
                    ControlSource ="T2O_UTME"
                    StatusBarText ="UTM East of Transect 2 origin"
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =1200
                            Top =2940
                            Width =900
                            Height =240
                            FontWeight =700
                            Name ="Label16"
                            Caption ="2"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3180
                    Top =2940
                    Width =1080
                    TabIndex =17
                    Name ="T2O_UTMN"
                    ControlSource ="T2O_UTMN"
                    StatusBarText ="UTM North of Transect 2 origin"
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4260
                    Top =2940
                    Width =960
                    TabIndex =18
                    Name ="T2O_Rebar"
                    ControlSource ="T2O_Rebar"
                    StatusBarText ="Distance from origin of rebar in meters - 1 decimal"
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5220
                    Top =2940
                    Width =1080
                    TabIndex =19
                    Name ="T2E_UTME"
                    ControlSource ="T2E_UTME"
                    StatusBarText ="UTM East of Transect 2 end"
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6300
                    Top =2940
                    Width =1080
                    TabIndex =20
                    Name ="T2E_UTMN"
                    ControlSource ="T2E_UTMN"
                    StatusBarText ="UTM North of Transect 2 end"
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7380
                    Top =2940
                    Width =960
                    TabIndex =21
                    Name ="T2E_Rebar"
                    ControlSource ="T2E_Rebar"
                    StatusBarText ="Distance from end of rebar in meters - 1 decimal"
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2100
                    Top =3180
                    Width =1080
                    TabIndex =23
                    Name ="T3O_UTME"
                    ControlSource ="T3O_UTME"
                    StatusBarText ="UTM East of Transect 3 origin"
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =1200
                            Top =3180
                            Width =900
                            Height =240
                            FontWeight =700
                            Name ="Label22"
                            Caption ="3"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3180
                    Top =3180
                    Width =1080
                    TabIndex =24
                    Name ="T3O_UTMN"
                    ControlSource ="T3O_UTMN"
                    StatusBarText ="UTM North of Transect 3 origin"
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4260
                    Top =3180
                    Width =960
                    TabIndex =25
                    Name ="T3O_Rebar"
                    ControlSource ="T3O_Rebar"
                    StatusBarText ="Distance from origin of rebar in meters - 1 decimal"
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5220
                    Top =3180
                    Width =1080
                    TabIndex =26
                    Name ="T3E_UTME"
                    ControlSource ="T3E_UTME"
                    StatusBarText ="UTM East of Transect 3 end"
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6300
                    Top =3180
                    Width =1080
                    TabIndex =27
                    Name ="T3E_UTMN"
                    ControlSource ="T3E_UTMN"
                    StatusBarText ="UTM North of Transect 3 end"
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7380
                    Top =3180
                    Width =960
                    TabIndex =28
                    Name ="T3E_Rebar"
                    ControlSource ="T3E_Rebar"
                    StatusBarText ="Distance from end of rebar in meters - 1 decimal"
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =1200
                    Top =3600
                    Width =9900
                    Height =903
                    TabIndex =30
                    Name ="Plot_Directions"
                    ControlSource ="Plot_Directions"
                    StatusBarText ="Directions to plot"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =240
                            Top =3600
                            Width =960
                            Height =480
                            FontWeight =700
                            Name ="Label28"
                            Caption ="Directions To Plot"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8340
                    Top =2700
                    Width =1080
                    TabIndex =15
                    Name ="T1_Elevation"
                    ControlSource ="T1_Elevation"
                    StatusBarText ="Elevation in meters of transect 1"
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =8340
                            Top =2040
                            Width =1080
                            Height =660
                            FontWeight =700
                            Name ="Label29"
                            Caption ="Elevation (m)"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8340
                    Top =2940
                    Width =1080
                    TabIndex =22
                    Name ="T2_Elevation"
                    ControlSource ="T2_Elevation"
                    StatusBarText ="Elevation in meters of transect 2"
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8340
                    Top =3180
                    Width =1080
                    TabIndex =29
                    Name ="T3_Elevation"
                    ControlSource ="T3_Elevation"
                    StatusBarText ="Elevation in meters of transect 3"
                End
                Begin ComboBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =735
                    Left =5880
                    Top =780
                    Width =2100
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Observer"
                    ControlSource ="Observer"
                    RowSourceType ="Table/Query"
                    RowSource ="qry_Contacts"
                    ColumnWidths ="0;735"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4980
                            Top =780
                            Width =900
                            Height =245
                            FontWeight =700
                            Name ="Observer_Label"
                            Caption ="Observer"
                        End
                    End
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =1200
                    Top =2040
                    Width =900
                    Height =660
                    FontWeight =700
                    Name ="Label34"
                    Caption ="Transect"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =3180
                    Top =2280
                    Width =1080
                    Height =420
                    FontWeight =700
                    Name ="Label35"
                    Caption ="UTM N"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    Left =2100
                    Top =2040
                    Width =2160
                    Height =240
                    FontWeight =700
                    Name ="Label36"
                    Caption ="Rebar at transect origin"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    Left =5220
                    Top =2040
                    Width =2160
                    Height =240
                    FontWeight =700
                    Name ="Label37"
                    Caption ="Rebar at transect end"
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    Left =180
                    Top =2040
                    Width =1020
                    Height =540
                    FontSize =10
                    FontWeight =700
                    Name ="Label38"
                    Caption ="Transect Location"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10080
                    Top =240
                    Width =1035
                    Height =300
                    TabIndex =41
                    ForeColor =255
                    Name ="ButtonClose"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =8520
                    Top =480
                    Width =780
                    TabIndex =42
                    Name ="SiteDate"
                    ControlSource ="SiteDate"
                    Format ="Short Date"
                    StatusBarText ="Date site characterized"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =8520
                    Top =720
                    Width =780
                    TabIndex =43
                    Name ="E_Coord"
                    ControlSource ="E_Coord"
                    StatusBarText ="UTM East of Centroid (X_Coord)"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =8520
                    Top =960
                    Width =780
                    TabIndex =44
                    Name ="N_Coord"
                    ControlSource ="N_Coord"
                    StatusBarText ="UTM North of Centroid (Y_Coord)"
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =2880
                    Top =4680
                    Width =480
                    TabIndex =32
                    Name ="SlopeA"
                    ControlSource ="SlopeA"
                    StatusBarText ="F/W Slope for plot side A in percent"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2640
                            Top =4680
                            Width =255
                            Height =240
                            FontWeight =700
                            Name ="Label117"
                            Caption ="A"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4800
                    Top =4680
                    Width =479
                    TabIndex =34
                    Name ="SlopeB"
                    ControlSource ="SlopeB"
                    StatusBarText ="F/W Slope for plot side B in percent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4500
                            Top =4680
                            Width =255
                            Height =240
                            FontWeight =700
                            Name ="Label118"
                            Caption ="B"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6720
                    Top =4680
                    Width =479
                    TabIndex =36
                    Name ="SlopeC"
                    ControlSource ="SlopeC"
                    StatusBarText ="F/W Slope for plot side C in percent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6420
                            Top =4680
                            Width =255
                            Height =240
                            FontWeight =700
                            Name ="Label119"
                            Caption ="C"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8640
                    Top =4680
                    Width =479
                    TabIndex =38
                    Name ="SlopeD"
                    ControlSource ="SlopeD"
                    StatusBarText ="F/W Slope for plot side D in percent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8340
                            Top =4680
                            Width =255
                            Height =240
                            FontWeight =700
                            Name ="Label120"
                            Caption ="D"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =87
                    Left =1020
                    Top =5100
                    Width =10365
                    TabIndex =40
                    Name ="fsub_FW_Monument"
                    SourceObject ="Form.fsub_FW_Monument"
                    LinkChildFields ="Location_ID"
                    LinkMasterFields ="Location_ID"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =60
                            Top =5100
                            Width =960
                            Height =420
                            FontWeight =700
                            Name ="fsub_FW_Monument Label"
                            Caption ="Monument Trees"
                            EventProcPrefix ="fsub_FW_Monument_Label"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =570
                    Left =3420
                    Top =4680
                    Width =780
                    TabIndex =33
                    Name ="SlopeAUD"
                    ControlSource ="SlopeAUD"
                    RowSourceType ="Value List"
                    RowSource ="\"up\";\"down\""
                    ColumnWidths ="570"
                End
                Begin Label
                    OverlapFlags =85
                    Left =180
                    Top =4680
                    Width =2160
                    Height =240
                    FontWeight =700
                    Name ="LabelSlope"
                    Caption ="Plot Side Slopes (%) U/D"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =570
                    Left =5340
                    Top =4680
                    Width =780
                    TabIndex =35
                    Name ="SlopeBUD"
                    ControlSource ="SlopeBUD"
                    RowSourceType ="Value List"
                    RowSource ="\"up\";\"down\""
                    ColumnWidths ="570"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =570
                    Left =7260
                    Top =4680
                    Width =780
                    TabIndex =37
                    Name ="SlopeCUD"
                    ControlSource ="SlopeCUD"
                    RowSourceType ="Value List"
                    RowSource ="\"up\";\"down\""
                    ColumnWidths ="570"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =570
                    Left =9180
                    Top =4680
                    Width =780
                    TabIndex =39
                    Name ="SlopeDUD"
                    ControlSource ="SlopeDUD"
                    RowSourceType ="Value List"
                    RowSource ="\"up\";\"down\""
                    ColumnWidths ="570"
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4185
                    Top =7380
                    Width =810
                    Height =300
                    TabIndex =45
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Bearing_A"
                    ControlSource ="Bearing_A"
                    StatusBarText ="Bearing of the plot slope + 180 in degrees"
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =4185
                            Top =7140
                            Width =810
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Bearing_A_Label"
                            Caption ="A"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =255
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4980
                    Top =7380
                    Width =810
                    Height =300
                    TabIndex =46
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Bearing_B"
                    ControlSource ="Bearing_B"
                    StatusBarText ="Bearing of transect 1 in degrees"
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =223
                            TextAlign =2
                            Left =4980
                            Top =7140
                            Width =810
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Bearing_B_Label"
                            Caption ="B"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =127
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5805
                    Top =7380
                    Width =810
                    Height =300
                    TabIndex =47
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Bearing_C"
                    ControlSource ="Bearing_C"
                    StatusBarText ="Bearing of transect 3 + 180"
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =5805
                            Top =7140
                            Width =810
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Bearing_C_Label"
                            Caption ="C"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6645
                    Top =7380
                    Width =810
                    Height =300
                    TabIndex =48
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Bearing_D"
                    ControlSource ="Bearing_D"
                    StatusBarText ="Bearing of the plot slope"
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =6645
                            Top =7140
                            Width =810
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Bearing_D_Label"
                            Caption ="D"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OldBorderStyle =1
                    OverlapFlags =127
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4185
                    Top =7680
                    Width =810
                    Height =300
                    TabIndex =49
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Slope_A"
                    ControlSource ="Slope_A"
                    StatusBarText ="Slope of transect A to nearest half percent"
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =2880
                            Top =7380
                            Width =1290
                            Height =300
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Slope_A_Label"
                            Caption ="bearing (deg)"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OldBorderStyle =1
                    OverlapFlags =255
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4980
                    Top =7680
                    Width =810
                    Height =300
                    TabIndex =50
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Slope_B"
                    ControlSource ="Slope_B"
                    StatusBarText ="Slope of transect B to nearest half percent"
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =119
                            TextAlign =2
                            Left =2880
                            Top =7680
                            Width =1290
                            Height =300
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Slope_B_Label"
                            Caption ="slope (%)"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OldBorderStyle =1
                    OverlapFlags =119
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5805
                    Top =7680
                    Width =810
                    Height =300
                    TabIndex =51
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Slope_C"
                    ControlSource ="Slope_C"
                    StatusBarText ="Slope of transect C to nearest half percent"
                End
                Begin TextBox
                    DecimalPlaces =1
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6645
                    Top =7680
                    Width =810
                    Height =300
                    TabIndex =52
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Slope_D"
                    ControlSource ="Slope_D"
                    StatusBarText ="Slope of transect C to nearest half percent"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =4140
                    Top =6780
                    Width =1860
                    Height =300
                    FontSize =10
                    FontWeight =700
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Fuels_Transect_Label"
                    Caption ="Fuels Transect"
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =11100
                    Top =3600
                    Width =306
                    Height =306
                    TabIndex =31
                    ForeColor =0
                    Name ="ButtonZoomPlotDirections"
                    Caption ="Zoom Caption"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x280000001000000010000000010004000000000080000000c40e0000c40e0000 ,
                        0x1000000000000000000000000000800000800000008080008000000080008000 ,
                        0x80800000c0c0c000808080000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00666666666666446666666666666474466666666666474446 ,
                        0x666666666474446666660000474446666600777f8444666660877777f8086666 ,
                        0x607770777f066666077770777770666607777077777066660700000007706666 ,
                        0x07f770777770666660ff707777066666608ff077780666666600777700666666 ,
                        0x6666000066666666
                    End
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0xc0c0c00080808000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="Zoom Caption"
                    Picture ="C:\\arcgis\\arcexe9x\\odetools\\Bitmaps\\zoomin.bmp"
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

Private Sub ButtonClose_Click()
On Error GoTo Err_ButtonClose_Click


    DoCmd.Close

Exit_ButtonClose_Click:
    Exit Sub

Err_ButtonClose_Click:
    MsgBox Err.Description
    Resume Exit_ButtonClose_Click
    
End Sub

Private Sub ButtonZoomPlotDirections_Click()
On Error GoTo Err_ButtonPlotDirections_Click

  Me!Plot_Directions.SetFocus
  SendKeys ("+{F2}")
  
Exit_ButtonPlotDirections_Click:
    Exit Sub

Err_ButtonPlotDirections_Click:
    MsgBox Err.Description
    Resume Exit_ButtonPlotDirections_Click

End Sub

Private Sub Form_Load()
  Dim Veg_Type As Variant
    Veg_Type = DLookup("[Vegetation_Type]", "tbl_Locations", "[Location_ID] = '" & Me!Location_ID & "'")
    If Not IsNull(Veg_Type) And (Veg_Type = "grassland/shrubland") Then
      Me!fsub_FW_Monument.Form.Visible = False
    End If
    If Not IsNull(Veg_Type) And ((Veg_Type = "oak scrub") Or (Veg_Type = "grassland/shrubland")) Then
      Me!SlopeA.Visible = False
      Me!SlopeB.Visible = False
      Me!SlopeC.Visible = False
      Me!SlopeD.Visible = False
      Me!SlopeAUD.Visible = False
      Me!SlopeBUD.Visible = False
      Me!SlopeCUD.Visible = False
      Me!SlopeDUD.Visible = False
      Me!LabelSlope.Visible = False
      Me!Fuels_Transect_Label.Visible = False  ' Hide fuels fields
      Me!Bearing_A_Label.Visible = False
      Me!Bearing_B_Label.Visible = False
      Me!Bearing_C_Label.Visible = False
      Me!Bearing_D_Label.Visible = False
      Me!Slope_A_Label.Visible = False
      Me!Slope_B_Label.Visible = False
      Me!Bearing_A.Visible = False
      Me!Bearing_B.Visible = False
      Me!Bearing_C.Visible = False
      Me!Bearing_D.Visible = False
      Me!Slope_A.Visible = False
      Me!Slope_B.Visible = False
      Me!Slope_C.Visible = False
      Me!Slope_D.Visible = False
    End If
If IsNull(Me!Plot_Date) Then
  Me!Plot_Date = Me!SiteDate  ' Set default
End If
If IsNull(Me!Plot_E_Coord) Then
  Me!Plot_E_Coord = E_Coord   ' plot
End If
If IsNull(Me!Plot_N_Coord) Then
  Me!Plot_N_Coord = N_Coord   ' values
End If

End Sub
