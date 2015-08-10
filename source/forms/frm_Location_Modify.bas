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
    Width =10080
    DatasheetFontHeight =9
    ItemSuffix =45
    Left =5685
    Top =3195
    Right =14325
    Bottom =8880
    DatasheetGridlinesColor =12632256
    Filter ="[Location_ID]='20090910180136-103022634.983063'"
    RecSrcDt = Begin
        0x527962004c51e340
    End
    RecordSource ="tbl_Locations"
    BeforeUpdate ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
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
            BorderLineStyle =0
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin Section
            Height =7200
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2565
                    Top =120
                    Width =5115
                    Height =480
                    FontSize =18
                    FontWeight =700
                    Name ="Label0"
                    Caption ="Location Coordinate Change"
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
                    Left =1560
                    Top =1260
                    Width =1200
                    TabIndex =3
                    Name ="Plot_E_Coord"
                    ControlSource ="E_Coord"
                    StatusBarText ="UTM East of Centroid if plot established on different visit from site characteri"
                        "zation"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =180
                            Top =1260
                            Width =1320
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
                    Left =4020
                    Top =1260
                    Width =1199
                    TabIndex =4
                    Name ="Plot_N_Coord"
                    ControlSource ="N_Coord"
                    StatusBarText ="UTM North of Centroid (Y_Coord) if plot established on different visit from site"
                        " characterization"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =3360
                            Top =1260
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
                    Top =1740
                    Width =480
                    TabIndex =5
                    Name ="Plot_Slope"
                    ControlSource ="Plot_Slope"
                    StatusBarText ="Plot slope in percent - 1 decimal"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =180
                            Top =1740
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
                    Top =1740
                    Width =480
                    TabIndex =6
                    Name ="Plot_Aspect"
                    ControlSource ="Plot_Aspect"
                    StatusBarText ="Plot aspect in degrees - 1 decimal"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =1620
                            Top =1740
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
                    Top =1740
                    Width =480
                    TabIndex =7
                    Name ="Azimuth"
                    ControlSource ="Azimuth"
                    StatusBarText ="Direction from origin to end of center transect in degrees - 1 decimal"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =3120
                            Top =1740
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
                    Left =1800
                    Top =3360
                    Width =1080
                    TabIndex =8
                    Name ="T1O_UTME"
                    ControlSource ="T1O_UTME"
                    StatusBarText ="UTM East of Transect 1 origin"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =900
                            Top =3360
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
                    Left =2880
                    Top =3360
                    Width =1080
                    TabIndex =9
                    Name ="T1O_UTMN"
                    ControlSource ="T1O_UTMN"
                    StatusBarText ="UTM North of Transect 1 origin"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =1800
                            Top =2940
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
                    Left =3960
                    Top =3360
                    Width =960
                    TabIndex =10
                    Name ="T1O_Rebar"
                    ControlSource ="T1O_Rebar"
                    StatusBarText ="Distance from origin of rebar in meters - 1 decimal"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =3960
                            Top =2700
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
                    Left =4920
                    Top =3360
                    Width =1080
                    TabIndex =11
                    Name ="T1E_UTME"
                    ControlSource ="T1E_UTME"
                    StatusBarText ="UTM East of Transect 1 end"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =4920
                            Top =2940
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
                    Left =6000
                    Top =3360
                    Width =1080
                    TabIndex =12
                    Name ="T1E_UTMN"
                    ControlSource ="T1E_UTMN"
                    StatusBarText ="UTM North of Transect 1 end"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =6000
                            Top =2940
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
                    Left =7080
                    Top =3360
                    Width =959
                    TabIndex =13
                    Name ="T1E_Rebar"
                    ControlSource ="T1E_Rebar"
                    StatusBarText ="Distance from end of rebar in meters - 1 decimal"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =7080
                            Top =2700
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
                    Left =1800
                    Top =3600
                    Width =1080
                    TabIndex =15
                    Name ="T2O_UTME"
                    ControlSource ="T2O_UTME"
                    StatusBarText ="UTM East of Transect 2 origin"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =900
                            Top =3600
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
                    Left =2880
                    Top =3600
                    Width =1080
                    TabIndex =16
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
                    Left =3960
                    Top =3600
                    Width =960
                    TabIndex =17
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
                    Left =4920
                    Top =3600
                    Width =1080
                    TabIndex =18
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
                    Left =6000
                    Top =3600
                    Width =1080
                    TabIndex =19
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
                    Left =7080
                    Top =3600
                    Width =960
                    TabIndex =20
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
                    Left =1800
                    Top =3840
                    Width =1080
                    TabIndex =22
                    Name ="T3O_UTME"
                    ControlSource ="T3O_UTME"
                    StatusBarText ="UTM East of Transect 3 origin"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =900
                            Top =3840
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
                    Left =2880
                    Top =3840
                    Width =1080
                    TabIndex =23
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
                    Left =3960
                    Top =3840
                    Width =960
                    TabIndex =24
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
                    Left =4920
                    Top =3840
                    Width =1080
                    TabIndex =25
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
                    Left =6000
                    Top =3840
                    Width =1080
                    TabIndex =26
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
                    Left =7080
                    Top =3840
                    Width =960
                    TabIndex =27
                    Name ="T3E_Rebar"
                    ControlSource ="T3E_Rebar"
                    StatusBarText ="Distance from end of rebar in meters - 1 decimal"

                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =1140
                    Top =4500
                    Width =7740
                    Height =1143
                    TabIndex =29
                    Name ="Plot_Directions"
                    ControlSource ="Plot_Directions"
                    StatusBarText ="Directions to plot"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =180
                            Top =4500
                            Width =960
                            Height =720
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
                    Left =8040
                    Top =3360
                    Width =1080
                    TabIndex =14
                    Name ="T1_Elevation"
                    ControlSource ="T1_Elevation"
                    StatusBarText ="Elevation in meters of transect 1"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =8040
                            Top =2700
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
                    Left =8040
                    Top =3600
                    Width =1080
                    TabIndex =21
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
                    Left =8040
                    Top =3840
                    Width =1080
                    TabIndex =28
                    Name ="T3_Elevation"
                    ControlSource ="T3_Elevation"
                    StatusBarText ="Elevation in meters of transect 3"

                End
                Begin ComboBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =735
                    Left =7440
                    Top =780
                    Width =2100
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Recorder"
                    RowSourceType ="Table/Query"
                    RowSource ="qry_Contacts"
                    ColumnWidths ="0;735"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =4800
                            Top =780
                            Width =2640
                            Height =245
                            FontWeight =700
                            Name ="Observer_Label"
                            Caption ="Changes made by (Required):"
                        End
                    End
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =900
                    Top =2700
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
                    Left =2880
                    Top =2940
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
                    Left =1800
                    Top =2700
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
                    Left =4920
                    Top =2700
                    Width =2160
                    Height =240
                    FontWeight =700
                    Name ="Label37"
                    Caption ="Rebar at transect end"
                End
                Begin Label
                    OverlapFlags =85
                    Left =900
                    Top =2400
                    Width =1920
                    Height =240
                    FontSize =10
                    FontWeight =700
                    Name ="Label38"
                    Caption ="Transect Location"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4560
                    Top =6540
                    Width =1035
                    Height =300
                    TabIndex =31
                    Name ="ButtonClose"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =300
                    Top =180
                    Width =960
                    TabIndex =32
                    Name ="Location_ID"
                    ControlSource ="Location_ID"
                    StatusBarText ="M. Location identifier (Loc_ID)"

                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =8880
                    Top =4500
                    Width =306
                    Height =306
                    TabIndex =30
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
                        0x000301000000000000000000
                    End
                    ControlTipText ="Zoom Caption"
                    Picture ="zoomin.bmp"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =2880
                    Top =5820
                    Width =480
                    TabIndex =33
                    Name ="SlopeA"
                    ControlSource ="SlopeA"
                    StatusBarText ="F/W Slope for plot side A in percent"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2640
                            Top =5820
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
                    Top =5820
                    Width =479
                    TabIndex =35
                    Name ="SlopeB"
                    ControlSource ="SlopeB"
                    StatusBarText ="F/W Slope for plot side B in percent"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4500
                            Top =5820
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
                    Top =5820
                    Width =479
                    TabIndex =37
                    Name ="SlopeC"
                    ControlSource ="SlopeC"
                    StatusBarText ="F/W Slope for plot side C in percent"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6420
                            Top =5820
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
                    Top =5820
                    Width =479
                    TabIndex =39
                    Name ="SlopeD"
                    ControlSource ="SlopeD"
                    StatusBarText ="F/W Slope for plot side D in percent"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8340
                            Top =5820
                            Width =255
                            Height =240
                            FontWeight =700
                            Name ="Label120"
                            Caption ="D"
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
                    Top =5820
                    Width =780
                    TabIndex =34
                    Name ="SlopeAUD"
                    ControlSource ="SlopeAUD"
                    RowSourceType ="Value List"
                    RowSource ="\"up\";\"down\""
                    ColumnWidths ="570"

                End
                Begin Label
                    OverlapFlags =85
                    Left =180
                    Top =5820
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
                    Top =5820
                    Width =780
                    TabIndex =36
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
                    Top =5820
                    Width =780
                    TabIndex =38
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
                    Top =5820
                    Width =780
                    TabIndex =40
                    Name ="SlopeDUD"
                    ControlSource ="SlopeDUD"
                    RowSourceType ="Value List"
                    RowSource ="\"up\";\"down\""
                    ColumnWidths ="570"

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

Private Sub Form_BeforeUpdate(Cancel As Integer)

    On Error GoTo Err_Handler
    Dim strMsg As String
    Dim db As Database
    Dim History As DAO.Recordset
    Dim OldLocation As DAO.Recordset
    Dim strSQL As String
    
    strMsg = "Are you sure you want to update location coordinates?"
    strMsg = strMsg & chr(13) & chr(10) & "Click Yes to Save or No to Discard changes."
    If MsgBox(strMsg, vbQuestion + vbYesNo, "Update Location?") = vbNo Then
      Me.Undo
      GoTo Exit_Form_BeforeUpdate
    ElseIf IsNull(Me!Recorder) Then
      MsgBox "You must select a recorder name!"
      Me.Undo
      GoTo Exit_Form_BeforeUpdate
    Else
      Set db = CurrentDb
      strSQL = "Select * from tbl_Locations WHERE Location_ID = '" & Me!Location_ID & "'"
      Set OldLocation = db.OpenRecordset(strSQL)  '  Get unmodified location record
      If OldLocation.EOF Then
        MsgBox "Location record not found."
        GoTo Exit_Form_BeforeUpdate
      Else
        OldLocation.MoveFirst
      End If
      Set History = db.OpenRecordset("tbl_Location_History")
        History.AddNew                     ' Create a Location History record
        History!Location_History_ID = fxnGUIDGen
        History!Location_ID = Me!Location_ID
        History!Modify_Date = Now()        ' Date of update
        History!Recorder = Me!Recorder     ' Person committing update
        History!Unit_Code = Me!Unit_Code
        History!Plot_ID = Me!Plot_ID
        ' Modified to populate plot centroid coordinates (E_Coord, N_Coord) correctly. [HMT, 3/16/2015]
        ' Plot_E_Coord, Plot_N_Coord are no longer used.
        ' History!E_Coord = OldLocation!Plot_E_Coord
        ' History!N_Coord = OldLocation!Plot_N_Coord
        History!E_Coord = OldLocation!E_Coord            ' UTM easting of plot centroid
        History!N_Coord = OldLocation!N_Coord            ' UTM northing of plot centroid
        History!Plot_Slope = OldLocation!Plot_Slope
        History!Plot_Aspect = OldLocation!Plot_Aspect
        History!Azimuth = OldLocation!Azimuth
        History!T1O_UTME = OldLocation!T1O_UTME
        History!T1O_UTMN = OldLocation!T1O_UTMN
        History!T1O_Rebar = OldLocation!T1O_Rebar
        History!T1E_UTME = OldLocation!T1E_UTME
        History!T1E_UTMN = OldLocation!T1E_UTMN
        History!T1E_Rebar = OldLocation!T1E_Rebar
        History!T1_Elevation = OldLocation!T1_Elevation
        History!T2O_UTME = OldLocation!T2O_UTME
        History!T2O_UTMN = OldLocation!T2O_UTMN
        History!T2O_Rebar = OldLocation!T2O_Rebar
        History!T2E_UTME = OldLocation!T2E_UTME
        History!T2E_UTMN = OldLocation!T2E_UTMN
        History!T2E_Rebar = OldLocation!T2E_Rebar
        History!T2_Elevation = OldLocation!T2_Elevation
        History!T3O_UTME = OldLocation!T3O_UTME
        History!T3O_UTMN = OldLocation!T3O_UTMN
        History!T3O_Rebar = OldLocation!T3O_Rebar
        History!T3E_UTME = OldLocation!T3E_UTME
        History!T3E_UTMN = OldLocation!T3E_UTMN
        History!T3E_Rebar = OldLocation!T3E_Rebar
        History!T3_Elevation = OldLocation!T3_Elevation
        History!Bearing_A = OldLocation!Bearing_A   ' Fuels bearings and slopes
        History!Bearing_B = OldLocation!Bearing_B
        History!Bearing_C = OldLocation!Bearing_C
        History!Bearing_D = OldLocation!Bearing_D
        History!Slope_A = OldLocation!Slope_A
        History!Slope_B = OldLocation!Slope_B
        History!Slope_C = OldLocation!Slope_C
        History!Slope_D = OldLocation!Slope_D
        History!SlopeA = OldLocation!SlopeA         ' Plot side slopes
        History!SlopeAUD = OldLocation!SlopeAUD
        History!SlopeB = OldLocation!SlopeB
        History!SlopeBUD = OldLocation!SlopeBUD
        History!SlopeC = OldLocation!SlopeC
        History!SlopeCUD = OldLocation!SlopeCUD
        History!SlopeD = OldLocation!SlopeD
        History!SlopeDUD = OldLocation!SlopeDUD
        History!Plot_Directions = OldLocation!Plot_Directions
        History.Update
        History.Close
        OldLocation.Close
    End If
    
Exit_Form_BeforeUpdate:
  Exit Sub
  
Err_Handler:
  MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
         "Error encountered (Update Location)"
  Resume Exit_Form_BeforeUpdate
End Sub
