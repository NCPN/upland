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
    ItemSuffix =52
    Left =4110
    Top =2910
    Right =14190
    Bottom =10095
    DatasheetGridlinesColor =12632256
    Filter ="[Location_ID]='{F4CE3EAB-E640-4E3F-8343-008314E75F39}'"
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
    OnLoad ="[Event Procedure]"
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
                    ColumnHeads = NotDefault
                    OverlapFlags =85
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
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Person making changes (required)"

                    LayoutCachedLeft =7440
                    LayoutCachedTop =780
                    LayoutCachedWidth =9540
                    LayoutCachedHeight =1020
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =5820
                            Top =780
                            Width =1560
                            Height =228
                            FontWeight =700
                            Name ="Observer_Label"
                            Caption ="Changes made by:"
                            ControlTipText ="Person making changes (required)"
                            LayoutCachedLeft =5820
                            LayoutCachedTop =780
                            LayoutCachedWidth =7380
                            LayoutCachedHeight =1008
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
                    Left =5220
                    Top =6540
                    Width =1575
                    Height =330
                    TabIndex =31
                    Name ="btnClose"
                    Caption ="Exit Without Saving"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =5220
                    LayoutCachedTop =6540
                    LayoutCachedWidth =6795
                    LayoutCachedHeight =6870
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
                Begin CommandButton
                    OverlapFlags =85
                    Left =3600
                    Top =6540
                    Width =1335
                    Height =330
                    TabIndex =41
                    Name ="btnUpdate"
                    Caption ="Update Location"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =3600
                    LayoutCachedTop =6540
                    LayoutCachedWidth =4935
                    LayoutCachedHeight =6870
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7380
                    Top =1260
                    Width =1320
                    TabIndex =42
                    Name ="tbxRecorderID"

                    LayoutCachedLeft =7380
                    LayoutCachedTop =1260
                    LayoutCachedWidth =8700
                    LayoutCachedHeight =1500
                End
                Begin Label
                    OverlapFlags =215
                    Left =5640
                    Top =720
                    Width =240
                    Height =240
                    FontSize =12
                    ForeColor =2366701
                    Name ="lblReqd"
                    Caption ="*"
                    ControlTipText ="Reqiured field"
                    LayoutCachedLeft =5640
                    LayoutCachedTop =720
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =960
                End
                Begin Label
                    OverlapFlags =93
                    Left =8280
                    Top =120
                    Width =240
                    Height =240
                    FontSize =12
                    ForeColor =2366701
                    Name ="lblKeyReqd"
                    Caption ="*"
                    ControlTipText ="Reqiured field"
                    LayoutCachedLeft =8280
                    LayoutCachedTop =120
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =360
                End
                Begin Label
                    OverlapFlags =87
                    Left =8520
                    Top =120
                    Width =1392
                    Height =252
                    FontSize =9
                    Name ="lblKeyReqdField"
                    Caption ="= Required Field"
                    ControlTipText ="Reqiured field"
                    LayoutCachedLeft =8520
                    LayoutCachedTop =120
                    LayoutCachedWidth =9912
                    LayoutCachedHeight =372
                    ForeThemeColorIndex =0
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
' MODULE:       frm_Location_Modify
' Level:        Form module
' Version:      1.02
' Description:  data functions & procedures specific to location modifications
'
' Source/date:  John R. Boetsch, June 2006
' Adapted:      Bonnie Campbell, 2/4/2016
' Revisions:    RDB - unknown  - 1.00 - initial version
'               BLC - 2/4/2016 - 1.01 - added documentation, adjusted form to require
'                                       recorder before save button enabled
'               BLC - 3/7/2016 - 1.02 - fix so person making changes is *not* the recorder
'                                       recorder is person who does site characterization (one-time event)
'                                       this needs to be whomever made changes (multiple time event)
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
' Adapted:      Bonnie Campbell, February 4, 2016 - for NCPN tools
' Revisions:
'   BLC, 2/4/2016  - initial version
'   BLC, 3/7/2016  - changed tbxRecorderID from pulling [Recorder] from site characterization
'                    to TempVars.item("User_ID") which is set from frm_Set_Defaults
' ---------------------------------
Private Sub Form_Load()
    On Error GoTo Err_Handler
       
    'set the default value for tbxRecorderID
    'value must be set here not in control data source as =[TempVars].[Item]("User_ID")
    'otherwise tbxRecorderID_AfterUpdate() will fail to update with
    'Error #2448 You can't assign a value to this object.
    tbxRecorderID = TempVars.item("User_ID")
    
    'set the value based on tbl_Locations.Recorder for this record
    '-----------------------------------------------------------------------
    ' NOTE:
    '   tbl_Locations.Recorder is NOT the recorder for site characterization
    '   it is the person making changes on this form which usually is the
    '   user selected when beginning the enter/edit data (frm_Set_Defaults)
    '   determine who this is via TempVars.item("User_ID")
    '-----------------------------------------------------------------------
    If Not IsNull(Me.tbxRecorderID.Value) Then
        Me.Recorder.Value = Me.tbxRecorderID.Value
    End If
       
    'enable the update button only if the recorder is entered
    btnUpdate.Enabled = False
    If Not IsNull(Me.Recorder) Then
        btnUpdate.Enabled = True
        Me.Recorder.backcolor = RGB(255, 255, 255) 'set background to white
    Else
        Me.Recorder.SetFocus
        Me.Recorder.backcolor = RGB(255, 255, 51) 'set background to yellow
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[Form_frm_Location_Modify])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_BeforeUpdate
' Description:  Handles form pre-update actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 4, 2016 - for NCPN tools
' Revisions:
'   JRB, 6/x/2006  - initial version
'   RDB, unknown   - ?
'   BLC, 2/4/2016  - added documentation, removed undo on recorder name reminder
' ---------------------------------
Private Sub Form_BeforeUpdate(Cancel As Integer)
    On Error GoTo Err_Handler
'    Dim strMsg As String
'    Dim db As Database
'    Dim History As DAO.Recordset
'    Dim OldLocation As DAO.Recordset
'    Dim strSQL As String
'
'    strMsg = "Are you sure you want to update location coordinates?"
'    strMsg = strMsg & chr(13) & chr(10) & "Click Yes to Save or No to Discard changes."
'
'    If MsgBox(strMsg, vbQuestion + vbYesNo, "Update Location?") = vbNo Then
'    '---------------- Cancel Location Update -----------
'      Me.Undo
'      GoTo Exit_Handler
'
'    ElseIf IsNull(Me!Recorder) Then
'    '---------------- Require Recorder -----------------
'      MsgBox "You must select a recorder name!"
'      Me.Recorder.SetFocus
'      Me.Recorder.BackColor = RGB(255, 255, 51) 'set background to yellow
'      'Me.Undo
'      'GoTo Exit_Handler
'
'    Else
'    '---------------- Process Changes ------------------
'      Set db = CurrentDb
'      strSQL = "Select * from tbl_Locations WHERE Location_ID = '" & Me!Location_ID & "'"
'
'      Set OldLocation = db.OpenRecordset(strSQL)  '  Get unmodified location record
'      If OldLocation.EOF Then
'        MsgBox "Location record not found."
'        GoTo Exit_Handler
'      Else
'        OldLocation.MoveFirst
'      End If
'
'      Set History = db.OpenRecordset("tbl_Location_History")
'      With History
'        .AddNew                     ' Create a Location History record
'        !Location_History_ID = fxnGUIDGen
'        !Location_ID = Me!Location_ID
'        !Modify_Date = Now()        ' Date of update
'        !Recorder = Me!Recorder     ' Person committing update
'        !Unit_Code = Me!Unit_Code
'        !Plot_ID = Me!Plot_ID
'        ' Modified to populate plot centroid coordinates (E_Coord, N_Coord) correctly. [HMT, 3/16/2015]
'        ' Plot_E_Coord, Plot_N_Coord are no longer used.
'        ' !E_Coord = OldLocation!Plot_E_Coord
'        ' !N_Coord = OldLocation!Plot_N_Coord
'        !E_Coord = OldLocation!E_Coord            ' UTM easting of plot centroid
'        !N_Coord = OldLocation!N_Coord            ' UTM northing of plot centroid
'        !Plot_Slope = OldLocation!Plot_Slope
'        !Plot_Aspect = OldLocation!Plot_Aspect
'        !Azimuth = OldLocation!Azimuth
'        !T1O_UTME = OldLocation!T1O_UTME
'        !T1O_UTMN = OldLocation!T1O_UTMN
'        !T1O_Rebar = OldLocation!T1O_Rebar
'        !T1E_UTME = OldLocation!T1E_UTME
'        !T1E_UTMN = OldLocation!T1E_UTMN
'        !T1E_Rebar = OldLocation!T1E_Rebar
'        !T1_Elevation = OldLocation!T1_Elevation
'        !T2O_UTME = OldLocation!T2O_UTME
'        !T2O_UTMN = OldLocation!T2O_UTMN
'        !T2O_Rebar = OldLocation!T2O_Rebar
'        !T2E_UTME = OldLocation!T2E_UTME
'        !T2E_UTMN = OldLocation!T2E_UTMN
'        !T2E_Rebar = OldLocation!T2E_Rebar
'        !T2_Elevation = OldLocation!T2_Elevation
'        !T3O_UTME = OldLocation!T3O_UTME
'        !T3O_UTMN = OldLocation!T3O_UTMN
'        !T3O_Rebar = OldLocation!T3O_Rebar
'        !T3E_UTME = OldLocation!T3E_UTME
'        !T3E_UTMN = OldLocation!T3E_UTMN
'        !T3E_Rebar = OldLocation!T3E_Rebar
'        !T3_Elevation = OldLocation!T3_Elevation
'        !Bearing_A = OldLocation!Bearing_A   ' Fuels bearings and slopes
'        !Bearing_B = OldLocation!Bearing_B
'        !Bearing_C = OldLocation!Bearing_C
'        !Bearing_D = OldLocation!Bearing_D
'        !Slope_A = OldLocation!Slope_A
'        !Slope_B = OldLocation!Slope_B
'        !Slope_C = OldLocation!Slope_C
'        !Slope_D = OldLocation!Slope_D
'        !SlopeA = OldLocation!SlopeA         ' Plot side slopes
'        !SlopeAUD = OldLocation!SlopeAUD
'        !SlopeB = OldLocation!SlopeB
'        !SlopeBUD = OldLocation!SlopeBUD
'        !SlopeC = OldLocation!SlopeC
'        !SlopeCUD = OldLocation!SlopeCUD
'        !SlopeD = OldLocation!SlopeD
'        !SlopeDUD = OldLocation!SlopeDUD
'        !Plot_Directions = OldLocation!Plot_Directions
'        .Update
'        .Close
'        End With
'        OldLocation.Close
'    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeUpdate[Form_frm_Location_Modify])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnUpdate_Click
' Description:  Handles form update button actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 4, 2016 - for NCPN tools
' Revisions:
'   BLC, 2/4/2016  - initial version
' ---------------------------------
Private Sub btnUpdate_Click()
On Error GoTo Err_Handler

    Dim strMsg As String
    Dim db As Database
    Dim History As DAO.Recordset
    Dim OldLocation As DAO.Recordset
    Dim strSQL As String
    
    strMsg = "Are you sure you want to update location coordinates?"
    strMsg = strMsg & chr(13) & chr(10) & "Click Yes to Save or No to Discard changes."
    
    If MsgBox(strMsg, vbQuestion + vbYesNo, "Update Location?") = vbNo Then
    '---------------- Cancel Location Update -----------
      Me.Undo
      GoTo Exit_Handler
    
'    ElseIf IsNull(Me!Recorder) Then
'    '---------------- Require Recorder -----------------
'      MsgBox "You must select a recorder name!"
'      Me.Recorder.SetFocus
'      Me.Recorder.BackColor = RGB(255, 255, 51) 'set background to yellow
'      'Me.Undo
'      'GoTo Exit_Handler
    
    Else
    '---------------- Process Changes ------------------
      Set db = CurrentDb
      strSQL = "Select * from tbl_Locations WHERE Location_ID = '" & Me!Location_ID & "'"
      
      Set OldLocation = db.OpenRecordset(strSQL)  '  Get unmodified location record
      If OldLocation.EOF Then
        MsgBox "Location record not found."
        GoTo Exit_Handler
      Else
        OldLocation.MoveFirst
      End If
      
      Set History = db.OpenRecordset("tbl_Location_History")
      With History
        .AddNew                     ' Create a Location History record
        !Location_History_ID = fxnGUIDGen
        !Location_ID = Me!Location_ID
        !Modify_Date = Now()        ' Date of update
        !Recorder = Me!Recorder     ' Person committing update
        !Unit_Code = Me!Unit_Code
        !Plot_ID = Me!Plot_ID
        ' Modified to populate plot centroid coordinates (E_Coord, N_Coord) correctly. [HMT, 3/16/2015]
        ' Plot_E_Coord, Plot_N_Coord are no longer used.
        ' !E_Coord = OldLocation!Plot_E_Coord
        ' !N_Coord = OldLocation!Plot_N_Coord
        !E_Coord = OldLocation!E_Coord            ' UTM easting of plot centroid
        !N_Coord = OldLocation!N_Coord            ' UTM northing of plot centroid
        !Plot_Slope = OldLocation!Plot_Slope
        !Plot_Aspect = OldLocation!Plot_Aspect
        !Azimuth = OldLocation!Azimuth
        !T1O_UTME = OldLocation!T1O_UTME
        !T1O_UTMN = OldLocation!T1O_UTMN
        !T1O_Rebar = OldLocation!T1O_Rebar
        !T1E_UTME = OldLocation!T1E_UTME
        !T1E_UTMN = OldLocation!T1E_UTMN
        !T1E_Rebar = OldLocation!T1E_Rebar
        !T1_Elevation = OldLocation!T1_Elevation
        !T2O_UTME = OldLocation!T2O_UTME
        !T2O_UTMN = OldLocation!T2O_UTMN
        !T2O_Rebar = OldLocation!T2O_Rebar
        !T2E_UTME = OldLocation!T2E_UTME
        !T2E_UTMN = OldLocation!T2E_UTMN
        !T2E_Rebar = OldLocation!T2E_Rebar
        !T2_Elevation = OldLocation!T2_Elevation
        !T3O_UTME = OldLocation!T3O_UTME
        !T3O_UTMN = OldLocation!T3O_UTMN
        !T3O_Rebar = OldLocation!T3O_Rebar
        !T3E_UTME = OldLocation!T3E_UTME
        !T3E_UTMN = OldLocation!T3E_UTMN
        !T3E_Rebar = OldLocation!T3E_Rebar
        !T3_Elevation = OldLocation!T3_Elevation
        !Bearing_A = OldLocation!Bearing_A   ' Fuels bearings and slopes
        !Bearing_B = OldLocation!Bearing_B
        !Bearing_C = OldLocation!Bearing_C
        !Bearing_D = OldLocation!Bearing_D
        !Slope_A = OldLocation!Slope_A
        !Slope_B = OldLocation!Slope_B
        !Slope_C = OldLocation!Slope_C
        !Slope_D = OldLocation!Slope_D
        !SlopeA = OldLocation!SlopeA         ' Plot side slopes
        !SlopeAUD = OldLocation!SlopeAUD
        !SlopeB = OldLocation!SlopeB
        !SlopeBUD = OldLocation!SlopeBUD
        !SlopeC = OldLocation!SlopeC
        !SlopeCUD = OldLocation!SlopeCUD
        !SlopeD = OldLocation!SlopeD
        !SlopeDUD = OldLocation!SlopeDUD
        !Plot_Directions = OldLocation!Plot_Directions
        .Update
        .Close
        End With
        OldLocation.Close
    End If

Exit_Handler:
    'close form
    DoCmd.Close
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnUpdate_Click[Form_frm_Location_Modify])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Recorder_AfterUpdate
' Description:  Handles form post-update actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 4, 2016 - for NCPN tools
' Revisions:
'   BLC, 2/4/2016  - initial version
' ---------------------------------
Private Sub Recorder_AfterUpdate()
On Error GoTo Err_Handler
    
    'set the tbxRecorderID to update the record
    Me.tbxRecorderID = Me.Recorder.Value

    'enable save when recorder isn't null
    If Not IsNull(Me!Recorder) Then
        Me.btnUpdate.Enabled = True
        Me.Recorder.backcolor = RGB(255, 255, 255) 'set background to white
    Else
        Me.btnUpdate.Enabled = False
        Me.Recorder.SetFocus
        Me.Recorder.backcolor = RGB(255, 255, 51) 'set background to yellow
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Recorder_AfterUpdate[Form_frm_Location_Modify])"
    End Select
    Resume Exit_Handler
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

' ---------------------------------
' SUB:          btnClose_Click
' Description:  Handles form button close actions (exit without saving)
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 4, 2016 - for NCPN tools
' Revisions:
'   JRB, 6/x/2006  - initial version
'   RDB, unknown   - ?
'   BLC, 2/4/2016  - added documentation, revised name to btnClose vs. ButtonClose
' ---------------------------------
Private Sub btnClose_Click()
On Error GoTo Err_Handler

    Me.Undo
    DoCmd.Close

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnClose_Click[Form_frm_Location_Modify])"
    End Select
    Resume Exit_Handler
End Sub
