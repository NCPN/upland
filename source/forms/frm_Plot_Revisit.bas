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
    FilterOn = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =127
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10800
    DatasheetFontHeight =9
    ItemSuffix =53
    Left =5988
    Top =2064
    Right =16788
    Bottom =11172
    DatasheetGridlinesColor =12632256
    Filter ="[Location_ID]='{0A0952FE-C4D0-49C5-B3DA-72D48B7EFC2B}'"
    RecSrcDt = Begin
        0x9becc7edac0fe340
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
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin Section
            CanGrow = NotDefault
            Height =9540
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =3225
                    Top =120
                    Width =3795
                    Height =480
                    FontSize =18
                    FontWeight =700
                    Name ="Label0"
                    Caption ="Plot Revisit"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =720
                    Top =360
                    Width =540
                    TabIndex =9
                    Name ="Unit_Code"
                    ControlSource ="Unit_Code"
                    StatusBarText ="Park Code."

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =180
                            Top =360
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
                    Top =360
                    Width =600
                    TabIndex =10
                    Name ="Plot_ID"
                    ControlSource ="Plot_ID"
                    StatusBarText ="Plot identifier"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =1560
                            Top =360
                            Width =660
                            Height =240
                            FontWeight =700
                            Name ="Label2"
                            Caption ="Plot ID"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =87
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1560
                    Top =2520
                    Width =1020
                    TabIndex =11
                    Name ="Plot_E_Coord"
                    ControlSource ="E_Coord"
                    StatusBarText ="UTM East of Centroid if plot established on different visit from site characteri"
                        "zation"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =180
                            Top =2520
                            Width =1380
                            Height =240
                            FontWeight =700
                            Name ="Label4"
                            Caption ="Centroid UTM E"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3540
                    Top =2520
                    Width =1019
                    TabIndex =12
                    Name ="Plot_N_Coord"
                    ControlSource ="N_Coord"
                    StatusBarText ="UTM North of Centroid (Y_Coord) if plot established on different visit from site"
                        " characterization"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =2880
                            Top =2520
                            Width =600
                            Height =240
                            FontWeight =700
                            Name ="Label5"
                            Caption ="UTM N"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5520
                    Top =2520
                    Width =480
                    TabIndex =13
                    Name ="Plot_Slope"
                    ControlSource ="Plot_Slope"
                    StatusBarText ="Plot slope in percent - 1 decimal"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =4860
                            Top =2520
                            Width =600
                            Height =240
                            FontWeight =700
                            Name ="Label7"
                            Caption ="Slope"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7020
                    Top =2520
                    Width =480
                    TabIndex =14
                    Name ="Plot_Aspect"
                    ControlSource ="Plot_Aspect"
                    StatusBarText ="Plot aspect in degrees - 1 decimal"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =6300
                            Top =2520
                            Width =660
                            Height =240
                            FontWeight =700
                            Name ="Label8"
                            Caption ="Aspect"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8700
                    Top =2520
                    Width =480
                    TabIndex =15
                    Name ="Azimuth"
                    ControlSource ="Azimuth"
                    StatusBarText ="Direction from origin to end of center transect in degrees - 1 decimal"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =7800
                            Top =2520
                            Width =840
                            Height =240
                            FontWeight =700
                            Name ="Label9"
                            Caption ="Azimuth"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2760
                    Top =3600
                    Width =1080
                    TabIndex =16
                    Name ="T1O_UTME"
                    ControlSource ="T1O_UTME"
                    StatusBarText ="UTM East of Transect 1 origin"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =1860
                            Top =3600
                            Width =900
                            Height =240
                            FontWeight =700
                            Name ="Label10"
                            Caption ="1"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3840
                    Top =3600
                    Width =1080
                    TabIndex =17
                    Name ="T1O_UTMN"
                    ControlSource ="T1O_UTMN"
                    StatusBarText ="UTM North of Transect 1 origin"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =2760
                            Top =3180
                            Width =1080
                            Height =420
                            FontWeight =700
                            Name ="Label11"
                            Caption ="UTM E"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =1
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4920
                    Top =3600
                    Width =960
                    TabIndex =18
                    Name ="T1O_Rebar"
                    ControlSource ="T1O_Rebar"
                    StatusBarText ="Distance from origin of rebar in meters - 1 decimal"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =4920
                            Top =2940
                            Width =960
                            Height =660
                            FontWeight =700
                            Name ="Label12"
                            Caption ="Rebar location (m)"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5880
                    Top =3600
                    Width =1080
                    TabIndex =19
                    Name ="T1E_UTME"
                    ControlSource ="T1E_UTME"
                    StatusBarText ="UTM East of Transect 1 end"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =5880
                            Top =3180
                            Width =1080
                            Height =420
                            FontWeight =700
                            Name ="Label13"
                            Caption ="UTM E"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6960
                    Top =3600
                    Width =1080
                    TabIndex =20
                    Name ="T1E_UTMN"
                    ControlSource ="T1E_UTMN"
                    StatusBarText ="UTM North of Transect 1 end"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =6960
                            Top =3180
                            Width =1080
                            Height =420
                            FontWeight =700
                            Name ="Label14"
                            Caption ="UTM N"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =1
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8040
                    Top =3600
                    Width =959
                    TabIndex =21
                    Name ="T1E_Rebar"
                    ControlSource ="T1E_Rebar"
                    StatusBarText ="Distance from end of rebar in meters - 1 decimal"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =8040
                            Top =2940
                            Width =959
                            Height =660
                            FontWeight =700
                            Name ="Label15"
                            Caption ="Rebar location (m)"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2760
                    Top =3840
                    Width =1080
                    TabIndex =23
                    Name ="T2O_UTME"
                    ControlSource ="T2O_UTME"
                    StatusBarText ="UTM East of Transect 2 origin"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =1860
                            Top =3840
                            Width =900
                            Height =240
                            FontWeight =700
                            Name ="Label16"
                            Caption ="2"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3840
                    Top =3840
                    Width =1080
                    TabIndex =24
                    Name ="T2O_UTMN"
                    ControlSource ="T2O_UTMN"
                    StatusBarText ="UTM North of Transect 2 origin"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =1
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4920
                    Top =3840
                    Width =960
                    TabIndex =25
                    Name ="T2O_Rebar"
                    ControlSource ="T2O_Rebar"
                    StatusBarText ="Distance from origin of rebar in meters - 1 decimal"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5880
                    Top =3840
                    Width =1080
                    TabIndex =26
                    Name ="T2E_UTME"
                    ControlSource ="T2E_UTME"
                    StatusBarText ="UTM East of Transect 2 end"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6960
                    Top =3840
                    Width =1080
                    TabIndex =27
                    Name ="T2E_UTMN"
                    ControlSource ="T2E_UTMN"
                    StatusBarText ="UTM North of Transect 2 end"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =1
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8040
                    Top =3840
                    Width =960
                    TabIndex =28
                    Name ="T2E_Rebar"
                    ControlSource ="T2E_Rebar"
                    StatusBarText ="Distance from end of rebar in meters - 1 decimal"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2760
                    Top =4080
                    Width =1080
                    TabIndex =30
                    Name ="T3O_UTME"
                    ControlSource ="T3O_UTME"
                    StatusBarText ="UTM East of Transect 3 origin"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =1860
                            Top =4080
                            Width =900
                            Height =240
                            FontWeight =700
                            Name ="Label22"
                            Caption ="3"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3840
                    Top =4080
                    Width =1080
                    TabIndex =31
                    Name ="T3O_UTMN"
                    ControlSource ="T3O_UTMN"
                    StatusBarText ="UTM North of Transect 3 origin"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =1
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4920
                    Top =4080
                    Width =960
                    TabIndex =32
                    Name ="T3O_Rebar"
                    ControlSource ="T3O_Rebar"
                    StatusBarText ="Distance from origin of rebar in meters - 1 decimal"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5880
                    Top =4080
                    Width =1080
                    TabIndex =33
                    Name ="T3E_UTME"
                    ControlSource ="T3E_UTME"
                    StatusBarText ="UTM East of Transect 3 end"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6960
                    Top =4080
                    Width =1080
                    TabIndex =34
                    Name ="T3E_UTMN"
                    ControlSource ="T3E_UTMN"
                    StatusBarText ="UTM North of Transect 3 end"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =1
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8040
                    Top =4080
                    Width =960
                    TabIndex =35
                    Name ="T3E_Rebar"
                    ControlSource ="T3E_Rebar"
                    StatusBarText ="Distance from end of rebar in meters - 1 decimal"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    SpecialEffect =0
                    OverlapFlags =87
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1800
                    Top =4500
                    Width =6300
                    Height =603
                    TabIndex =37
                    Name ="Plot_Directions"
                    ControlSource ="Plot_Directions"
                    StatusBarText ="Directions to plot"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =180
                            Top =4500
                            Width =1620
                            Height =240
                            FontWeight =700
                            Name ="Label28"
                            Caption ="Directions To Plot"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9000
                    Top =3600
                    Width =1080
                    TabIndex =22
                    Name ="T1_Elevation"
                    ControlSource ="T1_Elevation"
                    StatusBarText ="Elevation in meters of transect 1"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =9000
                            Top =2940
                            Width =1080
                            Height =660
                            FontWeight =700
                            Name ="Label29"
                            Caption ="Elevation (m)"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9000
                    Top =3840
                    Width =1080
                    TabIndex =29
                    Name ="T2_Elevation"
                    ControlSource ="T2_Elevation"
                    StatusBarText ="Elevation in meters of transect 2"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9000
                    Top =4080
                    Width =1080
                    TabIndex =36
                    Name ="T3_Elevation"
                    ControlSource ="T3_Elevation"
                    StatusBarText ="Elevation in meters of transect 3"

                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =1860
                    Top =2940
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
                    Left =3840
                    Top =3180
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
                    Left =2760
                    Top =2940
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
                    Left =5880
                    Top =2940
                    Width =2160
                    Height =240
                    FontWeight =700
                    Name ="Label37"
                    Caption ="Rebar at transect end"
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    Left =840
                    Top =2940
                    Width =1020
                    Height =540
                    FontSize =10
                    FontWeight =700
                    Name ="Label38"
                    Caption ="Transect Location"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2520
                    Top =8880
                    Width =1455
                    Height =300
                    TabIndex =47
                    Name ="ButtonClose"
                    Caption ="Cancel New Visit"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =180
                    Top =120
                    Width =720
                    TabIndex =48
                    Name ="Location_ID"
                    ControlSource ="Location_ID"
                    StatusBarText ="M. Location identifier (Loc_ID)"

                End
                Begin Subform
                    OverlapFlags =85
                    Left =1380
                    Top =780
                    Width =7410
                    Height =1560
                    Name ="fsub_Revisit"
                    SourceObject ="Form.fsub_Revisit"
                    LinkChildFields ="Location_ID"
                    LinkMasterFields ="Location_ID"

                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6480
                    Top =8880
                    Width =1454
                    Height =299
                    TabIndex =49
                    Name ="ButtonContinue"
                    Caption ="Continue"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8640
                    Top =4560
                    Width =1380
                    Height =480
                    TabIndex =50
                    ForeColor =255
                    Name ="ButtonModify"
                    Caption ="Modify Location Coordinates"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin Subform
                    OverlapFlags =87
                    Left =180
                    Top =5880
                    Width =10365
                    Height =1305
                    TabIndex =38
                    Name ="fsub_FW_Monument"
                    SourceObject ="Form.fsub_FW_Monument"
                    LinkChildFields ="Location_ID"
                    LinkMasterFields ="Location_ID"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =180
                            Top =5640
                            Width =1560
                            Height =240
                            FontWeight =700
                            Name ="fsub_FW_Monument Label"
                            Caption ="Monument Trees"
                            EventProcPrefix ="fsub_FW_Monument_Label"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =2880
                    Top =5280
                    Width =480
                    TabIndex =1
                    Name ="SlopeA"
                    ControlSource ="SlopeA"
                    StatusBarText ="F/W Slope for plot side A in percent"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2640
                            Top =5280
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
                    Top =5280
                    Width =479
                    TabIndex =3
                    Name ="SlopeB"
                    ControlSource ="SlopeB"
                    StatusBarText ="F/W Slope for plot side B in percent"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4500
                            Top =5280
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
                    Top =5280
                    Width =479
                    TabIndex =5
                    Name ="SlopeC"
                    ControlSource ="SlopeC"
                    StatusBarText ="F/W Slope for plot side C in percent"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6420
                            Top =5280
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
                    Top =5280
                    Width =479
                    TabIndex =7
                    Name ="SlopeD"
                    ControlSource ="SlopeD"
                    StatusBarText ="F/W Slope for plot side D in percent"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8340
                            Top =5280
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
                    Top =5280
                    Width =780
                    TabIndex =2
                    Name ="SlopeAUD"
                    ControlSource ="SlopeAUD"
                    RowSourceType ="Value List"
                    RowSource ="\"up\";\"down\""
                    ColumnWidths ="570"

                End
                Begin Label
                    OverlapFlags =85
                    Left =180
                    Top =5280
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
                    Top =5280
                    Width =780
                    TabIndex =4
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
                    Top =5280
                    Width =780
                    TabIndex =6
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
                    Top =5280
                    Width =780
                    TabIndex =8
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
                    Left =4305
                    Top =8040
                    Width =810
                    Height =300
                    TabIndex =39
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
                            Left =4305
                            Top =7800
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
                    Left =5100
                    Top =8040
                    Width =810
                    Height =300
                    TabIndex =40
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
                            Left =5100
                            Top =7800
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
                    Left =5925
                    Top =8040
                    Width =810
                    Height =300
                    TabIndex =41
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
                            Left =5925
                            Top =7800
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
                    Left =6765
                    Top =8040
                    Width =810
                    Height =300
                    TabIndex =42
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
                            Left =6765
                            Top =7800
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
                    Left =4305
                    Top =8340
                    Width =810
                    Height =300
                    TabIndex =43
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
                            Left =3000
                            Top =8040
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
                    Left =5100
                    Top =8340
                    Width =810
                    Height =300
                    TabIndex =44
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
                            Left =3000
                            Top =8340
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
                    Left =5925
                    Top =8340
                    Width =810
                    Height =300
                    TabIndex =45
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
                    Left =6765
                    Top =8340
                    Width =810
                    Height =300
                    TabIndex =46
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Slope_D"
                    ControlSource ="Slope_D"
                    StatusBarText ="Slope of transect C to nearest half percent"

                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =4260
                    Top =7440
                    Width =1860
                    Height =300
                    FontSize =10
                    FontWeight =700
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Fuels_Transect_Label"
                    Caption ="Fuels Transect"
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

Private Sub ButtonClose_Click()
On Error GoTo Err_ButtonClose_Click

    Dim stDocName As String
    Dim db As Database
    Dim Events As DAO.Recordset
    Dim strSQL As String
  If Not IsNull(Me!fsub_Revisit.Form!Event_ID) Then
    strSQL = "Select * FROM tbl_events WHERE event_ID = '" & Me!fsub_Revisit.Form!Event_ID & "'"
    Set db = CurrentDb
  ' Get the added events record
    Set Events = db.OpenRecordset(strSQL)
    If Not Events.EOF Then
      Events.Delete
'      Delete cancelled events record
      Events.Close
    End If
  End If
    DoCmd.Close

Exit_ButtonClose_Click:
    Exit Sub

Err_ButtonClose_Click:
    MsgBox Err.Description
    Resume Exit_ButtonClose_Click
    
End Sub

' ---------------------------------
' SUB:          Form_BeforeUpdate
' Description:  Populate centroid UTMs from tbl_Location_History
' Assumptions:  -
' Parameters:   Cancel - species to check (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Russ DenBleyker, date unkown, Northern Colorado Plateau Network
' Adapted:      -
' Revisions:
'   RD  - ?         - initial version
'   BLC - 8/11/2015 - fixed bug improperly populating plot centroid UTMs with
'                     tbl_Location_History deprecated Plot_E_Coord & Plot_N_Coord vs.
'                     E_Coord & N_Coord values, updated error handling & added documentation
'   BLC - 8/19/2015 - fixed bug improperly populating recorder (Me!fsub_Revisit.Form!Observer vs. Me!Recorder)
' ---------------------------------
Private Sub Form_BeforeUpdate(Cancel As Integer)
    On Error GoTo Err_Handler
    Dim db As Database
    Dim History As DAO.Recordset
    Dim OldLocation As DAO.Recordset
    Dim strSQL As String
    
    If IsNull(Me!fsub_Revisit.Form!Observer) Then
      MsgBox "You must select an observer name!"
      Me.Undo
      GoTo Exit_Sub
    Else
      Set db = CurrentDb
      strSQL = "Select * from tbl_Locations WHERE Location_ID = '" & Me!Location_ID & "'"
      Set OldLocation = db.OpenRecordset(strSQL)  '  Get unmodified location record
      If OldLocation.EOF Then
        MsgBox "Location record not found."
        GoTo Exit_Sub
      Else
        OldLocation.MoveFirst
      End If
      Set History = db.OpenRecordset("tbl_Location_History")
        History.AddNew                     ' Create a Location History record
        History!Location_History_ID = fxnGUIDGen
        History!Location_ID = Me!Location_ID
        History!Modify_Date = Now()        ' Date of update
        
        'populate individual committing update
        History!Recorder = Me!fsub_Revisit.Form!Observer
        'History!Recorder = Me!Recorder     ' Person committing update
        
        History!Unit_Code = Me!Unit_Code
        History!Plot_ID = Me!Plot_ID
        
        'populate plot centroid UTMs E & N Coord
        History!E_Coord = OldLocation!E_Coord
        History!N_Coord = OldLocation!N_Coord
        
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
        History!Plot_Directions = OldLocation!Plot_Directions
        
        ' Fuels bearings and slopes
        History!Bearing_A = OldLocation!Bearing_A
        History!Bearing_B = OldLocation!Bearing_B
        History!Bearing_C = OldLocation!Bearing_C
        History!Bearing_D = OldLocation!Bearing_D
        History!Slope_A = OldLocation!Slope_A
        History!Slope_B = OldLocation!Slope_B
        History!Slope_C = OldLocation!Slope_C
        History!Slope_D = OldLocation!Slope_D
        
        ' Plot side slopes
        History!SlopeA = OldLocation!SlopeA
        History!SlopeAUD = OldLocation!SlopeAUD
        History!SlopeB = OldLocation!SlopeB
        History!SlopeBUD = OldLocation!SlopeBUD
        History!SlopeC = OldLocation!SlopeC
        History!SlopeCUD = OldLocation!SlopeCUD
        History!SlopeD = OldLocation!SlopeD
        History!SlopeDUD = OldLocation!SlopeDUD
        History.Update
        History.Close
        OldLocation.Close
    End If
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - BeforeUpdate[Form_frm_Plot_Revisit])"
    End Select
    Resume Exit_Sub
End Sub

Private Sub Form_Load()
  Dim Veg_Type As Variant
  
    DoCmd.Restore
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
'  Forms![frm_Plot_Revisit]![fsub_FW_Monument].Form.AllowEdits = False
End Sub
Private Sub ButtonContinue_Click()
On Error GoTo Err_ButtonContinue_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    
    If IsNull(Me!fsub_Revisit.Form!Start_Date) Then
      MsgBox "Revisit date required"
      GoTo Exit_ButtonContinue_Click
    Else
    DoCmd.RunCommand acCmdSaveRecord  ' Save the new event record
    End If
    stDocName = "frm_Data_Entry"
    stLinkCriteria = "[Event_ID]=" & "'" & Me!fsub_Revisit.Form!Event_ID & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
        If Not IsNull(Me!Location_ID) Then
            ' Fill in Location
            Forms!frm_Data_Entry!cboLocation_ID = Me!Location_ID
            Forms!frm_Data_Entry.Update_Loc_Info
        End If
    DoCmd.Close acForm, "frm_Plot_Revisit"
Exit_ButtonContinue_Click:
    Exit Sub

Err_ButtonContinue_Click:
    MsgBox Err.Description
    Resume Exit_ButtonContinue_Click
    
End Sub

Private Sub ButtonModify_Click()
On Error GoTo Err_ButtonModify_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Location_Modify"
    
    stLinkCriteria = "[Location_ID]=" & "'" & Me![Location_ID] & "'"
    ' Check to see if modifications were made and, if so, save event ID.
    If (IsNull(Me!fsub_Revisit.Form!Start_Date) + IsNull(Me!fsub_Revisit.Form!Observer) + IsNull(Me!fsub_Revisit.Form!Comments)) > -3 Then
      Me!fsub_Revisit.Form!Event_Save = Me!fsub_Revisit.Form!Event_ID
    End If
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , acDialog
    Me.Requery
Exit_ButtonModify_Click:
    Exit Sub

Err_ButtonModify_Click:
    MsgBox Err.Description
    Resume Exit_ButtonModify_Click
    
End Sub
