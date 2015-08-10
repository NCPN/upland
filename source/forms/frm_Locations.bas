Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =13080
    DatasheetFontHeight =9
    ItemSuffix =127
    Left =4860
    Top =2805
    Right =18150
    Bottom =9705
    DatasheetGridlinesColor =12632256
    Filter ="[Location_ID]='{7EA92962-EF5B-4753-A542-6A754C5EEB62}'"
    RecSrcDt = Begin
        0xdca6db037508e340
    End
    RecordSource ="tbl_Locations"
    Caption =" Locations"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    OnClose ="[Event Procedure]"
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
            FontSize =18
        End
        Begin Line
            BorderLineStyle =0
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
            Height =10080
            BackColor =12574431
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =948
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txtLocation_ID"
                    ControlSource ="Location_ID"
                    StatusBarText ="Unique identifier for each sample location"

                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6720
                    Top =1140
                    Width =1080
                    TabIndex =8
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txtX_Coord"
                    ControlSource ="E_Coord"
                    StatusBarText ="M. X coordinate (X_Coord)"
                    Tag ="<data>"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6000
                            Top =1140
                            Width =690
                            Height =240
                            FontSize =8
                            FontWeight =700
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label40"
                            Caption ="UTM E"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8700
                    Top =1140
                    Width =1080
                    TabIndex =9
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txtY_Coord"
                    ControlSource ="N_Coord"
                    StatusBarText ="M. Y coordinate (Y_Coord)"
                    Tag ="<data>"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7980
                            Top =1140
                            Width =690
                            Height =240
                            FontSize =8
                            FontWeight =700
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label41"
                            Caption ="UTM N"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1320
                    Top =9600
                    Width =1800
                    TabIndex =45
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txtUpdated_Date"
                    ControlSource ="Updated_Date"
                    StatusBarText ="MA. Date of entry or last change (Upd_Date)"
                    DefaultValue ="=Now()"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =240
                            Top =9600
                            Width =1080
                            Height =240
                            FontSize =8
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label52"
                            Caption ="Updated Date"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =3150
                    Left =1140
                    Top =720
                    Width =960
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"10\""
                    Name ="Unit_Code"
                    ControlSource ="Unit_Code"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Parks.ParkCode, tlu_Parks.ParkName FROM tlu_Parks; "
                    ColumnWidths ="585;2565"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =720
                            Width =900
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="ParkCode_Label"
                            Caption ="Park Code"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =4260
                    Top =120
                    Width =4845
                    Height =420
                    FontSize =16
                    FontWeight =700
                    Name ="Label11"
                    Caption ="Edit/Add Site Characterization"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3000
                    Top =720
                    Width =540
                    TabIndex =2
                    Name ="PlotID"
                    ControlSource ="Plot_ID"
                    StatusBarText ="Plot identifier"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2280
                            Top =720
                            Width =660
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Label12"
                            Caption ="Plot ID"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4320
                    Top =720
                    Width =960
                    TabIndex =3
                    Name ="SiteDate"
                    ControlSource ="SiteDate"
                    Format ="Short Date"
                    StatusBarText ="Date site characterized"
                    InputMask ="99/99/0000;0;_"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3720
                            Top =720
                            Width =540
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Label13"
                            Caption ="Date"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1440
                    Top =1140
                    TabIndex =6
                    Name ="Soil_Map_Unit"
                    ControlSource ="Soil_Map_Unit"
                    StatusBarText ="Soil map unit in which the centroid lies"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =1140
                            Width =1200
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Label16"
                            Caption ="Soil Map Unit"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4380
                    Top =1140
                    TabIndex =7
                    Name ="GPS_File_Name"
                    ControlSource ="GPS_File_Name"
                    StatusBarText ="GPS file name for centroid coordinates"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3060
                            Top =1140
                            Width =1260
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Label17"
                            Caption ="GPS File Name"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =1620
                    Width =2100
                    TabIndex =10
                    Name ="Parent_Material"
                    ControlSource ="Parent_Material"
                    StatusBarText ="Geologic setting on site."

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =1620
                            Width =1440
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Label27"
                            Caption ="Parent Material"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4920
                    Top =1620
                    TabIndex =11
                    Name ="Landform"
                    ControlSource ="Landform"
                    StatusBarText ="Landform on site."

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3960
                            Top =1620
                            Width =900
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Label28"
                            Caption ="Landform"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7500
                    Top =2040
                    Width =480
                    TabIndex =16
                    Name ="Percent_Slope"
                    ControlSource ="Percent_Slope"
                    StatusBarText ="Percent slope at centroid - tenths"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6180
                            Top =2040
                            Width =1260
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Label33"
                            Caption ="Percent Slope"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =7140
                    Top =4680
                    Width =5220
                    Height =243
                    TabIndex =23
                    Name ="Dominant_Vegetation"
                    ControlSource ="Dominant_Vegetation"
                    StatusBarText ="Dominant vegetation at site"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5220
                            Top =4680
                            Width =1860
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Label42"
                            Caption ="Dominant Vegetation"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =3960
                    Top =5160
                    TabIndex =25
                    Name ="Soil_Assessment"
                    ControlSource ="Soil_Assessment"
                    StatusBarText ="Is it a soil assessment target match"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =5100
                            Width =3720
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Label43"
                            Caption ="Soil Assessment - (Check if a target match)"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =6180
                    Top =5100
                    Width =1800
                    TabIndex =26
                    Name ="Probable_Component"
                    ControlSource ="Probable_Component"
                    StatusBarText ="Probable soil component"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4320
                            Top =5100
                            Width =1860
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Label44"
                            Caption ="Probable Component"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =9480
                    Top =5100
                    Width =3240
                    Height =243
                    TabIndex =27
                    Name ="Soil_Comments"
                    ControlSource ="Soil_Comments"
                    StatusBarText ="Additional soil comments"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =8160
                            Top =5100
                            Width =1320
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Label47"
                            Caption ="Soil Comments"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =4440
                    Top =5580
                    TabIndex =28
                    Name ="Veg_Assessment"
                    ControlSource ="Veg_Assessment"
                    StatusBarText ="Is it a vegetation assessment target match"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =180
                            Top =5520
                            Width =4260
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Label48"
                            Caption ="Vegetation Assessment (Check if a target match)"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =6780
                    Top =5520
                    Width =3300
                    Height =243
                    TabIndex =29
                    Name ="Veg_Comments"
                    ControlSource ="Veg_Comments"
                    StatusBarText ="Vegetation assessment comments"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4800
                            Top =5520
                            Width =1980
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Label49"
                            Caption ="Vegetation Comments"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =5820
                    Top =6360
                    Width =420
                    TabIndex =33
                    Name ="Primary_Percent"
                    ControlSource ="Primary_Percent"
                    StatusBarText ="Percentage in tenths of macroplot comprised of primary site type"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =5100
                            Top =6360
                            Width =720
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Label51"
                            Caption ="% Area"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =11940
                    Top =6360
                    Width =420
                    TabIndex =35
                    Name ="Other_Percent"
                    ControlSource ="Other_Percent"
                    StatusBarText ="Percentage in tenths of macroplot comprised of other site type"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =11220
                            Top =6360
                            Width =720
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Label54"
                            Caption ="% Area"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =6060
                    Top =6780
                    Width =480
                    TabIndex =36
                    Name ="Transect_Length"
                    ControlSource ="Transect_Length"
                    StatusBarText ="Estimated total length of transect in meters that falls outside the target ecolo"
                        "gical site"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =180
                            Top =6780
                            Width =5880
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Label55"
                            Caption ="Estimated transect length in meters outside of target ecological site"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =3120
                    Top =7260
                    TabIndex =37
                    Name ="Site_Selection"
                    ControlSource ="Site_Selection"
                    StatusBarText ="Site accepted or rejected"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =7200
                            Width =2910
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Label56"
                            Caption ="Site Selection (Check if accepted)"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =4440
                    Top =7200
                    Width =7860
                    Height =243
                    TabIndex =38
                    Name ="Site_Selection_Comments"
                    ControlSource ="Site_Selection_Comments"
                    StatusBarText ="Additional site selection comments"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3480
                            Top =7200
                            Width =960
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Label57"
                            Caption ="Comments"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =2580
                    Top =7620
                    Width =360
                    TabIndex =40
                    Name ="Soil_Profile_Samples"
                    ControlSource ="Soil_Profile_Samples"
                    StatusBarText ="Number of soil profile samples collected"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =180
                            Top =7620
                            Width =2400
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Label58"
                            Caption ="Soil Profile (=10cm groups)"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4860
                    Top =7620
                    Width =360
                    TabIndex =41
                    Name ="2cm_Samples"
                    ControlSource ="2cm_Samples"
                    StatusBarText ="Number of 0-2 cm composite samples collected"
                    EventProcPrefix ="Ctl2cm_Samples"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3240
                            Top =7620
                            Width =1560
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Label59"
                            Caption ="0-2 cm composite"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =7200
                    Top =7620
                    Width =360
                    TabIndex =42
                    Name ="10cm_Samples"
                    ControlSource ="10cm_Samples"
                    StatusBarText ="Number of 2-10 cm composite samples collected"
                    EventProcPrefix ="Ctl10cm_Samples"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =5520
                            Top =7620
                            Width =1680
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Label60"
                            Caption ="2-10 cm composite"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Line
                    BorderWidth =2
                    OverlapFlags =85
                    Left =120
                    Top =1500
                    Width =12720
                    Name ="Line61"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    ListWidth =885
                    Left =8100
                    Top =1620
                    Width =1200
                    TabIndex =12
                    ColumnInfo ="\"\";\"\";\"10\";\"18\""
                    Name ="Hillslope_Position"
                    ControlSource ="Hillslope_Position"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Hillslope_Position.Position FROM tlu_Hillslope_Position; "
                    ColumnWidths ="885"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6540
                            Top =1620
                            Width =1560
                            Height =245
                            FontSize =8
                            FontWeight =700
                            Name ="Hillslope Position_Label"
                            Caption ="Hillslope Position"
                            FontName ="Tahoma"
                            EventProcPrefix ="Hillslope_Position_Label"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =765
                    Left =11040
                    Top =1620
                    Width =1080
                    TabIndex =13
                    Name ="Slope_Complexity"
                    ControlSource ="Slope_Complexity"
                    RowSourceType ="Value List"
                    RowSource ="\"complex\";\"simple\""
                    ColumnWidths ="765"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9480
                            Top =1620
                            Width =1500
                            Height =245
                            FontSize =8
                            FontWeight =700
                            Name ="Slope Complexity_Label"
                            Caption ="Slope Complexity"
                            FontName ="Tahoma"
                            EventProcPrefix ="Slope_Complexity_Label"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =780
                    Left =1860
                    Top =2040
                    Width =1080
                    TabIndex =14
                    ColumnInfo ="\"\";\"\";\"10\";\"14\""
                    Name ="Slope_Shape_Down"
                    ControlSource ="Slope_Shape_Down"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Slope_Shape.Shape FROM tlu_Slope_Shape; "
                    ColumnWidths ="780"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =2040
                            Width =1620
                            Height =245
                            FontSize =8
                            FontWeight =700
                            Name ="Slope Shape Down_Label"
                            Caption ="Slope Shape Down"
                            FontName ="Tahoma"
                            EventProcPrefix ="Slope_Shape_Down_Label"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =780
                    Left =4920
                    Top =2040
                    Width =1080
                    TabIndex =15
                    ColumnInfo ="\"\";\"\";\"10\";\"14\""
                    Name ="Slope_Shape_Across"
                    ControlSource ="Slope_Shape_Across"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Slope_Shape.Shape FROM tlu_Slope_Shape; "
                    ColumnWidths ="780"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3120
                            Top =2040
                            Width =1740
                            Height =245
                            FontSize =8
                            FontWeight =700
                            Name ="Slope Shape Across_Label"
                            Caption ="Slope Shape Across"
                            FontName ="Tahoma"
                            EventProcPrefix ="Slope_Shape_Across_Label"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =1035
                    Left =1440
                    Top =2460
                    Width =2100
                    TabIndex =18
                    ColumnInfo ="\"\";\"\";\"10\";\"50\""
                    Name ="Sand_Modifier"
                    ControlSource ="Sand_Modifier"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Sand_Modifier.Modifier FROM tlu_Sand_Modifier; "
                    ColumnWidths ="1035"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =2460
                            Width =1200
                            Height =245
                            FontSize =8
                            FontWeight =700
                            Name ="Sand Modifier_Label"
                            Caption ="Sand Modifier"
                            FontName ="Tahoma"
                            EventProcPrefix ="Sand_Modifier_Label"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2085
                    Left =5760
                    Top =2460
                    Width =1080
                    TabIndex =19
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"18\""
                    Name ="Rock_Fragment_Qty"
                    ControlSource ="Rock_Fragment_Qty"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Rock_Frag_Q.Quantity, tlu_Rock_Frag_Q.Descriptor FROM tlu_Rock_Frag_Q"
                        "; "
                    ColumnWidths ="900;1185"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3720
                            Top =2460
                            Width =1980
                            Height =245
                            FontSize =8
                            FontWeight =700
                            Name ="Rock Fragment Quantity Modifier_Label"
                            Caption ="Rock Frag Qty Modifier"
                            FontName ="Tahoma"
                            EventProcPrefix ="Rock_Fragment_Quantity_Modifier_Label"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =855
                    Left =9120
                    Top =2460
                    Width =1080
                    TabIndex =20
                    ColumnInfo ="\"\";\"\";\"10\";\"16\""
                    Name ="Rock_Fragment_Size"
                    ControlSource ="Rock_Fragment_Size"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Rock_Frag_Size.Size FROM tlu_Rock_Frag_Size; "
                    ColumnWidths ="855"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7020
                            Top =2460
                            Width =2040
                            Height =245
                            FontSize =8
                            FontWeight =700
                            Name ="Rock Fragment Size Modifier_Label"
                            Caption ="Rock Frag Size Modifier"
                            FontName ="Tahoma"
                            EventProcPrefix ="Rock_Fragment_Size_Modifier_Label"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =870
                    Left =11700
                    Top =2460
                    Width =1020
                    TabIndex =21
                    ColumnInfo ="\"\";\"\";\"10\";\"22\""
                    Name ="Effervescence"
                    ControlSource ="Effervescence"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Effervescence.Effervescence FROM tlu_Effervescence; "
                    ColumnWidths ="870"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10380
                            Top =2460
                            Width =1260
                            Height =245
                            FontSize =8
                            FontWeight =700
                            Name ="Effervescence_Label"
                            Caption ="Effervescence"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Line
                    BorderWidth =2
                    OverlapFlags =85
                    Left =120
                    Top =2820
                    Width =12720
                    Name ="Line78"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =1395
                    Left =1140
                    Top =4680
                    Width =3720
                    TabIndex =22
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="Depth"
                    ControlSource ="Depth"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Soil_Depth.Depth FROM tlu_Soil_Depth; "
                    ColumnWidths ="1395"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =4680
                            Width =900
                            Height =245
                            FontSize =8
                            FontWeight =700
                            Name ="Soil Depth_Label"
                            Caption ="Soil Depth"
                            FontName ="Tahoma"
                            EventProcPrefix ="Soil_Depth_Label"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =1260
                    Top =3000
                    Width =10260
                    Height =1500
                    TabIndex =44
                    Name ="fsub_Soil_Profile"
                    SourceObject ="Form.fsub_Soil_Profile"
                    LinkChildFields ="Location_ID"
                    LinkMasterFields ="Location_ID"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =480
                            Top =3000
                            Width =720
                            Height =480
                            FontSize =8
                            FontWeight =700
                            Name ="fsub_Soil_Profile Label"
                            Caption ="Soil Profile"
                            FontName ="Tahoma"
                            EventProcPrefix ="fsub_Soil_Profile_Label"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1980
                    Top =180
                    Width =1545
                    Height =300
                    TabIndex =46
                    Name ="Button_Plots"
                    Caption ="Plot Establishment"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =6300
                    Top =720
                    Width =1980
                    TabIndex =4
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Recorder"
                    ControlSource ="Recorder"
                    RowSourceType ="Table/Query"
                    RowSource ="qry_Contacts"
                    ColumnWidths ="0;1440"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =5460
                            Top =720
                            Width =840
                            Height =245
                            FontSize =8
                            FontWeight =700
                            Name ="Recorder_Label"
                            Caption ="Recorder"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1140
                    Top =60
                    Width =420
                    TabIndex =47
                    Name ="UTM_Zone"
                    ControlSource ="UTM_Zone"
                    StatusBarText ="MA. UTM Zone (UTM_Zone)"

                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1140
                    Top =360
                    Width =420
                    TabIndex =48
                    Name ="Datum"
                    ControlSource ="Datum"
                    StatusBarText ="M. Datum of mapping ellipsoid (Datum)"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =1500
                    Left =10020
                    Top =720
                    Width =2760
                    TabIndex =5
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="Soil Survey Area"
                    ControlSource ="Soil_Survey_Area"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Soil_Survey_Area.Soil_Survey_Area FROM tlu_Soil_Survey_Area; "
                    ColumnWidths ="1500"
                    EventProcPrefix ="Soil_Survey_Area"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8460
                            Top =720
                            Width =1500
                            Height =245
                            FontSize =8
                            FontWeight =700
                            Name ="Soil Survey Area_Label"
                            Caption ="Soil Survey Area"
                            FontName ="Tahoma"
                            EventProcPrefix ="Soil_Survey_Area_Label"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ListWidth =1320
                    Left =9300
                    Top =2040
                    Width =2820
                    TabIndex =17
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="Soil_Texture"
                    ControlSource ="Soil_Texture"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Soil_Texture.Soil_Texture FROM tlu_Soil_Texture; "
                    ColumnWidths ="1320"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =8160
                            Top =2040
                            Width =1920
                            Height =245
                            FontSize =8
                            FontWeight =700
                            Name ="Soil Texture_Label"
                            Caption ="Soil Texture"
                            FontName ="Tahoma"
                            EventProcPrefix ="Soil_Texture_Label"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ListWidth =1080
                    Left =2160
                    Top =6360
                    Width =2880
                    TabIndex =32
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="Primary_Eco_Site"
                    ControlSource ="Primary_Eco_Site"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Eco_Site.Eco_Site FROM tlu_Eco_Site; "
                    ColumnWidths ="1080"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =180
                            Top =6360
                            Width =4545
                            Height =245
                            FontSize =8
                            FontWeight =700
                            Name ="Primary Ecological Site_Label"
                            Caption ="Primary Ecological Site"
                            FontName ="Tahoma"
                            EventProcPrefix ="Primary_Ecological_Site_Label"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =10740
                    Top =7500
                    Width =2100
                    TabIndex =43
                    Name ="Geologic_Setting"
                    ControlSource ="Geologic_Setting"
                    StatusBarText ="Geologic setting on site."

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =9120
                            Top =7500
                            Width =1605
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Label97"
                            Caption ="Geological Setting"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =6360
                    Top =6000
                    TabIndex =31
                    Name ="Grazed_field"
                    ControlSource ="Grazed_field"
                    StatusBarText ="Is it a grazed field?"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =5640
                            Top =5940
                            Width =720
                            Height =270
                            FontSize =8
                            FontWeight =700
                            Name ="Label99"
                            Caption ="Grazed"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    ListWidth =1635
                    Left =1680
                    Top =5940
                    Width =3360
                    TabIndex =30
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="Vegetation_Type"
                    ControlSource ="Vegetation_Type"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Veg_Type.Vegetation_Type FROM tlu_Veg_Type; "
                    ColumnWidths ="1635"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =180
                            Top =5940
                            Width =1500
                            Height =245
                            FontSize =8
                            FontWeight =700
                            Name ="Vegetation Type_Label"
                            Caption ="Vegetation Type"
                            FontName ="Tahoma"
                            EventProcPrefix ="Vegetation_Type_Label"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    ListWidth =1080
                    Left =8280
                    Top =6360
                    Width =2880
                    TabIndex =34
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="Other_Eco_Site"
                    ControlSource ="Other_Eco_Site"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Eco_Site.Eco_Site FROM tlu_Eco_Site; "
                    ColumnWidths ="1080"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6480
                            Top =6360
                            Width =1800
                            Height =245
                            FontSize =8
                            FontWeight =700
                            Name ="Other Ecological Site_Label"
                            Caption ="Other Ecological Site"
                            FontName ="Tahoma"
                            EventProcPrefix ="Other_Ecological_Site_Label"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6060
                    Top =8820
                    Width =1035
                    Height =300
                    TabIndex =49
                    Name ="ButtonClose"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =12300
                    Top =7200
                    Width =306
                    Height =306
                    TabIndex =39
                    ForeColor =0
                    Name ="ButtonZoomSiteSelectionComments"
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
                Begin CommandButton
                    OverlapFlags =87
                    Left =12360
                    Top =4680
                    Width =306
                    Height =306
                    TabIndex =24
                    ForeColor =0
                    Name ="ButtonZoomDominantVegetation"
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

Private Sub ButtonZoomDominantVegetation_Click()
On Error GoTo Err_ButtonDominantVegetation_Click

  Me!Dominant_Vegetation.SetFocus
  SendKeys ("+{F2}")
  
Exit_ButtonDominantVegetation_Click:
    Exit Sub

Err_ButtonDominantVegetation_Click:
    MsgBox Err.Description
    Resume Exit_ButtonDominantVegetation_Click

End Sub

Private Sub ButtonZoomSiteSelectionComments_Click()
On Error GoTo Err_ButtonSiteSelectionComments_Click

  Me!Site_Selection_Comments.SetFocus
  SendKeys ("+{F2}")
  
Exit_ButtonSiteSelectionComments_Click:
    Exit Sub

Err_ButtonSiteSelectionComments_Click:
    MsgBox Err.Description
    Resume Exit_ButtonSiteSelectionComments_Click

End Sub



' =================================
' Description:  Locations entry form
' Data source:  tbl_Locations
' Data access:  edit, add, delete
' Pages:        none
' Functions:    none
' References:   fxnGUIDGen
' Source/date:  Rescued from Simon Kingston, Sept. 2006 by Russ DenBleyker
' Revisions:    <name, date, desc - add lines as you go>
' =================================


Private Sub Form_BeforeUpdate(Cancel As Integer)
'check to see if a primary key is needed and add it (used for string GUIDs)

    Me!txtUpdated_date = Now()
    If IsNull(Me!txtLocation_ID) Then
        If GetDataType("tbl_Locations", "Location_ID") = dbText Then
            Me!txtLocation_ID = fxnGUIDGen
        End If
    End If

End Sub

Private Sub Form_Close()

'update control as necessary on calling form to reflect new location values
UpdateControl Me.OpenArgs
End Sub

Private Sub Form_Current()
'check to see if a primary key is needed and add it (used for string GUIDs)
  If Me.NewRecord Then
    If GetDataType("tbl_Locations", "Location_ID") = dbText Then
        Me!txtLocation_ID = fxnGUIDGen
    End If
  End If
  If Not IsNull(Me!Unit_Code) Then
    Me.Primary_Eco_Site.RowSource = "SELECT Eco_Site FROM xref_Park_EcoSite WHERE [Unit_Code] = '" & Me!Unit_Code & "' ORDER BY Eco_Site"
    Me.Primary_Eco_Site.Requery
    Me.Other_Eco_Site.RowSource = "SELECT Eco_Site FROM xref_Park_EcoSite WHERE [Unit_Code] = '" & Me!Unit_Code & "' ORDER BY Eco_Site"
    Me.Other_Eco_Site.Requery
  End If
End Sub
Private Sub Button_Plots_Click()
On Error GoTo Err_Button_Plots_Click
Dim stDocName As String
Dim stLinkCriteria As String

  If Me!Site_Selection = 0 Or IsNull(Me!Site_Selection) Then
    MsgBox "This site has not been accepted for monitoring - action cancelled."
  Else
    DoCmd.RunCommand acCmdSaveRecord  ' Save record so next form can find it.
    stDocName = "frm_Plot_Establishment"
    stLinkCriteria = "[Location_ID]=" & "'" & Me![txtLocation_ID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
  End If
Exit_Button_Plots_Click:
    Exit Sub

Err_Button_Plots_Click:
    MsgBox Err.Description
    Resume Exit_Button_Plots_Click
    
End Sub

Private Sub Form_Load()
  Dim strCriteria As String
  If Me.OpenArgs = "New Record" Then
    strCriteria = "Project = 'NCPN Upland Monitoring'"
    Me!Unit_Code = DLookup("Park", "tsys_App_Defaults", strCriteria)
    Me!Datum = DLookup("Datum", "tsys_App_Defaults", strCriteria)
    Me!UTM_Zone = DLookup("Zone", "tsys_App_Defaults", strCriteria)
    Me!Soil_Survey_Area = DLookup("Soil_Survey_Area", "tsys_App_Defaults", strCriteria)
  End If

'  DoCmd.Maximize

End Sub

Private Sub PlotID_AfterUpdate()
  Dim strCriteria As String
  Dim strMessage As String
  
  strCriteria = "Unit_Code = '" & Me!Unit_Code & "' And Plot_Id = " & Me!Plot_ID
  If Not IsNull(DLookup("Location_ID", "tbl_Locations", strCriteria)) Then
    strMessage = "Plot " & Me!Plot_ID & " already exists for " & Me!Unit_Code & "."
    MsgBox strMessage
    DoCmd.CancelEvent
    SendKeys "{ESC}"
    Me!Unit_Code.SetFocus
    If Me.OpenArgs = "New Record" Then
      strCriteria = "Project = 'NCPN Upland Monitoring'"
      Me!Unit_Code = DLookup("Park", "tsys_App_Defaults", strCriteria)
      Me!Datum = DLookup("Datum", "tsys_App_Defaults", strCriteria)
      Me!UTM_Zone = DLookup("Zone", "tsys_App_Defaults", strCriteria)
    End If
  End If

End Sub
Private Sub ButtonClose_Click()
On Error GoTo Err_ButtonClose_Click


    DoCmd.Close

Exit_ButtonClose_Click:
    Exit Sub

Err_ButtonClose_Click:
    MsgBox Err.Description
    Resume Exit_ButtonClose_Click
    
End Sub

Private Sub Unit_Code_AfterUpdate()
  If Not IsNull(Me!Unit_Code) Then
    Me.Primary_Eco_Site.RowSource = "SELECT Eco_Site FROM xref_Park_EcoSite WHERE [Unit_Code] = '" & Me!Unit_Code & "' ORDER BY Eco_Site"
    Me.Primary_Eco_Site.Requery
    Me.Other_Eco_Site.RowSource = "SELECT Eco_Site FROM xref_Park_EcoSite WHERE [Unit_Code] = '" & Me!Unit_Code & "' ORDER BY Eco_Site"
    Me.Other_Eco_Site.Requery
  End If
End Sub
