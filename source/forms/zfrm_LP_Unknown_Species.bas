Version =21
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =124
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =11520
    DatasheetFontHeight =9
    ItemSuffix =43
    Left =1935
    Top =3450
    Right =13080
    Bottom =9735
    DatasheetGridlinesColor =12632256
    Filter ="[Species_ID]='{A49DAA5E-5889-4FAB-B295-3EC3CDAE8101}'"
    RecSrcDt = Begin
        0x6139e0c6cd5be340
    End
    RecordSource ="tbl_Unknown_Species"
    Caption ="Enter Unknown Species"
    BeforeInsert ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
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
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =7200
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =10140
                    Top =120
                    Width =570
                    ColumnWidth =2310
                    Name ="Unknown_ID"
                    ControlSource ="Unknown_ID"
                    StatusBarText ="Unique record identifier - primary key"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10800
                    Top =120
                    Width =570
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Species_ID"
                    ControlSource ="Species_ID"
                    StatusBarText ="Foreign key to tbl_Quadrat_Species"

                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1980
                    Top =1740
                    Width =7380
                    TabIndex =6
                    Name ="Plant_Description"
                    ControlSource ="Plant_Description"
                    StatusBarText ="General description"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =180
                            Top =1740
                            Width =1800
                            Height =240
                            FontWeight =700
                            Name ="Plant_Description_Label"
                            Caption ="General Description"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =2040
                    Top =2220
                    Width =7320
                    TabIndex =7
                    Name ="Salient_Feature"
                    ControlSource ="Salient_Feature"
                    StatusBarText ="Most salient feature"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =180
                            Top =2220
                            Width =1860
                            Height =240
                            FontWeight =700
                            Name ="Salient_Feature_Label"
                            Caption ="Most Salient Feature"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1140
                    Top =2700
                    Width =3180
                    ColumnWidth =2310
                    TabIndex =8
                    Name ="Leaf_Type"
                    ControlSource ="Leaf_Type"
                    StatusBarText ="Leaf type: compound/simple, arrangement"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =180
                            Top =2700
                            Width =960
                            Height =240
                            FontWeight =700
                            Name ="Leaf_Type_Label"
                            Caption ="Leaf Type"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1260
                    Top =3180
                    Width =3060
                    ColumnWidth =2310
                    TabIndex =9
                    Name ="Margin"
                    ControlSource ="Margin"
                    StatusBarText ="Leaf margin"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =180
                            Top =3180
                            Width =1110
                            Height =240
                            FontWeight =700
                            Name ="Margin_Label"
                            Caption ="Leaf Margin"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =2520
                    Top =3660
                    Width =6840
                    TabIndex =10
                    Name ="Other_Characteristics"
                    ControlSource ="Other_Characteristics"
                    StatusBarText ="Other leaf characteristics:  pubescence, sap, stipules"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =180
                            Top =3660
                            Width =2340
                            Height =240
                            FontWeight =700
                            Name ="Other_Characteristics_Label"
                            Caption ="Other Leaf Characteristics"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =2040
                    Top =4140
                    Width =7320
                    TabIndex =11
                    Name ="Stem_Characteristics"
                    ControlSource ="Stem_Characteristics"
                    StatusBarText ="Stem characteristics: shape, pubescence, bud"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =180
                            Top =4140
                            Width =1860
                            Height =240
                            FontWeight =700
                            Name ="Stem_Characteristics_Label"
                            Caption ="Stem Characteristics"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =2160
                    Top =4620
                    Width =7200
                    TabIndex =12
                    Name ="Flower_Characteristics"
                    ControlSource ="Flower_Characteristics"
                    StatusBarText ="Flower characteristics: color location floral formula"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =180
                            Top =4620
                            Width =1980
                            Height =240
                            FontWeight =700
                            Name ="Flower_Characteristics_Label"
                            Caption ="Flower Characteristics"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =3720
                    Top =5100
                    Width =5640
                    TabIndex =13
                    Name ="General_Characteristics"
                    ControlSource ="General_Characteristics"
                    StatusBarText ="General and microhabitat characteristics"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =180
                            Top =5100
                            Width =3540
                            Height =240
                            FontWeight =700
                            Name ="General_Characteristics_Label"
                            Caption ="General and Microhabitat Characteristics"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =1200
                    Top =5640
                    Width =735
                    Height =300
                    TabIndex =14
                    Name ="Collected"
                    ControlSource ="Collected"
                    StatusBarText ="Was plant collected"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =5580
                            Width =960
                            Height =240
                            FontWeight =700
                            Name ="Collected_Label"
                            Caption ="Collected?"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1260
                    Top =6060
                    Width =2310
                    ColumnWidth =2310
                    TabIndex =17
                    Name ="Best_Guess"
                    ControlSource ="Best_Guess"
                    StatusBarText ="Best guess species name"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =180
                            Top =6060
                            Width =1080
                            Height =240
                            FontWeight =700
                            Name ="Best_Guess_Label"
                            Caption ="Best Guess"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1620
                    Top =6540
                    Width =2310
                    ColumnWidth =2310
                    TabIndex =18
                    Name ="Confirmed"
                    ControlSource ="Confirmed"
                    StatusBarText ="Confirmed species name"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =180
                            Top =6540
                            Width =1455
                            Height =240
                            FontWeight =700
                            Name ="Confirmed_Label"
                            Caption ="Confirmed to be"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =4020
                    Top =120
                    Width =3420
                    Height =420
                    FontSize =14
                    FontWeight =700
                    Name ="Label28"
                    Caption ="Unknown Plant Species"
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =87
                    IMESentenceMode =3
                    ListWidth =675
                    Left =1200
                    Top =1260
                    Width =1380
                    TabIndex =3
                    Name ="Plant_Type"
                    ControlSource ="Plant_Type"
                    RowSourceType ="Value List"
                    RowSource ="\"tree\";\"shrub\";\"grass\";\"forb\";\"other\""
                    ColumnWidths ="675"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =180
                            Top =1260
                            Width =1020
                            Height =245
                            FontWeight =700
                            Name ="Plant Type_Label"
                            Caption ="Plant Type"
                            EventProcPrefix ="Plant_Type_Label"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =9300
                    Top =180
                    Width =1020
                    Height =300
                    TabIndex =21
                    Name ="ButtonClose"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1560
                    Top =780
                    Width =4200
                    TabIndex =2
                    Name ="Unknown_Code"
                    ControlSource ="Unknown_Code"
                    StatusBarText ="Temporary code for unknown species - Line point form"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =180
                            Top =780
                            Width =1395
                            Height =240
                            FontWeight =700
                            Name ="Label32"
                            Caption ="Unknown Code"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =87
                    IMESentenceMode =3
                    ListWidth =795
                    Left =4560
                    Top =1260
                    Width =1200
                    TabIndex =4
                    Name ="Forb_Grass_Type"
                    ControlSource ="Forb_Grass_Type"
                    RowSourceType ="Value List"
                    RowSource ="\"annual\";\"perennial\""
                    ColumnWidths ="795"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2880
                            Top =1260
                            Width =1680
                            Height =245
                            FontWeight =700
                            Name ="Forbs and Grasses_Label"
                            Caption ="Forbs and Grasses"
                            EventProcPrefix ="Forbs_and_Grasses_Label"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =1050
                    Left =7680
                    Top =1260
                    TabIndex =5
                    Name ="Perennial_Grasses"
                    ControlSource ="Perennial_Grasses"
                    RowSourceType ="Value List"
                    RowSource ="\"bunchgrass\";\"rhizomatous\""
                    ColumnWidths ="1050"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6060
                            Top =1260
                            Width =1575
                            Height =245
                            FontWeight =700
                            Name ="Perennial Grasses_Label"
                            Caption ="Perennial Grasses"
                            EventProcPrefix ="Perennial_Grasses_Label"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =119
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =1980
                    Left =2940
                    Top =5580
                    Width =1620
                    TabIndex =15
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Collected_by"
                    ControlSource ="Collected_by"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Contacts.Contact_ID, tlu_Contacts.Last_Name, tlu_Contacts.First_Name "
                        "FROM tlu_Contacts; "
                    ColumnWidths ="0;990;990"
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =1740
                            Top =5580
                            Width =1200
                            Height =245
                            FontWeight =700
                            Name ="Collected by_Label"
                            Caption ="Collected by"
                            EventProcPrefix ="Collected_by_Label"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =1980
                    Left =5400
                    Top =6540
                    Width =1620
                    TabIndex =19
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Identified_by"
                    ControlSource ="Identified_by"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Contacts.Contact_ID, tlu_Contacts.Last_Name, tlu_Contacts.First_Name "
                        "FROM tlu_Contacts; "
                    ColumnWidths ="0;990;990"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4260
                            Top =6540
                            Width =1140
                            Height =245
                            FontWeight =700
                            Name ="Identified by_Label"
                            Caption ="Identified by"
                            EventProcPrefix ="Identified_by_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =7800
                    Top =6540
                    Width =960
                    TabIndex =20
                    Name ="Identified_Date"
                    ControlSource ="Identified_Date"
                    Format ="Short Date"
                    StatusBarText ="Date of identification - Line point form"
                    InputMask ="99/99/0000;0;_"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =7320
                            Top =6540
                            Width =480
                            Height =240
                            FontWeight =700
                            Name ="Label41"
                            Caption ="Date"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =5760
                    Top =5640
                    TabIndex =16
                    Name ="Have_Photos"
                    ControlSource ="Have_Photos"
                    StatusBarText ="Are there photos? - Line point form"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4920
                            Top =5580
                            Width =780
                            Height =240
                            FontWeight =700
                            Name ="Label42"
                            Caption ="Photos"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
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

    DoCmd.RunCommand acCmdSaveRecord  ' Save record.
    DoCmd.Close

Exit_ButtonClose_Click:
    Exit Sub

Err_ButtonClose_Click:
    MsgBox Err.Description
    Resume Exit_ButtonClose_Click
    
End Sub



Private Sub Form_BeforeInsert(Cancel As Integer)
    On Error GoTo Err_Handler
    ' Create the GUID primary key value
    If IsNull(Me!Unknown_ID) Then
        If GetDataType("tbl_Unknown_Species", "Unknown_ID") = dbText Then
            Me.Unknown_ID = fxnGUIDGen
        End If
    End If
    Me!Species_ID = Me.OpenArgs  ' Set foreign key

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub
