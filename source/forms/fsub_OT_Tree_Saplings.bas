Version =21
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11340
    DatasheetFontHeight =9
    ItemSuffix =31
    Top =270
    Right =9330
    Bottom =3735
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x384b3f359387e340
    End
    RecordSource ="tbl_OT_Tree_Saplings"
    Caption ="fsub_OT_Tree_Saplings"
    BeforeInsert ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnDeactivate ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
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
            Height =1200
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Rectangle
                    BackStyle =1
                    OverlapFlags =93
                    Left =4725
                    Top =660
                    Width =1008
                    Height =540
                    BackColor =13434828
                    Name ="rct2"
                    LayoutCachedLeft =4725
                    LayoutCachedTop =660
                    LayoutCachedWidth =5733
                    LayoutCachedHeight =1200
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =225
                    Top =720
                    Width =1335
                    Height =240
                    FontWeight =700
                    Name ="Species_Label"
                    Caption ="Tree Species"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2580
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
                    Left =3660
                    Top =960
                    Width =930
                    Height =240
                    FontWeight =700
                    Name ="HC25_Label"
                    Caption ="2.5-5.0cm"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =3660
                    LayoutCachedTop =960
                    LayoutCachedWidth =4590
                    LayoutCachedHeight =1200
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =4703
                    Top =960
                    Width =1035
                    Height =240
                    FontWeight =700
                    Name ="HC50_Label"
                    Caption ="5.1-10.0cm"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =4703
                    LayoutCachedTop =960
                    LayoutCachedWidth =5738
                    LayoutCachedHeight =1200
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =5730
                    Top =960
                    Width =1140
                    Height =240
                    FontWeight =700
                    Name ="HC100_Label"
                    Caption ="10.1-15.0cm"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =5730
                    LayoutCachedTop =960
                    LayoutCachedWidth =6870
                    LayoutCachedHeight =1200
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =3600
                    Top =480
                    Width =3255
                    Height =240
                    FontWeight =700
                    BackColor =14277081
                    Name ="Label22"
                    Caption ="Diameter Class Totals"
                    LayoutCachedLeft =3600
                    LayoutCachedTop =480
                    LayoutCachedWidth =6855
                    LayoutCachedHeight =720
                    BackThemeColorIndex =1
                    BackShade =85.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1620
                    Top =60
                    Width =5760
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label23"
                    Caption ="Number of Tree Saplings in 5 Meter Belt Transect"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7500
                    Top =120
                    Width =1545
                    Height =300
                    Name ="btnMaster"
                    Caption ="Master Lookup"
                    OnClick ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"

                    LayoutCachedLeft =7500
                    LayoutCachedTop =120
                    LayoutCachedWidth =9045
                    LayoutCachedHeight =420
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7500
                    Top =540
                    Width =1545
                    Height =300
                    TabIndex =1
                    Name ="btnUnknown"
                    Caption ="Unknown Species"
                    OnClick ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"

                    LayoutCachedLeft =7500
                    LayoutCachedTop =540
                    LayoutCachedWidth =9045
                    LayoutCachedHeight =840
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =4020
                    Top =735
                    Width =195
                    Height =240
                    FontSize =5
                    FontWeight =700
                    Name ="lbl1"
                    Caption ="1"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =4020
                    LayoutCachedTop =735
                    LayoutCachedWidth =4215
                    LayoutCachedHeight =975
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =5100
                    Top =735
                    Width =195
                    Height =240
                    FontSize =5
                    FontWeight =700
                    BackColor =13434828
                    Name ="lbl2"
                    Caption ="2"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =5100
                    LayoutCachedTop =735
                    LayoutCachedWidth =5295
                    LayoutCachedHeight =975
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =6180
                    Top =735
                    Width =195
                    Height =240
                    FontSize =5
                    FontWeight =700
                    Name ="lbl3"
                    Caption ="3"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =6180
                    LayoutCachedTop =735
                    LayoutCachedWidth =6375
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
                    Left =4740
                    Width =1008
                    Height =420
                    BackColor =13434828
                    Name ="rct2data"
                    LayoutCachedLeft =4740
                    LayoutCachedWidth =5748
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =360
                    Height =255
                    ColumnWidth =2310
                    TabIndex =6
                    Name ="Shrub_ID"
                    ControlSource ="TS_ID"
                    StatusBarText ="Unique record identifier - primary key"

                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =660
                    Top =60
                    Width =300
                    Height =255
                    ColumnWidth =2310
                    TabIndex =7
                    Name ="Transect_ID"
                    ControlSource ="Event_ID"
                    StatusBarText ="Foreign key to tbl_Canopy_Transect"

                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3855
                    Top =60
                    Width =540
                    Height =255
                    ColumnWidth =600
                    TabIndex =2
                    BackColor =65535
                    Name ="HC25"
                    ControlSource ="D25"
                    StatusBarText ="10.1-25cm height class total"
                    DefaultValue ="Null"
                    OnGotFocus ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000012010000030000000100000000000000000000001800000001000000 ,
                        0x00000000ffffff00010000000000000019000000310000000100000000000000 ,
                        0xffff0000010000000000000032000000580000000100000000000000ffffff00 ,
                        0x4900490066002800490073004e0075006c006c0028005b004800430032003500 ,
                        0x5d0029002c0030002c0031002900000000004900490066002800490073004e00 ,
                        0x75006c006c0028005b0048004300320035005d0029002c0031002c0030002900 ,
                        0x000000005b0050006100720065006e0074005d002e005b006300620078004e00 ,
                        0x6f005300610070006c0069006e00670073005d002e005b00560061006c007500 ,
                        0x65005d003d00540072007500650000000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000400000001000000000000000100000000000000ffffff00170000004900 ,
                        0x490066002800490073004e0075006c006c0028005b0048004300320035005d00 ,
                        0x29002c0030002c00310029000000000000000000000000000000000000000000 ,
                        0x0001000000000000000100000000000000ffff00001700000049004900660028 ,
                        0x00490073004e0075006c006c0028005b0048004300320035005d0029002c0031 ,
                        0x002c003000290000000000000000000000000000000000000000000001000000 ,
                        0x000000000100000000000000ffffff00250000005b0050006100720065006e00 ,
                        0x74005d002e005b006300620078004e006f005300610070006c0069006e006700 ,
                        0x73005d002e005b00560061006c00750065005d003d0054007200750065000000 ,
                        0x0000000000000000000000000000000000000001000000000000000100000000 ,
                        0x000000ffff6600170000005b0048004300320035005d002b005b004800430035 ,
                        0x0030005d002b005b00480043003100300030005d003d00300000000000000000 ,
                        0x0000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4995
                    Top =60
                    Width =540
                    Height =255
                    ColumnWidth =600
                    TabIndex =3
                    BackColor =65535
                    Name ="HC50"
                    ControlSource ="D51"
                    StatusBarText ="25.1-50cm height class total"
                    DefaultValue ="Null"
                    OnGotFocus ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000012010000030000000100000000000000000000001800000001000000 ,
                        0x00000000ffffff00010000000000000019000000310000000100000000000000 ,
                        0xffff0000010000000000000032000000580000000100000000000000ffffff00 ,
                        0x4900490066002800490073004e0075006c006c0028005b004800430035003000 ,
                        0x5d0029002c0030002c0031002900000000004900490066002800490073004e00 ,
                        0x75006c006c0028005b0048004300350030005d0029002c0031002c0030002900 ,
                        0x000000005b0050006100720065006e0074005d002e005b006300620078004e00 ,
                        0x6f005300610070006c0069006e00670073005d002e005b00560061006c007500 ,
                        0x65005d003d00540072007500650000000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000400000001000000000000000100000000000000ffffff00170000004900 ,
                        0x490066002800490073004e0075006c006c0028005b0048004300350030005d00 ,
                        0x29002c0030002c00310029000000000000000000000000000000000000000000 ,
                        0x0001000000000000000100000000000000ffff00001700000049004900660028 ,
                        0x00490073004e0075006c006c0028005b0048004300350030005d0029002c0031 ,
                        0x002c003000290000000000000000000000000000000000000000000001000000 ,
                        0x000000000100000000000000ffffff00250000005b0050006100720065006e00 ,
                        0x74005d002e005b006300620078004e006f005300610070006c0069006e006700 ,
                        0x73005d002e005b00560061006c00750065005d003d0054007200750065000000 ,
                        0x0000000000000000000000000000000000000001000000000000000100000000 ,
                        0x000000ffff6600170000005b0048004300320035005d002b005b004800430035 ,
                        0x0030005d002b005b00480043003100300030005d003d00300000000000000000 ,
                        0x0000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6015
                    Top =60
                    Width =540
                    Height =255
                    ColumnWidth =600
                    TabIndex =4
                    BackColor =65535
                    Name ="HC100"
                    ControlSource ="D101"
                    StatusBarText ="50.1-100cm height class total"
                    DefaultValue ="Null"
                    OnGotFocus ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000016010000030000000100000000000000000000001900000001000000 ,
                        0x00000000ffffff0001000000000000001a000000330000000100000000000000 ,
                        0xffff00000100000000000000340000005a0000000100000000000000ffffff00 ,
                        0x4900490066002800490073004e0075006c006c0028005b004800430031003000 ,
                        0x30005d0029002c0030002c003100290000000000490049006600280049007300 ,
                        0x4e0075006c006c0028005b00480043003100300030005d0029002c0031002c00 ,
                        0x30002900000000005b0050006100720065006e0074005d002e005b0063006200 ,
                        0x78004e006f005300610070006c0069006e00670073005d002e005b0056006100 ,
                        0x6c00750065005d003d00540072007500650000000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000400000001000000000000000100000000000000ffffff00180000004900 ,
                        0x490066002800490073004e0075006c006c0028005b0048004300310030003000 ,
                        0x5d0029002c0030002c0031002900000000000000000000000000000000000000 ,
                        0x00000001000000000000000100000000000000ffff0000180000004900490066 ,
                        0x002800490073004e0075006c006c0028005b00480043003100300030005d0029 ,
                        0x002c0031002c0030002900000000000000000000000000000000000000000000 ,
                        0x01000000000000000100000000000000ffffff00250000005b00500061007200 ,
                        0x65006e0074005d002e005b006300620078004e006f005300610070006c006900 ,
                        0x6e00670073005d002e005b00560061006c00750065005d003d00540072007500 ,
                        0x6500000000000000000000000000000000000000000000010000000000000001 ,
                        0x00000000000000ffff6600170000005b0048004300320035005d002b005b0048 ,
                        0x004300350030005d002b005b00480043003100300030005d003d003000000000 ,
                        0x000000000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =375
                    Left =2580
                    Top =60
                    Width =780
                    TabIndex =1
                    Name ="Alive"
                    ControlSource ="Alive"
                    RowSourceType ="Value List"
                    RowSource ="-1;\"Yes\";0;\"No\""
                    ColumnWidths ="0;375"
                    BeforeUpdate ="[Event Procedure]"
                    DefaultValue ="-1"
                    OnGotFocus ="[Event Procedure]"

                End
                Begin ComboBox
                    OverlapFlags =247
                    TextFontFamily =0
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =6480
                    Left =60
                    Top =60
                    Width =2304
                    BackColor =65535
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    ConditionalFormat = Begin
                        0x0100000030010000030000000100000000000000000000001a00000001000000 ,
                        0x00000000ffffff0001000000000000001b000000420000000100000000000000 ,
                        0xffff0000010000000000000043000000670000000100000000000000ffffff00 ,
                        0x49004900660028004c0065006e0028005b005300700065006300690065007300 ,
                        0x5d0029003e0030002c0031002c00300029000000000049004900660028004900 ,
                        0x73004e0075006c006c0028005b0048004300320035005d002b005b0048004300 ,
                        0x350030005d002b005b00480043003100300030005d0029002c0030002c003100 ,
                        0x2900000000005b0050006100720065006e0074005d002e005b00630062007800 ,
                        0x4e006f005300610070006c0069006e00670073005d002e00560061006c007500 ,
                        0x65003d00540072007500650000000000
                    End
                    Name ="Species"
                    ControlSource ="Species"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT  q.Master_PLANT_Code, q.LU_Code, q.Utah_Species,   q.Lifeform  F"
                        "ROM qryU_Top_Canopy q WHERE ((q.[Utah_Species] Is Not Null)  AND (q.[Lifeform]='"
                        "Tree')) OR (q.[LU_Code] = 'JUNIPERUS')  ORDER BY q.LU_Code    UNION   (SELECT DI"
                        "STINCT  u.Unknown_Code, u.Unknown_Code,   u.Plant_Type + \" - \" + u.Plant_Descr"
                        "iption, u.Plant_Type AS Lifeform  FROM tbl_Unknown_Species u WHERE u.Plant_Type "
                        "IN ('Tree','Other')  OR u.Plant_Type IS NULL  ORDER BY u.Unknown_Code);"
                    ColumnWidths ="0;2160;4320"
                    BeforeUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    ConditionalFormat14 = Begin
                        0x01000300000001000000000000000100000000000000ffffff00190000004900 ,
                        0x4900660028004c0065006e0028005b0053007000650063006900650073005d00 ,
                        0x29003e0030002c0031002c003000290000000000000000000000000000000000 ,
                        0x000000000001000000000000000100000000000000ffff000026000000490049 ,
                        0x0066002800490073004e0075006c006c0028005b0048004300320035005d002b ,
                        0x005b0048004300350030005d002b005b00480043003100300030005d0029002c ,
                        0x0030002c00310029000000000000000000000000000000000000000000000100 ,
                        0x0000000000000100000000000000ffffff00230000005b005000610072006500 ,
                        0x6e0074005d002e005b006300620078004e006f005300610070006c0069006e00 ,
                        0x670073005d002e00560061006c00750065003d00540072007500650000000000 ,
                        0x0000000000000000000000000000000000
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =7740
                    Top =60
                    Width =705
                    Height =300
                    TabIndex =5
                    ForeColor =255
                    Name ="btnDelete"
                    Caption ="Delete"
                    OnClick ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"

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
                    Left =3480
                    Top =60
                    Width =606
                    Height =288
                    Name ="btnA1"
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
                    Left =4200
                    Top =60
                    Width =606
                    Height =288
                    TabIndex =1
                    Name ="btnA5"
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
                    Left =4920
                    Top =60
                    Width =606
                    Height =288
                    TabIndex =2
                    Name ="btnS1"
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
                    Left =5640
                    Top =60
                    Width =606
                    Height =288
                    TabIndex =3
                    Name ="btnS5"
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
                    Left =6360
                    Top =60
                    Width =606
                    Height =288
                    TabIndex =4
                    Name ="btnZero"
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
' MODULE:       Form_fsub_OT_Tree_Saplings
' Level:        Form module
' Version:      1.06
' Description:  data functions & procedures specific to overstory tree sapling monitoring
'
' Source/date:  Bonnie Campbell, 2/11/2016
' Revisions:    RDB - unknown  - 1.00 - initial version
'               BLC - 2/11/2016 - 1.01 - added documentation, set checkbox notifications for no species found
'               BLC - 3/8/2016 - 1.02 - added documentation, Species_GotFocus()
'               BLC - 3/29/2016 - 1.03 - added field highlighting
'               BLC - 4/13/2016 - 1.04 - added refresh for underlying subforms for conditional formatting
'               BLC - 8/9/2017  - 1.05 - revised to avoid clearing 0 values entered for diameter class totals,
'                                        added documentation, error handling & renamed button prefixes to btnXX
'               BLC - 2/2/2018 - 1.06 - added ToggleTallyButtons, SetTallyButtons, and HC25, HC50, HC100 got
'                                       focus events to avoid error #438 object doesn't support this
'                                       property/method & control when tally buttons are available
' =================================

' ---------------------------------
' SUB:          Form_Load
' Description:  handles form loading actions
' Parameters:
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 29, 2016 - for NCPN tools
' Revisions:
'       BLC, 3/29/2016 - initial version
'       BLC, 2/2/2018 - disable tally buttons unless diameter class field has focus
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler

    'default
    btnA1.Enabled = False
    btnA5.Enabled = False
    btnS1.Enabled = False
    btnS5.Enabled = False
    btnZero.Enabled = False

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[Form_fsub_OT_Tree_Saplings])"
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

    ' Make sure there is an events record
    If IsNull(Me.Parent!Start_Date) Then
      MsgBox "Missing site visit date."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
      GoTo Exit_Handler
    End If
    ' Create the GUID primary key value
    If IsNull(Me!TS_ID) Then
        If GetDataType("tbl_OT_Tree_Saplings", "TS_ID") = dbText Then
            Me.TS_ID = fxnGUIDGen
        End If
    End If

    '-----------------------------------
    ' update the NoDataCollected info
    '-----------------------------------
    Dim noData As Scripting.Dictionary
    
    'remove the no data collected record
    Set noData = SetNoDataCollected(Me.Parent.Form.Controls("Event_ID"), "E", "OverstoryTree-Sapling", 0)
        
    'update checkbox/rectangle
    Me.Parent.Form.Controls("cbxNoSaplings") = 0
    Me.Parent.Form.Controls("cbxNoSaplings").Enabled = False
    Me.Parent.Form.Controls("rctNoSaplings").Visible = False

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeInsert[Form_fsub_OT_Tree_Saplings])"
    End Select
    Resume Exit_Handler
End Sub

'==================================
'   HC25-50-100 Tally Toggling
'==================================

' ---------------------------------
' SUB:          HC25_GotFocus
' Description:  Handles actions when control has focus
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, February 2, 2018 - for NCPN tools
' Adapted:
' Revisions:
'   BLC, 2/2/2018  - initial version
' ---------------------------------
Private Sub HC25_GotFocus()
On Error GoTo Err_Handler

'    Debug.Print "HC25_GotFocus"

    SetTallyButtons HC25
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - HC25_GotFocus[Form_fsub_OT_Tree_Saplings])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          HC50_GotFocus
' Description:  Handles actions when control has focus
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, February 2, 2018 - for NCPN tools
' Adapted:
' Revisions:
'   BLC, 2/2/2018  - initial version
' ---------------------------------
Private Sub HC50_GotFocus()
On Error GoTo Err_Handler

'    Debug.Print "HC50_GotFocus"

    SetTallyButtons HC50

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - HC50_GotFocus[Form_fsub_OT_Tree_Saplings])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          HC100_GotFocus
' Description:  Handles actions when control has focus
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, February 2, 2018 - for NCPN tools
' Adapted:
' Revisions:
'   BLC, 2/2/2018  - initial version
' ---------------------------------
Private Sub HC100_GotFocus()
On Error GoTo Err_Handler
  
'    Debug.Print "HC100_GotFocus"

    SetTallyButtons HC100

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - HC100_GotFocus[Form_fsub_OT_Tree_Saplings])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          HC25_LostFocus
' Description:  Handles actions when control lost focus
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, February 2, 2018 - for NCPN tools
' Adapted:
' Revisions:
'   BLC, 2/2/2018  - initial version
' ---------------------------------
Private Sub HC25_LostFocus()
On Error GoTo Err_Handler

    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - HC25_LostFocus[Form_fsub_OT_Tree_Saplings])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          HC50_LostFocus
' Description:  Handles actions when control lost focus
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, February 2, 2018 - for NCPN tools
' Adapted:
' Revisions:
'   BLC, 2/2/2018  - initial version
' ---------------------------------
Private Sub HC50_LostFocus()
On Error GoTo Err_Handler


Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - HC50_LostFocus[Form_fsub_OT_Tree_Saplings])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          HC100_LostFocus
' Description:  Handles actions when control lost focus
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, February 2, 2018 - for NCPN tools
' Adapted:
' Revisions:
'   BLC, 2/2/2018  - initial version
' ---------------------------------
Private Sub HC100_LostFocus()
On Error GoTo Err_Handler
  

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - HC100_LostFocus[Form_fsub_OT_Tree_Saplings])"
    End Select
    Resume Exit_Handler
End Sub

'==================================
'   HC25-50-100 Highlighting
'==================================
' ---------------------------------
' SUB:          SetHCHighlight
' Description:  Handles HC highlighting
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, March 29, 2016 - for NCPN tools
' Adapted:
' Revisions:
'   BLC, 3/29/2016  - initial version
'   BLC, 8/9/2017   - revised to remove clearing action where
'                     user chooses to enter 0 per issue:
'                     https://github.com/NCPN/upland/issues/103
' ---------------------------------
Private Sub SetHCHighlighting()
On Error GoTo Err_Handler

    'clear HC25-50-100 values to get rid of 0 if not set
'    If Not Me.HC25 <> 0 Then Me.HC25 = Null
'    If Not Me.HC50 <> 0 Then Me.HC50 = Null
'    If Not Me.HC100 <> 0 Then Me.HC100 = Null

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetHCHighlighting[Form_fsub_OT_Tree_Saplings])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          HC25_Change
' Description:  Handles actions when control has been changed
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, March 29, 2016 - for NCPN tools
' Adapted:
' Revisions:
'   BLC, 3/29/2016  - initial version
' ---------------------------------
Private Sub HC25_Change()
On Error GoTo Err_Handler

    SetHCHighlighting

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - HC25_Change[Form_fsub_OT_Tree_Saplings])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          HC50_Change
' Description:  Handles actions when control has been changed
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, March 29, 2016 - for NCPN tools
' Adapted:
' Revisions:
'   BLC, 3/29/2016  - initial version
' ---------------------------------
Private Sub HC50_Change()
On Error GoTo Err_Handler

    SetHCHighlighting

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - HC50_Change[Form_fsub_OT_Tree_Saplings])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          HC100_Change
' Description:  Handles actions when control has been changed
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, March 29, 2016 - for NCPN tools
' Adapted:
' Revisions:
'   BLC, 3/29/2016  - initial version
' ---------------------------------
Private Sub HC100_Change()
On Error GoTo Err_Handler

    SetHCHighlighting

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - HC100_Change[Form_fsub_OT_Tree_Saplings])"
    End Select
    Resume Exit_Handler
End Sub

'==================================
'   Disable Tallies for Non-HC Control Foci
'==================================

' ---------------------------------
' SUB:          Alive_GotFocus
' Description:  Handles actions when control has focus
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, February 2, 2018 - for NCPN tools
' Adapted:
' Revisions:
'   BLC, 2/2/2018  - initial version
' ---------------------------------
Private Sub Alive_GotFocus()
On Error GoTo Err_Handler

    'disable all by passing non-diameter class control
    SetTallyButtons Me.Alive

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Alive_GotFocus[Form_fsub_OT_Tree_Saplings])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnDelete_GotFocus
' Description:  Handles actions when control has focus
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, February 2, 2018 - for NCPN tools
' Adapted:
' Revisions:
'   BLC, 2/2/2018  - initial version
' ---------------------------------
Private Sub btnDelete_GotFocus()
On Error GoTo Err_Handler

    'disable all by passing non-diameter class control
    SetTallyButtons Me.btnDelete

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnDelete_GotFocus[Form_fsub_OT_Tree_Saplings])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnMaster_GotFocus
' Description:  Handles actions when control has focus
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, February 2, 2018 - for NCPN tools
' Adapted:
' Revisions:
'   BLC, 2/2/2018  - initial version
' ---------------------------------
Private Sub btnMaster_GotFocus()
On Error GoTo Err_Handler

    'disable all by passing non-diameter class control
    SetTallyButtons Me.btnMaster

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnMaster_GotFocus[Form_fsub_OT_Tree_Saplings])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnUnknown_GotFocus
' Description:  Handles actions when control has focus
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, February 2, 2018 - for NCPN tools
' Adapted:
' Revisions:
'   BLC, 2/2/2018  - initial version
' ---------------------------------
Private Sub btnUnknown_GotFocus()
On Error GoTo Err_Handler

    'disable all by passing non-diameter class control
    SetTallyButtons Me.btnUnknown

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnUnknown_GotFocus[Form_fsub_OT_Tree_Saplings])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_Deactivate
' Description:  Handles actions when form deactivates
'               Deactivate vs. LostFocus event is used because the latter does not
'               fire when moving to the main form or another subform
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, February 2, 2018 - for NCPN tools
' Adapted:
' Revisions:
'   BLC, 2/2/2018  - initial version
' ---------------------------------
Private Sub Form_Deactivate()
On Error GoTo Err_Handler

    'disable all by passing non-diameter class control
    SetTallyButtons Me.Species

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Deactivate[Form_fsub_OT_Tree_Saplings])"
    End Select
    Resume Exit_Handler
End Sub

'==================================
'   Species Record Methods
'==================================

' ---------------------------------
' SUB:          Species_GotFocus
' Description:  Handles species actions when control has focus
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, March 8, 2016 - for NCPN tools
' Adapted:
' Revisions:
'   BLC, 3/8/2016  - initial version
'   BLC, 2/2/2018  - disable tally buttons
' ---------------------------------
Private Sub Species_GotFocus()
On Error GoTo Err_Handler

    'update the data to ensure new unknowns are added
    Me.ActiveControl.Requery

    'disable all by passing non-diameter class control
    SetTallyButtons Me.Species

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Species_GotFocus[Form_fsub_OT_Tree_Saplings])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Species_Change
' Description:  Handles species actions when control has been changed
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, March 29, 2016 - for NCPN tools
' Adapted:
' Revisions:
'   BLC, 3/29/2016  - initial version
' ---------------------------------
Private Sub Species_Change()
On Error GoTo Err_Handler

    SetHCHighlighting

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Species_Change[Form_fsub_OT_Tree_Saplings])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Species_BeforeUpdate
' Description:  Handles species actions before update
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, August 9, 2017 - for NCPN tools
' Adapted:
' Revisions:
'   RDB, unknown   - initial version
'   BLC, 8/9/2017  - added documentation, error handling
' ---------------------------------
Private Sub Species_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

    Dim Reply As Integer
    Dim TextMsg As String

    If Not IsNull(DLookup("[TS_ID]", "tbl_OT_Tree_Saplings", "[Event_ID] = '" & Me!Event_ID & "' AND [Species] = '" & Me!Species & "' AND [Alive] = " & Me!Alive)) Then
     If Me!Alive Then
       TextMsg = "This species already exists as alive on this point.  Would you like to set it to dead?"
     Else
       TextMsg = "This species already exists as dead on this point.  Would you like to set it to alive?"
     End If
     Reply = MsgBox(TextMsg, vbYesNo, "Species Verification")
     If Reply = vbYes Then
       Me!Alive = IIf(Me!Alive = True, False, True)
       If Not IsNull(DLookup("[TS_ID]", "tbl_OT_Tree_Saplings", "[Event_ID] = '" & Me!Event_ID & "' AND [Species] = '" & Me!Species & "' AND [Alive] = " & Me!Alive)) Then
         MsgBox "This species is already recorded for this point."
         DoCmd.CancelEvent
         SendKeys "{ESC}"
         Exit Sub
       End If
     Else
       DoCmd.CancelEvent
       SendKeys "{ESC}"
       Exit Sub
     End If
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Species_BeforeUpdate[Form_fsub_OT_Tree_Saplings])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Alive_BeforeUpdate
' Description:  Handles alive before update actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 11, 2016 - for NCPN tools
' Revisions:
'   JRB, 6/x/2006  - initial version
'   RDB, unknown   - ?
'   BLC, 2/11/2016  - added documentation
'   BLC, 8/9/0217   - added error handling
' ---------------------------------
Private Sub Alive_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

    If Not IsNull(DLookup("[TS_ID]", "tbl_OT_Tree_Saplings", "[Event_ID] = '" & Me!Event_ID & "' AND [Species] = '" & Me!Species & "' AND [Alive] = " & Me!Alive)) Then
      MsgBox "This species is already recorded for this transect."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
    End If
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Alive_BeforeUpdate[Form_fsub_OT_Tree_Saplings])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnDelete_Click
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
'   BLC, 4/13/2016 - added requery of related subform to clear/set conditional formatting on change
'   BLC, 8/9/2017  - renamed button to btnDelete
' ---------------------------------
Private Sub btnDelete_Click()
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
    
        Dim noData As Scripting.Dictionary
        
        'remove the no data collected record
        Set noData = SetNoDataCollected(Me.Parent.Form.Controls("Event_ID"), "E", "OverstoryTree-Sapling", 1)
    
        'update checkbox/rectangle
        Me.Parent.Form.Controls("cbxNoSaplings") = 1
        Me.Parent.Form.Controls("cbxNoSaplings").Enabled = True
        Me.Parent.Form.Controls("rctNoSaplings").Visible = True
        
        'refresh the subform to clear conditional formatting
        Me.Requery
        
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnDelete_Click[Form_fsub_OT_Tree_Saplings])"
    End Select
    Resume Exit_Handler
End Sub

'==================================
'   Tally Buttons
'==================================

' ---------------------------------
' SUB:          btnA1_Click
' Description:  Handles A1 button click actions (add 1 to value)
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, August 9, 2017 - for NCPN tools
' Adapted:
' Revisions:
'   RDB, unknown   - initial version
'   BLC, 8/9/2017  - added documentation, error handling
' ---------------------------------
Private Sub btnA1_Click()
On Error GoTo Err_Handler

  If Screen.PreviousControl.Name <> "Species" And Not IsNull(Me!Species) Then
    If IsNull(Screen.PreviousControl.Value) Then
      Screen.PreviousControl.Value = 1
    Else
      Screen.PreviousControl.Value = Screen.PreviousControl.Value + 1
    End If
  End If
  Screen.PreviousControl.SetFocus

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnA1_Click[Form_fsub_OT_Tree_Saplings])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnA5_Click
' Description:  Handles A5 button click actions (add 5 to value)
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, August 9, 2017 - for NCPN tools
' Adapted:
' Revisions:
'   RDB, unknown   - initial version
'   BLC, 8/9/2017  - added documentation, error handling
' ---------------------------------
Private Sub btnA5_Click()
On Error GoTo Err_Handler
  If Screen.PreviousControl.Name <> "Species" And Not IsNull(Me!Species) Then
    If IsNull(Screen.PreviousControl.Value) Then
      Screen.PreviousControl.Value = 5
    Else
      Screen.PreviousControl.Value = Screen.PreviousControl.Value + 5
    End If
  End If
  Screen.PreviousControl.SetFocus

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnA5_Click[Form_fsub_OT_Tree_Saplings])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnS5_Click
' Description:  Handles S5 button click actions (subtract 5 from value)
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, August 9, 2017 - for NCPN tools
' Adapted:
' Revisions:
'   RDB, unknown   - initial version
'   BLC, 8/9/2017  - added documentation, error handling
' ---------------------------------
Private Sub btnS1_Click()
On Error GoTo Err_Handler

  If Screen.PreviousControl.Name <> "Species" And Not IsNull(Me!Species) Then
    If IsNull(Screen.PreviousControl.Value) Then
      Screen.PreviousControl.Value = 0
    ElseIf Screen.PreviousControl.Value - 1 < 0 Then
      MsgBox "Total cannot be negative.", , "Belt Shrubs"
      Exit Sub
    Else
      Screen.PreviousControl.Value = Screen.PreviousControl.Value - 1
    End If
  End If
  Screen.PreviousControl.SetFocus

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnS1_Click[Form_fsub_OT_Tree_Saplings])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnS5_Click
' Description:  Handles A5 button click actions (subtract 5 from value)
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, August 9, 2017 - for NCPN tools
' Adapted:
' Revisions:
'   RDB, unknown   - initial version
'   BLC, 8/9/2017  - added documentation, error handling
' ---------------------------------
Private Sub btnS5_Click()
On Error GoTo Err_Handler

  If Screen.PreviousControl.Name <> "Species" And Not IsNull(Me!Species) Then
    If IsNull(Screen.PreviousControl.Value) Then
      Screen.PreviousControl.Value = 0
    ElseIf Screen.PreviousControl.Value - 5 < 0 Then
      MsgBox "Total cannot be negative.", , "Belt Shrubs"
      Exit Sub
    Else
      Screen.PreviousControl.Value = Screen.PreviousControl.Value - 5
    End If
  End If
  Screen.PreviousControl.SetFocus
  
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnS5_Click[Form_fsub_OT_Tree_Saplings])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnZero_Click
' Description:  Handles Zero button click actions (set value as 0)
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, August 9, 2017 - for NCPN tools
' Adapted:
' Revisions:
'   RDB, unknown   - initial version
'   BLC, 8/9/2017  - added documentation, error handling
' ---------------------------------
Private Sub btnZero_Click()
On Error GoTo Err_Handler

  If Screen.PreviousControl.Name <> "Species" And Not IsNull(Me!Species) Then
      Screen.PreviousControl.Value = 0
  End If
  Screen.PreviousControl.SetFocus
  
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnZero_Click[Form_fsub_OT_Tree_Saplings])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnUnknown_Click
' Description:  Handles Unknown button click actions (opens unknown form)
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, August 9, 2017 - for NCPN tools
' Adapted:
' Revisions:
'   RDB, unknown   - initial version
'   BLC, 8/9/2017  - added documentation, error handling
' ---------------------------------
Private Sub btnUnknown_Click()
On Error GoTo Err_Handler

    Dim stDocName As String
    Dim stLinkCriteria As String

    DoCmd.OpenForm "frm_List_Unknown", , , , , acDialog
    Me.Refresh
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnUnknown_Click[Form_fsub_OT_Tree_Saplings])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnMaster_Click
' Description:  Handles Master button click actions (opens master plants)
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, August 9, 2017 - for NCPN tools
' Adapted:
' Revisions:
'   RDB, unknown   - initial version
'   BLC, 8/9/2017  - added documentation, error handling
' ---------------------------------
Private Sub btnMaster_Click()
On Error GoTo Err_Handler

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Master_Species"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnMaster_Click[Form_fsub_OT_Tree_Saplings])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          ToggleTallyButtons
' Description:  Enables or disables tally buttons based on current state
' Assumptions:  -
' Parameters:   btn - tally button to set (command button object, optional)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, February 2, 2018 - for NCPN tools
' Adapted:
' Revisions:
'   BLC, 2/2/2018  - initial version
' ---------------------------------
Private Sub ToggleTallyButtons(Optional btn As CommandButton)
On Error GoTo Err_Handler

    'toggle single tally button
    If Not btn Is Nothing Then
        btn.Enabled = Not btn.Enabled
        GoTo Exit_Handler
    End If
    
    'toggle all tally buttons
    btnA1.Enabled = Not btnA1.Enabled
    btnA5.Enabled = Not btnA5.Enabled
    btnS1.Enabled = Not btnS1.Enabled
    btnA5.Enabled = Not btnS5.Enabled
    btnZero.Enabled = Not btnZero.Enabled
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ToggleTallyButtons[Form_fsub_OT_Tree_Saplings])"
    End Select
    Resume Exit_Handler
End Sub
    
' ---------------------------------
' SUB:          SetTallyButtons
' Description:  Enables or disables tally buttons based on textbox value
' Assumptions:  1) If run from controls other than HC25, HC50, HC100 all tally
'                  buttons should be disabled
'               2) No other controls on the form begin with "HC"
'               3) Lost focus events for HC25/HC50/HC100 trigger this subroutine
'                  to disable tally buttons when the focus is NOT HC25, HC50, HC100
' Parameters:   tbx - control to evaluate (Textbox object)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, February 2, 2018 - for NCPN tools
' Adapted:
' Revisions:
'   BLC, 2/2/2018  - initial version
' ---------------------------------
Private Sub SetTallyButtons(ctrl As Control)
On Error GoTo Err_Handler

    'ensure the control is HC25, HC50, HC100
    If ctrl.ControlType <> acTextBox Or _
       (ctrl.ControlType = acTextBox And _
        Left(ctrl.Name, 2) <> "HC") Then
        
        'no buttons are viable (+5, +1, 0, -1, -5 disabled)
        If btnA1.Enabled = True Then ToggleTallyButtons btnA1
        If btnA5.Enabled = True Then ToggleTallyButtons btnA5
        If btnZero.Enabled = True Then ToggleTallyButtons btnZero
        If btnS1.Enabled = True Then ToggleTallyButtons btnS1
        If btnS5.Enabled = True Then ToggleTallyButtons btnS5
        
        GoTo Exit_Handler
    End If

    Select Case ctrl 'HC25, HC50, HC100
        Case Is > 5, Is = 5
            'all buttons are viable (+5, +1, 0, -1, -5)
            If btnA1.Enabled = False Then ToggleTallyButtons btnA1
            If btnA5.Enabled = False Then ToggleTallyButtons btnA5
            If btnZero.Enabled = False Then ToggleTallyButtons btnZero
            If btnS1.Enabled = False Then ToggleTallyButtons btnS1
            If btnS5.Enabled = False Then ToggleTallyButtons btnS5
        Case Is > 1, Is = 1
            'all buttons except -5 are viable (+5, +1, 0, -1)
            If btnA1.Enabled = False Then ToggleTallyButtons btnA1
            If btnA5.Enabled = False Then ToggleTallyButtons btnA5
            If btnZero.Enabled = False Then ToggleTallyButtons btnZero
            If btnS1.Enabled = False Then ToggleTallyButtons btnS1
            If btnS5.Enabled = True Then ToggleTallyButtons btnS5
        Case Else 'includes Is = 0, Is < 0
            'only add/zero buttons are viable (+5, +1, 0)
            If btnA1.Enabled = False Then ToggleTallyButtons btnA1
            If btnA5.Enabled = False Then ToggleTallyButtons btnA5
            If btnZero.Enabled = False Then ToggleTallyButtons btnZero
            If btnS1.Enabled = True Then ToggleTallyButtons btnS1
            If btnS5.Enabled = True Then ToggleTallyButtons btnS5
    End Select
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetTallyButtons[Form_fsub_OT_Tree_Saplings])"
    End Select
    Resume Exit_Handler
End Sub
