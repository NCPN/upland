Version =20
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
    Top =336
    Right =11136
    Bottom =3792
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
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
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
                    OverlapFlags =95
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
                    Name ="ButtonMaster"
                    Caption ="Master Lookup"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =7500
                    LayoutCachedTop =120
                    LayoutCachedWidth =9045
                    LayoutCachedHeight =420
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7500
                    Top =540
                    Width =1545
                    Height =300
                    TabIndex =1
                    Name ="ButtonUnknown"
                    Caption ="Unknown Species"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =7500
                    LayoutCachedTop =540
                    LayoutCachedWidth =9045
                    LayoutCachedHeight =840
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
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
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =360
                    Height =255
                    ColumnWidth =2310
                    Name ="Shrub_ID"
                    ControlSource ="TS_ID"
                    StatusBarText ="Unique record identifier - primary key"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =660
                    Top =60
                    Width =300
                    Height =255
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Transect_ID"
                    ControlSource ="Event_ID"
                    StatusBarText ="Foreign key to tbl_Canopy_Transect"

                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =3855
                    Top =60
                    Width =540
                    Height =255
                    ColumnWidth =600
                    TabIndex =4
                    BackColor =65535
                    Name ="HC25"
                    ControlSource ="D25"
                    StatusBarText ="10.1-25cm height class total"
                    DefaultValue ="Null"
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
                        0x01000300000001000000000000000100000000000000ffffff00170000004900 ,
                        0x490066002800490073004e0075006c006c0028005b0048004300320035005d00 ,
                        0x29002c0030002c00310029000000000000000000000000000000000000000000 ,
                        0x0001000000000000000100000000000000ffff00001700000049004900660028 ,
                        0x00490073004e0075006c006c0028005b0048004300320035005d0029002c0031 ,
                        0x002c003000290000000000000000000000000000000000000000000001000000 ,
                        0x000000000100000000000000ffffff00250000005b0050006100720065006e00 ,
                        0x74005d002e005b006300620078004e006f005300610070006c0069006e006700 ,
                        0x73005d002e005b00560061006c00750065005d003d0054007200750065000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =4995
                    Top =60
                    Width =540
                    Height =255
                    ColumnWidth =600
                    TabIndex =5
                    BackColor =65535
                    Name ="HC50"
                    ControlSource ="D51"
                    StatusBarText ="25.1-50cm height class total"
                    DefaultValue ="Null"
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
                        0x01000300000001000000000000000100000000000000ffffff00170000004900 ,
                        0x490066002800490073004e0075006c006c0028005b0048004300350030005d00 ,
                        0x29002c0030002c00310029000000000000000000000000000000000000000000 ,
                        0x0001000000000000000100000000000000ffff00001700000049004900660028 ,
                        0x00490073004e0075006c006c0028005b0048004300350030005d0029002c0031 ,
                        0x002c003000290000000000000000000000000000000000000000000001000000 ,
                        0x000000000100000000000000ffffff00250000005b0050006100720065006e00 ,
                        0x74005d002e005b006300620078004e006f005300610070006c0069006e006700 ,
                        0x73005d002e005b00560061006c00750065005d003d0054007200750065000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =6015
                    Top =60
                    Width =540
                    Height =255
                    ColumnWidth =600
                    TabIndex =6
                    BackColor =65535
                    Name ="HC100"
                    ControlSource ="D101"
                    StatusBarText ="50.1-100cm height class total"
                    DefaultValue ="Null"
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
                        0x01000300000001000000000000000100000000000000ffffff00180000004900 ,
                        0x490066002800490073004e0075006c006c0028005b0048004300310030003000 ,
                        0x5d0029002c0030002c0031002900000000000000000000000000000000000000 ,
                        0x00000001000000000000000100000000000000ffff0000180000004900490066 ,
                        0x002800490073004e0075006c006c0028005b00480043003100300030005d0029 ,
                        0x002c0031002c0030002900000000000000000000000000000000000000000000 ,
                        0x01000000000000000100000000000000ffffff00250000005b00500061007200 ,
                        0x65006e0074005d002e005b006300620078004e006f005300610070006c006900 ,
                        0x6e00670073005d002e005b00560061006c00750065005d003d00540072007500 ,
                        0x6500000000000000000000000000000000000000000000
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
                    TabIndex =3
                    Name ="Alive"
                    ControlSource ="Alive"
                    RowSourceType ="Value List"
                    RowSource ="-1;\"Yes\";0;\"No\""
                    ColumnWidths ="0;375"
                    BeforeUpdate ="[Event Procedure]"
                    DefaultValue ="-1"

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
                    TabIndex =2
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
                    RowSource ="SELECT DISTINCT qryU_Top_Canopy.Master_PLANT_Code, qryU_Top_Canopy.LU_Code, qryU"
                        "_Top_Canopy.Utah_Species,   qryU_Top_Canopy.Lifeform FROM qryU_Top_Canopy WHERE "
                        "(((qryU_Top_Canopy.Utah_Species) Is Not Null) AND ((qryU_Top_Canopy.[Lifeform])="
                        "'Tree')) ORDER BY qryU_Top_Canopy.LU_Code  UNION  (SELECT DISTINCT tbl_Unknown_S"
                        "pecies.Unknown_Code, tbl_Unknown_Species.Unknown_Code,   tbl_Unknown_Species.Pla"
                        "nt_Type + \" - \" + tbl_Unknown_Species.Plant_Description, tbl_Unknown_Species.P"
                        "lant_Type AS Lifeform FROM tbl_Unknown_Species WHERE tbl_Unknown_Species.Plant_T"
                        "ype IN ('Tree','Other') OR tbl_Unknown_Species.Plant_Type IS NULL ORDER BY tbl_U"
                        "nknown_Species.Unknown_Code);"
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
                    TabIndex =7
                    ForeColor =255
                    Name ="ButtonDelete"
                    Caption ="Delete"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
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
                    Name ="ButtonA1"
                    Caption ="+ 1"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4200
                    Top =60
                    Width =606
                    Height =288
                    TabIndex =1
                    Name ="ButtonA5"
                    Caption ="+ 5"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4920
                    Top =60
                    Width =606
                    Height =288
                    TabIndex =2
                    Name ="ButtonS1"
                    Caption ="- 1"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5640
                    Top =60
                    Width =606
                    Height =288
                    TabIndex =3
                    Name ="ButtonS5"
                    Caption ="- 5"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6360
                    Top =60
                    Width =606
                    Height =288
                    TabIndex =4
                    Name ="ButtonZero"
                    Caption ="0"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
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
' Version:      1.04
' Description:  data functions & procedures specific to overstory tree sapling monitoring
'
' Source/date:  Bonnie Campbell, 2/11/2016
' Revisions:    RDB - unknown  - 1.00 - initial version
'               BLC - 2/11/2016 - 1.01 - added documentation, set checkbox notifications for no species found
'               BLC - 3/8/2016 - 1.02 - added documentation, Species_GotFocus()
'               BLC - 3/29/2016 - 1.03 - added field highlighting
'               BLC - 4/13/2016 - 1.04 - added refresh for underlying subforms for conditional formatting
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
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler

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
    Dim NoData As Scripting.Dictionary
    
    'remove the no data collected record
    Set NoData = SetNoDataCollected(Me.Parent.Form.Controls("Event_ID"), "E", "OverstoryTree-Sapling", 0)
        
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
' ---------------------------------
Private Sub Species_GotFocus()
On Error GoTo Err_Handler

    'update the data to ensure new unknowns are added
    Me.ActiveControl.Requery

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
' ---------------------------------
Private Sub SetHCHighlighting()
On Error GoTo Err_Handler

    'clear HC25-50-100 values to get rid of 0 if not set
    If Not Me.HC25 <> 0 Then Me.HC25 = Null
    If Not Me.HC50 <> 0 Then Me.HC50 = Null
    If Not Me.HC100 <> 0 Then Me.HC100 = Null

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
' ---------------------------------
Private Sub Alive_BeforeUpdate(Cancel As Integer)
    If Not IsNull(DLookup("[TS_ID]", "tbl_OT_Tree_Saplings", "[Event_ID] = '" & Me!Event_ID & "' AND [Species] = '" & Me!Species & "' AND [Alive] = " & Me!Alive)) Then
      MsgBox "This species is already recorded for this transect."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
    End If
End Sub

Private Sub ButtonA1_Click()

  If Screen.PreviousControl.Name <> "Species" And Not IsNull(Me!Species) Then
    If IsNull(Screen.PreviousControl.value) Then
      Screen.PreviousControl.value = 1
    Else
      Screen.PreviousControl.value = Screen.PreviousControl.value + 1
    End If
  End If
  Screen.PreviousControl.SetFocus
End Sub

Private Sub ButtonA5_Click()
  If Screen.PreviousControl.Name <> "Species" And Not IsNull(Me!Species) Then
    If IsNull(Screen.PreviousControl.value) Then
      Screen.PreviousControl.value = 5
    Else
      Screen.PreviousControl.value = Screen.PreviousControl.value + 5
    End If
  End If
  Screen.PreviousControl.SetFocus
End Sub

Private Sub ButtonS1_Click()
  If Screen.PreviousControl.Name <> "Species" And Not IsNull(Me!Species) Then
    If IsNull(Screen.PreviousControl.value) Then
      Screen.PreviousControl.value = 0
    ElseIf Screen.PreviousControl.value - 1 < 0 Then
      MsgBox "Total cannot be negative.", , "Belt Shrubs"
      Exit Sub
    Else
      Screen.PreviousControl.value = Screen.PreviousControl.value - 1
    End If
  End If
  Screen.PreviousControl.SetFocus
End Sub

Private Sub ButtonS5_Click()
  If Screen.PreviousControl.Name <> "Species" And Not IsNull(Me!Species) Then
    If IsNull(Screen.PreviousControl.value) Then
      Screen.PreviousControl.value = 0
    ElseIf Screen.PreviousControl.value - 5 < 0 Then
      MsgBox "Total cannot be negative.", , "Belt Shrubs"
      Exit Sub
    Else
      Screen.PreviousControl.value = Screen.PreviousControl.value - 5
    End If
  End If
  Screen.PreviousControl.SetFocus
End Sub

Private Sub ButtonUnknown_Click()
    Dim stDocName As String
    Dim stLinkCriteria As String

    DoCmd.OpenForm "frm_List_Unknown", , , , , acDialog
    Me.Refresh
End Sub

Private Sub ButtonZero_Click()
  If Screen.PreviousControl.Name <> "Species" And Not IsNull(Me!Species) Then
      Screen.PreviousControl.value = 0
  End If
  Screen.PreviousControl.SetFocus
End Sub

Private Sub Species_BeforeUpdate(Cancel As Integer)
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


End Sub

Private Sub ButtonMaster_Click()
On Error GoTo Err_ButtonMaster_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Master_Species"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_ButtonMaster_Click:
    Exit Sub

Err_ButtonMaster_Click:
    MsgBox Err.Description
    Resume Exit_ButtonMaster_Click
    
End Sub

' ---------------------------------
' SUB:          ButtonDelete_Click
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
' ---------------------------------
Private Sub ButtonDelete_Click()
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
    
        Dim NoData As Scripting.Dictionary
        
        'remove the no data collected record
        Set NoData = SetNoDataCollected(Me.Parent.Form.Controls("Event_ID"), "E", "OverstoryTree-Sapling", 1)
    
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
            "Error encountered (#" & Err.Number & " - ButtonDelete_Click[Form_fsub_LP_Belt_Shrub])"
    End Select
    Resume Exit_Handler
End Sub
