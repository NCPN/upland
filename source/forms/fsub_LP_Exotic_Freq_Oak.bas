﻿Version =21
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10440
    DatasheetFontHeight =9
    ItemSuffix =57
    Left =2340
    Top =1350
    Right =10290
    Bottom =5595
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x69259af5aed1e340
    End
    RecordSource ="tbl_LP_Exotic_Freq"
    Caption ="fsub_LP_Exotic_Freq_Oak"
    BeforeInsert ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
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
            Height =780
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =240
                    Top =540
                    Width =1335
                    Height =240
                    FontWeight =700
                    Name ="Species_Label"
                    Caption ="Species Name"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =5820
                    Top =540
                    Width =360
                    Height =240
                    FontWeight =700
                    Name ="Label38"
                    Caption ="18"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =2580
                    Top =540
                    Width =360
                    Height =240
                    FontWeight =700
                    Name ="Label39"
                    Caption ="0"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =2940
                    Top =540
                    Width =360
                    Height =240
                    FontWeight =700
                    Name ="Label40"
                    Caption ="2"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =3300
                    Top =540
                    Width =360
                    Height =240
                    FontWeight =700
                    Name ="Label41"
                    Caption ="4"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =3660
                    Top =540
                    Width =360
                    Height =240
                    FontWeight =700
                    Name ="Label42"
                    Caption ="6"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =4020
                    Top =540
                    Width =360
                    Height =240
                    FontWeight =700
                    Name ="Label43"
                    Caption ="8"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =4380
                    Top =540
                    Width =360
                    Height =240
                    FontWeight =700
                    Name ="Label44"
                    Caption ="10"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =4740
                    Top =540
                    Width =360
                    Height =240
                    FontWeight =700
                    Name ="Label45"
                    Caption ="12"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =5100
                    Top =540
                    Width =360
                    Height =240
                    FontWeight =700
                    Name ="Label46"
                    Caption ="14"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =5460
                    Top =540
                    Width =360
                    Height =240
                    FontWeight =700
                    Name ="Label47"
                    Caption ="16"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    Left =2580
                    Top =300
                    Width =3600
                    Height =240
                    FontWeight =700
                    Name ="Label48"
                    Caption ="Meter"
                    Tag ="DetachedLabel"
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =6300
                    Top =480
                    Width =839
                    Height =300
                    Name ="ButtonMaster"
                    Caption ="Master "
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =6300
                    Top =60
                    Width =840
                    Height =300
                    TabIndex =1
                    Name ="ButtonUnknown"
                    Caption ="Unknown"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =600
                    Top =120
                    Width =600
                    ColumnOrder =0
                    TabIndex =2
                    Name ="tbxRecordCount"

                    LayoutCachedLeft =600
                    LayoutCachedTop =120
                    LayoutCachedWidth =1200
                    LayoutCachedHeight =360
                End
            End
        End
        Begin Section
            Height =420
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =2580
                    Width =3660
                    Height =300
                    BackColor =65535
                    Name ="tbxRctVisible"
                    ConditionalFormat = Begin
                        0x0100000064020000030000000100000000000000000000006e00000001000000 ,
                        0x00000000ffff000001000000000000006f000000dd0000000100000000000000 ,
                        0xffffff000100000000000000de000000010100000100000000000000ffffff00 ,
                        0x41006200730028005b004d0030005d0029002b0041006200730028005b004d00 ,
                        0x35005d0029002b0041006200730028005b004d00310030005d0029002b004100 ,
                        0x6200730028005b004d00310035005d0029002b0041006200730028005b004d00 ,
                        0x320030005d0029002b0041006200730028005b004d00320035005d0029002b00 ,
                        0x41006200730028005b004d00330030005d0029002b0041006200730028005b00 ,
                        0x4d00330035005d0029002b0041006200730028005b004d00340030005d002900 ,
                        0x2b0041006200730028005b004d00340035005d0029003d003000000000004100 ,
                        0x6200730028005b004d0030005d0029002b0041006200730028005b004d003500 ,
                        0x5d0029002b0041006200730028005b004d00310030005d0029002b0041006200 ,
                        0x730028005b004d00310035005d0029002b0041006200730028005b004d003200 ,
                        0x30005d0029002b0041006200730028005b004d00320035005d0029002b004100 ,
                        0x6200730028005b004d00330030005d0029002b0041006200730028005b004d00 ,
                        0x330035005d0029002b0041006200730028005b004d00340030005d0029002b00 ,
                        0x41006200730028005b004d00340035005d0029003e003000000000005b005000 ,
                        0x6100720065006e0074005d002e005b006300620078004e006f00450078006f00 ,
                        0x74006900630073005d002e00560061006c00750065003d005400720075006500 ,
                        0x00000000
                    End

                    LayoutCachedLeft =2580
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =300
                    ConditionalFormat14 = Begin
                        0x01000300000001000000000000000100000000000000ffff00006d0000004100 ,
                        0x6200730028005b004d0030005d0029002b0041006200730028005b004d003500 ,
                        0x5d0029002b0041006200730028005b004d00310030005d0029002b0041006200 ,
                        0x730028005b004d00310035005d0029002b0041006200730028005b004d003200 ,
                        0x30005d0029002b0041006200730028005b004d00320035005d0029002b004100 ,
                        0x6200730028005b004d00330030005d0029002b0041006200730028005b004d00 ,
                        0x330035005d0029002b0041006200730028005b004d00340030005d0029002b00 ,
                        0x41006200730028005b004d00340035005d0029003d0030000000000000000000 ,
                        0x0000000000000000000000000001000000000000000100000000000000ffffff ,
                        0x006d00000041006200730028005b004d0030005d0029002b0041006200730028 ,
                        0x005b004d0035005d0029002b0041006200730028005b004d00310030005d0029 ,
                        0x002b0041006200730028005b004d00310035005d0029002b0041006200730028 ,
                        0x005b004d00320030005d0029002b0041006200730028005b004d00320035005d ,
                        0x0029002b0041006200730028005b004d00330030005d0029002b004100620073 ,
                        0x0028005b004d00330035005d0029002b0041006200730028005b004d00340030 ,
                        0x005d0029002b0041006200730028005b004d00340035005d0029003e00300000 ,
                        0x0000000000000000000000000000000000000000010000000000000001000000 ,
                        0x00000000ffffff00220000005b0050006100720065006e0074005d002e005b00 ,
                        0x6300620078004e006f00450078006f0074006900630073005d002e0056006100 ,
                        0x6c00750065003d00540072007500650000000000000000000000000000000000 ,
                        0x0000000000
                    End
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
                    TabIndex =13
                    Name ="Shrub_ID"
                    ControlSource ="Exotic_ID"
                    StatusBarText ="Unique record identifier - primary key"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =300
                    Top =60
                    Width =300
                    Height =255
                    ColumnWidth =2310
                    TabIndex =14
                    Name ="Transect_ID"
                    ControlSource ="Transect_ID"
                    StatusBarText ="Foreign key to tbl_Canopy_Transect"

                End
                Begin ComboBox
                    OverlapFlags =247
                    TextFontFamily =0
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =6480
                    Left =180
                    Top =60
                    Width =2304
                    TabIndex =1
                    BackColor =65535
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    ConditionalFormat = Begin
                        0x01000000ce010000030000000100000000000000000000001a00000001000000 ,
                        0x00000000ffffff0001000000000000001b0000003e0000000100000000000000 ,
                        0xffffff0001000000000000003f000000b60000000100000000000000ffff0000 ,
                        0x49004900660028004c0065006e0028005b005300700065006300690065007300 ,
                        0x5d0029003e0030002c0031002c0030002900000000005b005000610072006500 ,
                        0x6e0074005d002e005b006300620078004e006f00450078006f00740069006300 ,
                        0x73005d002e00560061006c00750065003d005400720075006500000000004900 ,
                        0x49006600280041006200730028005b004d0030005d0029002b00410062007300 ,
                        0x28005b004d0035005d0029002b0041006200730028005b004d00310030005d00 ,
                        0x29002b0041006200730028005b004d00310035005d0029002b00410062007300 ,
                        0x28005b004d00320030005d0029002b0041006200730028005b004d0032003500 ,
                        0x5d0029002b0041006200730028005b004d00330030005d0029002b0041006200 ,
                        0x730028005b004d00330035005d0029002b0041006200730028005b004d003400 ,
                        0x30005d0029002b0041006200730028005b004d00340035005d0029003e003000 ,
                        0x2c0031002c003000290000000000
                    End
                    Name ="Species"
                    ControlSource ="Species"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT * FROM (SELECT DISTINCT qryU_Top_Canopy.Master_PLANT_Code, qryU_Top_Canop"
                        "y.LU_Code AS LUcode, qryU_Top_Canopy.Utah_Species,   qryU_Top_Canopy.Nativity  F"
                        "ROM qryU_Top_Canopy  WHERE (((qryU_Top_Canopy.Utah_Species) Is Not Null) AND ((q"
                        "ryU_Top_Canopy.[Nativity])='NonNative')) )   UNION  (SELECT DISTINCT tbl_Unknown"
                        "_Species.Unknown_Code, tbl_Unknown_Species.Unknown_Code AS LUcode,   tbl_Unknown"
                        "_Species.Plant_Type+ \" - \" + tbl_Unknown_Species.Plant_Description, NULL AS Na"
                        "tivity  FROM tbl_Unknown_Species  ) ORDER BY LUcode;"
                    ColumnWidths ="0;2160;4320"
                    BeforeUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"
                    ConditionalFormat14 = Begin
                        0x01000300000001000000000000000100000000000000ffffff00190000004900 ,
                        0x4900660028004c0065006e0028005b0053007000650063006900650073005d00 ,
                        0x29003e0030002c0031002c003000290000000000000000000000000000000000 ,
                        0x000000000001000000000000000100000000000000ffffff00220000005b0050 ,
                        0x006100720065006e0074005d002e005b006300620078004e006f00450078006f ,
                        0x0074006900630073005d002e00560061006c00750065003d0054007200750065 ,
                        0x0000000000000000000000000000000000000000000001000000000000000100 ,
                        0x000000000000ffff000076000000490049006600280041006200730028005b00 ,
                        0x4d0030005d0029002b0041006200730028005b004d0035005d0029002b004100 ,
                        0x6200730028005b004d00310030005d0029002b0041006200730028005b004d00 ,
                        0x310035005d0029002b0041006200730028005b004d00320030005d0029002b00 ,
                        0x41006200730028005b004d00320035005d0029002b0041006200730028005b00 ,
                        0x4d00330030005d0029002b0041006200730028005b004d00330035005d002900 ,
                        0x2b0041006200730028005b004d00340030005d0029002b004100620073002800 ,
                        0x5b004d00340035005d0029003e0030002c0031002c0030002900000000000000 ,
                        0x000000000000000000000000000000
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =6300
                    Top =60
                    Width =705
                    Height =300
                    TabIndex =12
                    ForeColor =255
                    Name ="ButtonDelete"
                    Caption ="Delete"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =2700
                    Top =60
                    TabIndex =2
                    Name ="M0"
                    ControlSource ="Oak0"
                    StatusBarText ="Zero point quadrat - species detected checkbox"

                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =3060
                    Top =60
                    TabIndex =3
                    Name ="M5"
                    ControlSource ="Oak2"
                    StatusBarText ="5 meter quadrat"

                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =3420
                    Top =60
                    TabIndex =4
                    Name ="M10"
                    ControlSource ="Oak4"
                    StatusBarText ="10 meter quadrat"

                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =3780
                    Top =60
                    TabIndex =5
                    Name ="M15"
                    ControlSource ="Oak6"
                    StatusBarText ="15 meter quadrat"

                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =4140
                    Top =60
                    TabIndex =6
                    Name ="M20"
                    ControlSource ="Oak8"
                    StatusBarText ="20 meter quadrat"

                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =4500
                    Top =60
                    TabIndex =7
                    Name ="M25"
                    ControlSource ="Oak10"
                    StatusBarText ="25 meter quadrat"

                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =4860
                    Top =60
                    TabIndex =8
                    Name ="M30"
                    ControlSource ="Oak12"
                    StatusBarText ="30 meter quadrat"

                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =5220
                    Top =60
                    TabIndex =9
                    Name ="M35"
                    ControlSource ="Oak14"
                    StatusBarText ="35 meter quadrat"

                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =5580
                    Top =60
                    TabIndex =10
                    Name ="M40"
                    ControlSource ="Oak16"
                    StatusBarText ="40 meter quadrat"

                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =5940
                    Top =60
                    TabIndex =11
                    Name ="M45"
                    ControlSource ="Oak18"
                    StatusBarText ="45 meter quadrat"

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
Option Explicit

' =================================
' MODULE:       Form_fsub_Exotic_Freq_Oak
' Level:        Form module
' Version:      1.03
' Description:  data functions & procedures specific to oak exotic frequency monitoring
'
' Source/date:  Bonnie Campbell, 2/2/2016
' Revisions:    RDB - unknown  - 1.00 - initial version
'               BLC - 2/2/2016 - 1.01 - added documentation, checkbox for no species found
'               BLC - 3/23/2016 -1.02 - updated button delete click to handle no exotics label not displaying
'                                       after deleting last record
'               BLC - 4/13/2016 - 1.03 - added refresh for underlying subforms for conditional formatting
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
' Adapted:      Bonnie Campbell, February 2, 2016 - for NCPN tools
' Revisions:
'   BLC, 2/2/2016  - initial version
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler
    
    tbxRecordCount.Value = Me.RecordsetClone.RecordCount

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[Form_fsub_LP_Exotic_Freq_Oak])"
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

    ' Create the GUID primary key value
    If IsNull(Me!Exotic_ID) Then
        If GetDataType("tbl_LP_Exotic_Freq", "Exotic_ID") = dbText Then
            Me.Exotic_ID = fxnGUIDGen
        End If
    End If
    
    '-----------------------------------
    ' update the NoDataCollected info
    '-----------------------------------
    Dim noData As Scripting.Dictionary
    
    'remove the no data collected record
    Set noData = SetNoDataCollected(Me.Parent!Transect_ID, "T", "1mBelt-Exotics", 0)
        
    'update checkbox/rectangle
    Me.Parent.Form.Controls("cbxNoExotics") = 0
    Me.Parent.Form.Controls("cbxNoExotics").Enabled = False
    Me.Parent.Form.Controls("rctNoExotics").Visible = False

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeInsert[Form_fsub_LP_Exotic_Freq_Oak])"
    End Select
    Resume Exit_Handler
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

Private Sub Button_Master_Species_Click()
On Error GoTo Err_Button_Master_Species_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim strOpenArg As String

    strOpenArg = "fsub_LP_Exotic_Frequency"
    stDocName = "frm_Master_Species"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , , strOpenArg

Exit_Button_Master_Species_Click:
    Exit Sub

Err_Button_Master_Species_Click:
    MsgBox Err.Description
    Resume Exit_Button_Master_Species_Click
 
End Sub

' ---------------------------------
' SUB:          Species_BeforeUpdate
' Description:  Handles species pre-update actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 2, 2016 - for NCPN tools
' Revisions:
'   RDB, unknown  - initial version
'   BLC, 2/2/2016  - added documentation, disable checkbox if species exist
'   BLC, 3/29/2016 - removed no data collected changes as this is taken care of with Form_BeforeInsert()
' ---------------------------------
Private Sub Species_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

    If Not IsNull(DLookup("[Exotic_ID]", "tbl_LP_Exotic_Freq", "[Transect_ID] = '" & Me!Transect_ID & "' AND [Species] = '" & Me!Species & "'")) Then
      MsgBox "This species is already recorded for this transect."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
      Me.Undo
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Species_BeforeUpdate[Form_fsub_LP_Exotic_Freq_Oak])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub ButtonUnknown_Click()
On Error GoTo Err_ButtonUnknown_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    DoCmd.OpenForm "frm_List_Unknown", , , , , acDialog
    Me!Species.Requery
    Me.Refresh

Exit_ButtonUnknown_Click:
    Exit Sub

Err_ButtonUnknown_Click:
    MsgBox Err.Description
    Resume Exit_ButtonUnknown_Click
    
End Sub

' ---------------------------------
' SUB:          Species_GotFocus
' Description:  Handles species actions when control has focus
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Russ DenBleyker, unknown
' Adapted:      Bonnie Campbell, February 9, 2016 - for NCPN tools
' Revisions:
'   RDB, unknown  - initial version
'   BLC, 2/9/2016 - added error handling, documentation, refresh list to catch unknowns
' ---------------------------------
Private Sub Species_GotFocus()
On Error GoTo Err_Handler

    If IsNull(Me.Parent!Visit_Date) Then    ' If they didn't bother to enter a date, default to event date.
      Me.Parent!Visit_Date = Me.Parent.Parent!Start_Date
      Me.Parent.Refresh   ' Force save of transect record
    End If

    'update the data to ensure new unknowns are added
    Me.ActiveControl.Requery

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Species_GotFocus[Form_fsub_LP_Exotic_Freq_Oak])"
    End Select
    Resume Exit_Handler
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
'   BLC, 3/23/2016 - added lblNoExotics.Visible since it was not properly appearing after deleting
'                    last record
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
    
        Dim noData As Scripting.Dictionary
        
        'remove the no data collected record
        Set noData = SetNoDataCollected(Me.Parent.Form.Controls("Transect_ID"), "T", "1mBelt-Exotics", 1)
    
        'update checkbox/rectangle
        Me.Parent.Form.Controls("cbxNoExotics") = 1
        Me.Parent.Form.Controls("cbxNoExotics").Enabled = True
        Me.Parent.Form.Controls("lblNoExotics").Visible = True
        Me.Parent.Form.Controls("rctNoExotics").Visible = True
        
        'refresh the subform to clear conditional formatting
        Me.Requery
        
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ButtonDelete_Click[Form_fsub_LP_Exotic_Freq_Oak])"
    End Select
    Resume Exit_Handler
End Sub
