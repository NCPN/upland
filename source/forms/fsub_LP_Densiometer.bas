Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    TabularFamily =124
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5760
    DatasheetFontHeight =9
    ItemSuffix =16
    Left =420
    Top =1290
    Right =6465
    Bottom =3855
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xf4627643ab16e540
    End
    RecordSource ="qry_LP_Densiometer"
    Caption ="fsub_LP_Densiometer"
    OnCurrent ="[Event Procedure]"
    BeforeInsert ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
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
            Height =480
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =360
                    Top =60
                    Width =540
                    Height =240
                    FontWeight =700
                    Name ="Point_Label"
                    Caption ="Point"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =4740
                    Top =30
                    Width =720
                    FontSize =11
                    ForeColor =4210752
                    Name ="btnAddLocations"
                    Caption ="Add Record"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Add location records @ 5, 15, 25, 35, 45 meters (if no records exist)"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b09880ff201010ff201010ff201010ff201010ff201010ff ,
                        0x201010ff201010ff201010ff201010ff201010ff201010ff201010ff00000000 ,
                        0x0000000000000000c0a090fffff8f0fffff8f0fffff0f0fffff0e0fff0e8e0ff ,
                        0xf0e8d0fff0e0d0fff0e0d0fff0e0d0fff0d8d0fff0d8d0ff201810ff00000000 ,
                        0x0000000000000000c0a090ffffffffffd07850ffd07840ffd07040ffc07040ff ,
                        0xc06840ffc06840ffc06840ffc07040ffa06040fff0e0d0ff403830ff00000000 ,
                        0x0000000000000000c0a890ffffffffffd07850fff0b8a0fff0b090fff0a880ff ,
                        0xf0a080fff09870fff09870fff0a880ffc09880fffff0f0ff909090ff00000000 ,
                        0x0000000000000000c0a890ffffffffffd07850ffd07850ffd07840ffd07040ff ,
                        0xc07040ffc07050ffd09070ff70b8c0ff90d8f0ff90f0ffff40c0e0ffa0f0ffff ,
                        0xa0e8ffff90d8f0ffc0a8a0fffffffffffffffffffffffffffffffffffff8f0ff ,
                        0xfff8f0fffff8f0fffff8f0ffb0e8ffff30b8e0ff80e8ffff60c8e0ff90f0ffff ,
                        0x30b8e0ffa0e8ffffc0a8a0ffc0a8a0ffc0a890ffc0a090ffc0a090ffc0a090ff ,
                        0xc09880ffc0a090ffd0c0b0ffa0e8ffff90f0ffffc0f8ffffb0e8f0ffc0f8ffff ,
                        0x90f0ffffa0f0ffff000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000020a8e0ff50c0e0ffb0e8f0fff0ffffffb0e8f0ff ,
                        0x50c0e0ff30b8e0ff000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000080e8ffc090f0ffffc0f8ffffb0e8f0ffc0f8ffff ,
                        0x90f0ffff90d8e0ff000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000050d8ff8030b8e0ff90f0ffff60c0e0ff90f0ffff ,
                        0x30b8e0ff50d0f080000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000030b0e0a040c8f09080e8ffc020b0e0ff70e8ffc0 ,
                        0x50d8f08030b0e080000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =4740
                    LayoutCachedTop =30
                    LayoutCachedWidth =5460
                    LayoutCachedHeight =390
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Shape =1
                    Gradient =12
                    BackColor =14136213
                    BackThemeColorIndex =4
                    BackTint =60.0
                    BorderColor =14136213
                    BorderThemeColorIndex =4
                    BorderTint =60.0
                    ThemeFontIndex =1
                    HoverColor =65280
                    HoverTint =40.0
                    PressedColor =9592887
                    PressedThemeColorIndex =4
                    PressedShade =75.0
                    HoverForeColor =4210752
                    HoverForeThemeColorIndex =0
                    HoverForeTint =75.0
                    PressedForeColor =4210752
                    PressedForeThemeColorIndex =0
                    PressedForeTint =75.0
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
        Begin Section
            Height =375
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =120
                    Top =60
                    Width =420
                    Height =255
                    ColumnWidth =2310
                    Name ="SD_ID"
                    ControlSource ="SD_ID"
                    StatusBarText ="Unique record identifier - primary key"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =660
                    Top =60
                    Width =420
                    Height =255
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Transect_ID"
                    ControlSource ="Transect_ID"
                    StatusBarText ="Foreign key to tbl_LP_Belt_Transect"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1320
                    Top =60
                    Width =600
                    Height =255
                    ColumnWidth =600
                    TabIndex =3
                    BackColor =6750207
                    Name ="Total1"
                    ControlSource ="Total1"
                    StatusBarText ="Total count"
                    FontName ="Tahoma"
                    ConditionalFormat = Begin
                        0x01000000ba000000030000000000000006000000000000000200000001000000 ,
                        0x00000000ffffff00000000000500000003000000050000000100000000000000 ,
                        0xffff66000100000000000000060000002c0000000100000000000000ffff6600 ,
                        0x3000000000003000000000005b0054006f00740061006c0031005d002b005b00 ,
                        0x54006f00740061006c0032005d002b005b0054006f00740061006c0033005d00 ,
                        0x2b005b0054006f00740061006c0034005d003d00300000000000
                    End

                    LayoutCachedLeft =1320
                    LayoutCachedTop =60
                    LayoutCachedWidth =1920
                    LayoutCachedHeight =315
                    ConditionalFormat14 = Begin
                        0x01000400000000000000060000000100000000000000ffffff00010000003000 ,
                        0x0000000000000000000000000000000000000000000000000005000000010000 ,
                        0x0000000000ffff66000100000030000000000000000000000000000000000000 ,
                        0x0000000001000000000000000100000000000000ffff6600250000005b005400 ,
                        0x6f00740061006c0031005d002b005b0054006f00740061006c0032005d002b00 ,
                        0x5b0054006f00740061006c0033005d002b005b0054006f00740061006c003400 ,
                        0x5d003d0030000000000000000000000000000000000000000000000100000000 ,
                        0x0000000100000000000000ffff660010000000490073004e0075006c006c0028 ,
                        0x005b0054006f00740061006c0031005d00290000000000000000000000000000 ,
                        0x0000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2160
                    Top =60
                    Width =600
                    Height =255
                    ColumnWidth =600
                    TabIndex =4
                    BackColor =6750207
                    Name ="Total2"
                    ControlSource ="Total2"
                    StatusBarText ="Total count"
                    FontName ="Tahoma"
                    ConditionalFormat = Begin
                        0x01000000ba000000030000000000000006000000000000000200000001000000 ,
                        0x00000000ffffff00000000000500000003000000050000000100000000000000 ,
                        0xffff66000100000000000000060000002c0000000100000000000000ffff6600 ,
                        0x3000000000003000000000005b0054006f00740061006c0031005d002b005b00 ,
                        0x54006f00740061006c0032005d002b005b0054006f00740061006c0033005d00 ,
                        0x2b005b0054006f00740061006c0034005d003d00300000000000
                    End

                    LayoutCachedLeft =2160
                    LayoutCachedTop =60
                    LayoutCachedWidth =2760
                    LayoutCachedHeight =315
                    ConditionalFormat14 = Begin
                        0x01000300000000000000060000000100000000000000ffffff00010000003000 ,
                        0x0000000000000000000000000000000000000000000000000005000000010000 ,
                        0x0000000000ffff66000100000030000000000000000000000000000000000000 ,
                        0x0000000001000000000000000100000000000000ffff6600250000005b005400 ,
                        0x6f00740061006c0031005d002b005b0054006f00740061006c0032005d002b00 ,
                        0x5b0054006f00740061006c0033005d002b005b0054006f00740061006c003400 ,
                        0x5d003d003000000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3000
                    Top =60
                    Width =600
                    Height =255
                    ColumnWidth =600
                    TabIndex =5
                    BackColor =6750207
                    Name ="Total3"
                    ControlSource ="Total3"
                    StatusBarText ="Total count"
                    FontName ="Tahoma"
                    ConditionalFormat = Begin
                        0x01000000ba000000030000000000000006000000000000000200000001000000 ,
                        0x00000000ffffff00000000000500000003000000050000000100000000000000 ,
                        0xffff66000100000000000000060000002c0000000100000000000000ffff6600 ,
                        0x3000000000003000000000005b0054006f00740061006c0031005d002b005b00 ,
                        0x54006f00740061006c0032005d002b005b0054006f00740061006c0033005d00 ,
                        0x2b005b0054006f00740061006c0034005d003d00300000000000
                    End

                    LayoutCachedLeft =3000
                    LayoutCachedTop =60
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =315
                    ConditionalFormat14 = Begin
                        0x01000300000000000000060000000100000000000000ffffff00010000003000 ,
                        0x0000000000000000000000000000000000000000000000000005000000010000 ,
                        0x0000000000ffff66000100000030000000000000000000000000000000000000 ,
                        0x0000000001000000000000000100000000000000ffff6600250000005b005400 ,
                        0x6f00740061006c0031005d002b005b0054006f00740061006c0032005d002b00 ,
                        0x5b0054006f00740061006c0033005d002b005b0054006f00740061006c003400 ,
                        0x5d003d003000000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3840
                    Top =60
                    Width =600
                    Height =255
                    ColumnWidth =600
                    TabIndex =6
                    BackColor =6750207
                    Name ="Total4"
                    ControlSource ="Total4"
                    StatusBarText ="Total count"
                    FontName ="Tahoma"
                    ConditionalFormat = Begin
                        0x01000000ba000000030000000000000006000000000000000200000001000000 ,
                        0x00000000ffffff00000000000500000003000000050000000100000000000000 ,
                        0xffff66000100000000000000060000002c0000000100000000000000ffff6600 ,
                        0x3000000000003000000000005b0054006f00740061006c0031005d002b005b00 ,
                        0x54006f00740061006c0032005d002b005b0054006f00740061006c0033005d00 ,
                        0x2b005b0054006f00740061006c0034005d003d00300000000000
                    End

                    LayoutCachedLeft =3840
                    LayoutCachedTop =60
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =315
                    ConditionalFormat14 = Begin
                        0x01000300000000000000060000000100000000000000ffffff00010000003000 ,
                        0x0000000000000000000000000000000000000000000000000005000000010000 ,
                        0x0000000000ffff66000100000030000000000000000000000000000000000000 ,
                        0x0000000001000000000000000100000000000000ffff6600250000005b005400 ,
                        0x6f00740061006c0031005d002b005b0054006f00740061006c0032005d002b00 ,
                        0x5b0054006f00740061006c0033005d002b005b0054006f00740061006c003400 ,
                        0x5d003d003000000000000000000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    ListWidth =540
                    Left =180
                    Top =60
                    Width =900
                    Height =255
                    TabIndex =2
                    BackColor =6750207
                    ColumnInfo ="\"\";\"\";\"10\";\"8\""
                    ConditionalFormat = Begin
                        0x01000000ec000000020000000100000000000000000000000f00000001000000 ,
                        0x00000000ffffff00010000000000000010000000450000000100000000000000 ,
                        0xffff660000000000000000000000000000000000000000000000000000000000 ,
                        0x4c0065006e0028005b0050006f0069006e0074005d0029003e00300000000000 ,
                        0x4900490066002800490073004e0075006c006c0028005b0054006f0074006100 ,
                        0x6c0031005d002b005b0054006f00740061006c0032005d002b005b0054006f00 ,
                        0x740061006c0033005d002b005b0054006f00740061006c0034005d0029002c00 ,
                        0x31002c003000290000000000
                    End
                    Name ="Point"
                    ControlSource ="Point"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qry_Densiometer_LU.Point FROM qry_Densiometer_LU; "
                    ColumnWidths ="540"
                    FontName ="Tahoma"
                    OnGotFocus ="[Event Procedure]"

                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000ffffff000e0000004c00 ,
                        0x65006e0028005b0050006f0069006e0074005d0029003e003000000000000000 ,
                        0x00000000000000000000000000000001000000000000000100000000000000ff ,
                        0xff6600340000004900490066002800490073004e0075006c006c0028005b0054 ,
                        0x006f00740061006c0031005d002b005b0054006f00740061006c0032005d002b ,
                        0x005b0054006f00740061006c0033005d002b005b0054006f00740061006c0034 ,
                        0x005d0029002c0031002c00300029000000000000000000000000000000000000 ,
                        0x00000000
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
Option Explicit

' =================================
' Form:         fsub_LP_Densiometer
' Level:        Application form
' Version:      1.01
' Basis:        -
'
' Description:  LP densiometer form object related properties, events, functions & procedures for UI display
'
' Data source:  qry_LP_Intercept
' Data access:  view and delete records
' Pages:        none
' Functions:    -
' Source/date:  John R. Boetsch, June 7, 2006
' References:   -
' Revisions:    unknown - unknown  - 1.00 - initial version
'               BLC - 3/29/2018 - 1.01 - added CallingForm properties
'                                        added documentation, error handling
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Private m_CallingForm As String

'---------------------
' Event Declarations
'---------------------
Public Event InvalidCallingForm(Value As String)

'---------------------
' Properties
'---------------------
Public Property Let CallingForm(Value As String)
        m_CallingForm = Value
End Property

Public Property Get CallingForm() As String
    CallingForm = m_CallingForm
End Property

'---------------------
' Methods
'---------------------

' ---------------------------------
' Sub:          Form_Open
' Description:  form opening actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:   Bonnie Campbell March 29, 2018 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/29/2018 - initial version
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    'set hover
    btnAddLocations.HoverColor = lngGreen

    'automatic record generation
    'generate 5 records @ 5, 15, 25, 35, and 45 meter fixed locations IF no records exist
    'determine record count
    
    Dim rs As DAO.Recordset
    Dim locs() As Variant
    Dim loc As Variant
    
    locs = Array(5, 15, 25, 35, 45)
    
    Set rs = Me.RecordsetClone
    
    If Not (rs.BOF And rs.EOF) Then rs.MoveLast
Debug.Print rs.RecordCount

    If rs.RecordCount = 0 Then
        
Debug.Print "open rs = 0"
        Me.btnAddLocations.Enabled = True
'        For Each loc In locs
'            rs.AddNew
'            rs("Transect_ID") = Me.Parent.Form.Controls("Transect_ID")
'            rs("Point") = loc
'            'rs("Total1") = 0
'            'rs("Total2") = 0
'
'            rs.Update
''                Me.Point = loc
'
'        Next
    Else
        btnAddLocations.Enabled = False
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[fsub_LP_Densiometer form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Current
' Description:  form current actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:   Bonnie Campbell March 29, 2018 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/29/2018 - initial version
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler

    'determine record count
    Dim rs As DAO.Recordset
    Set rs = Me.RecordsetClone
    'rs.MoveFirst
    If Not (rs.BOF And rs.EOF) Then rs.MoveLast
Debug.Print rs.RecordCount

    If rs.RecordCount = 0 Then
        
Debug.Print "current open rs = 0"
        Me.btnAddLocations.Enabled = True
    Else
        btnAddLocations.Enabled = False
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case 3021 'no current record
        Resume Next
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[fsub_LP_Densiometer form])"
    End Select
    Resume Exit_Handler
End Sub


' ---------------------------------
' Sub:          Form_BeforeInsert
' Description:  form actions prior to record insert
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:   unknown, unknown
' Adapted:      Bonnie Campbell March 29, 2018
' Revisions:
'   BLC - 3/29/2018 - initial version
' ---------------------------------
Private Sub Form_BeforeInsert(Cancel As Integer)
On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim Points As DAO.Recordset
    Dim strSQL As String
    On Error GoTo Err_Handler

    ' Check for species overlap
    ' Set SQL
    Set db = CurrentDb
    strSQL = "SELECT tbl_LP_Belt_Transect.Transect_ID, Count(SD_ID) AS PointCount FROM tbl_LP_Belt_Transect LEFT JOIN tbl_LP_Densiometer ON tbl_LP_Belt_Transect.Transect_ID = tbl_LP_Densiometer.Transect_ID GROUP BY tbl_LP_Belt_Transect.Transect_ID HAVING tbl_LP_Belt_Transect.Transect_ID = '" & Me.Parent!Transect_ID & "'"
    Set Points = db.OpenRecordset(strSQL)
    If Points!PointCount > 4 Then
      MsgBox "5 points maximum!", , "Spherical Densiometer"
      DoCmd.CancelEvent
      SendKeys "{ESC}"
      GoTo Exit_Handler
    End If
    
    ' Create the GUID primary key value
    If IsNull(Me!SD_ID) Then
        If GetDataType("tbl_LP_Densiometer", "SD_ID") = dbText Then
            Me.SD_ID = fxnGUIDGen
        End If
    End If
'    DoCmd.RunCommand acCmdSaveRecord  ' Save it.

Exit_Handler:
    Points.Close
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeInsert[fsub_LP_Densiometer form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Point_GotFocus
' Description:  actions when point has focus
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:   unknown, unknown
' Adapted:      Bonnie Campbell March 29, 2018
' Revisions:
'   BLC - 3/29/2018 - initial version
' ---------------------------------
Private Sub Point_GotFocus()
On Error GoTo Err_Handler
    
    If IsNull(Me.Parent!Visit_Date) Then    ' If they didn't bother to enter a date, default to event date.
      Me.Parent!Visit_Date = Me.Parent.Parent!Start_Date
      Me.Parent.Refresh   ' Force save of transect record
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Point_GotFocus[fsub_LP_Densiometer form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnAddLocations_Click
' Description:  create locations button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Tony Toews, October 27, 2009
'   https://stackoverflow.com/questions/1628267/autonumber-value-of-last-inserted-row-ms-access-vba
' Source/date:   Bonnie Campbell March 29, 2018 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/29/2018 - initial version
'   HMT - 8/5/2018  - Removed code that inserted records into parent table since it was not needed
'                     and was creating duplicate records.
' ---------------------------------
Private Sub btnAddLocations_Click()
On Error GoTo Err_Handler

    'automatic record generation deferred - stub only
    'generate 5 records @ 5, 15, 25, 35, and 45 meter fixed locations IF no records exist

Debug.Print Me.Parent.Form.Controls("Transect_ID")
    
    Dim rs As DAO.Recordset
    Dim locs() As Variant
    Dim loc As Variant
    
    locs = Array(5, 15, 25, 35, 45)
Debug.Print Me.Form.Name

    Set rs = Me.RecordsetClone
    
    'determine record count
    If Not (rs.BOF And rs.EOF) Then rs.MoveLast
Debug.Print rs.RecordCount

    'add only if records don't exist
    If rs.RecordCount = 0 Then
        
Debug.Print "rs = 0"
        For Each loc In locs
            rs.AddNew
            rs("Transect_ID") = Me.Parent.Form.Controls("Transect_ID")
            rs("Point") = CStr(loc) & "m"
            'rs("Total1") = 0
            'rs("Total2") = 0
            'rs("Total3") = 0
            'rs("Total4") = 0

            rs.Update
        Next
    
        'disable additions
        btnAddLocations.Enabled = False
    
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnAddLocations_Click[fsub_LP_Densiometer form])"
    End Select
    Resume Exit_Handler
End Sub
