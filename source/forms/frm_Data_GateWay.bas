Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    FilterOn = NotDefault
    OrderByOn = NotDefault
    ScrollBars =2
    TabularFamily =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10560
    DatasheetFontHeight =10
    ItemSuffix =37
    Left =2490
    Top =1875
    Right =13050
    Bottom =7725
    DatasheetGridlinesColor =12632256
    Filter ="Unit_code = 'ARCH' AND Site_Selection = -1"
    OrderBy ="Plot_ID"
    RecSrcDt = Begin
        0x29b5dcdf75fbe240
    End
    RecordSource ="qfrm_Data_Gateway"
    Caption ="Data Gateway"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =255
    AllowLayoutView =0
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
            Height =1248
            BackColor =11056034
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =2460
                    Top =1020
                    Width =1680
                    Height =228
                    Name ="lblUpdated_Date"
                    Caption ="Entered/updated*"
                    FontName ="Arial"
                    OnDblClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =105
                    Top =1020
                    Width =795
                    Height =225
                    Name ="lblUnit_Code"
                    Caption ="Unit*"
                    FontName ="Arial"
                    OnDblClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9720
                    Top =120
                    Width =720
                    Height =354
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    Name ="btnClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Close the data entry form"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =120
                    Width =4860
                    Height =408
                    BackColor =16777215
                    ForeColor =0
                    Name ="labOverview"
                    Caption ="* Double-click on the field label to change sort order.  Double-click on a Plot "
                        "ID to open the Site form for that record."
                    FontName ="Arial"
                    ControlTipText ="View mode"
                End
                Begin ComboBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =660
                    Width =960
                    ColumnOrder =1
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"10\";\"8\""
                    Name ="cbxPark"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT tbl_Locations.Unit_Code FROM tbl_Locations ORDER BY tbl_Location"
                        "s.Unit_Code; "
                    StatusBarText ="Park code"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =120
                            Top =660
                            Width =1320
                            Height =228
                            FontWeight =700
                            Name ="lblPark"
                            Caption ="Filter by:  Park"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ToggleButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =5580
                    Top =660
                    Width =480
                    Height =300
                    ColumnOrder =0
                    Name ="tglFilterByPark"
                    Caption ="Filter on"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadad0000adadaddadada0660dadadaadadad0660adadaddadada0f80dadada ,
                        0xadadad0f80adadaddadad088860adadaadad06888660adaddad068f888660ada ,
                        0xad068f88888660add068fff88886660aa00000000000000ddadadadadadadada ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    FontName ="Arial"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Turn the park filter on or off"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =1260
                    Top =1020
                    Width =660
                    Height =224
                    FontWeight =700
                    Name ="lblPlot_ID"
                    Caption ="Plot ID*"
                    FontName ="Arial"
                    OnDblClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7920
                    Top =120
                    Width =1560
                    FontSize =9
                    FontWeight =700
                    TabIndex =4
                    Name ="btnNewSite"
                    Caption ="Add a new site"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    Left =4140
                    Top =1020
                    Width =1800
                    Height =225
                    Name ="lblSelForMonitoring"
                    Caption ="Selected for Monitoring"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6720
                    Top =120
                    Width =960
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    Name ="btnRefresh"
                    Caption ="Refresh"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =93
                    IMESentenceMode =3
                    ListWidth =390
                    Left =4680
                    Top =660
                    Width =720
                    ColumnOrder =2
                    TabIndex =5
                    Name ="cbxMonitor"
                    RowSourceType ="Value List"
                    RowSource ="\"On\";\"Off\""
                    ColumnWidths ="390"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"On\""

                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =2640
                    Top =660
                    Width =2040
                    Height =239
                    FontWeight =700
                    Name ="lblMonitor"
                    Caption ="Selected for Monitoring"
                    FontName ="Arial"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            Height =420
            BackColor =11056034
            Name ="Detail"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =2400
                    Top =60
                    Width =1920
                    ColumnWidth =1710
                    TabIndex =1
                    Name ="tbxUpdated"
                    ControlSource ="Updated_Date"
                    Format ="yyyy mmm dd hh:nn"
                    StatusBarText ="Date on which data entry occurred"
                    FontName ="Arial"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =120
                    Top =60
                    Width =780
                    ColumnWidth =2310
                    Name ="tbxUnit_Code"
                    ControlSource ="Unit_Code"
                    StatusBarText ="Unit code"
                    FontName ="Arial"

                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =2160
                    Top =60
                    Width =420
                    TabIndex =2
                    Name ="tbxLocation_ID"
                    ControlSource ="Location_ID"
                    StatusBarText ="Name of the location"
                    FontName ="Arial"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1320
                    Top =60
                    Width =600
                    TabIndex =3
                    Name ="tbxPlot_ID"
                    ControlSource ="Plot_ID"
                    StatusBarText ="Plot identifier"

                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5700
                    Top =60
                    Width =1320
                    Height =300
                    TabIndex =4
                    Name ="btnVisitList"
                    Caption ="View Visits"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =4920
                    Top =120
                    TabIndex =5
                    Name ="Site_Selection"
                    ControlSource ="Site_Selection"
                    StatusBarText ="Site accepted or rejected"

                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7200
                    Top =60
                    Width =1319
                    Height =299
                    TabIndex =6
                    Name ="btnNewVisit"
                    Caption ="Add New Visit"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8700
                    Top =60
                    Width =1650
                    Height =300
                    TabIndex =7
                    Name ="btnSiteChar"
                    Caption ="Site Characterization"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
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
' Form:         frm_Data_Gateway
' Level:        Application form
' Version:      1.03
' Basis:        Dropdown form
'
' Description:  Plot field check form object related properties, events, functions & procedures for UI display
'
' Data source:  qfrm_Data_Gateway
' Data access:  view and delete records (delete by cmdDeleteRec)
' Pages:        none
' Functions:    SortRecords
' Source/date:  John R. Boetsch, June 7, 2006
' References:   -
' Revisions:    JRB - 6/7/2006 - 1.00 - initial version
'               Simon Kingston, 9/x/2006 - 1.01 added CorrectText calls where strings were being used in criteria
'                                               updated cmdDeleteRec_Click() event to use appropriate criteria depending on primary key
'               BLC - 3/24/2017 - 1.02 - added CallingForm, CallingRecordID properties
'               BLC - 8/9/2017  - 1.03 - added documentation, error handling
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Dim strSortField As String    ' Keeps track of current sort settings
Dim strSortOrder As String
Dim strSortFieldLabel As String

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
' Source/date:   John R. Boetsch, June 7, 2006 - for NCPN tools
' Adapted:      -
' Revisions:
'   JRB - 6/7/2006 - initial version
'   BLC - 8/9/2017 - added documentation, error handling
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Err_Handler

    Dim varReturn As Variant

    ' On opening the form, set the initial sort order
    strSortFieldLabel = "lblPlot_ID"
    varReturn = SortRecords("Unit_Code", "Plot_ID")
    ' Set the filter
    If fxnSwitchboardIsOpen Then
        Me.cbxPark = Forms!frm_Switchboard.cPark
        Me.Filter = "Unit_code = " & CorrectText(Me.cbxPark) & " AND Site_Selection = " & -1
        Me.FilterOn = True
        Me.lblPark.FontBold = True
        Me!lblMonitor.FontBold = True
        Me.tglFilterByPark = True
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[frm_Data_Gateway form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
'   Form Filtering
' ---------------------------------

' ---------------------------------
' Sub:          cbxPark_AfterUpdate
' Description:  park selection dropdown after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:   John R. Boetsch, June 7, 2006 - for NCPN tools
' Adapted:      -
' Revisions:
'   JRB - 6/7/2006 - initial version
'   BLC - 8/9/2017 - added documentation, error handling
' ---------------------------------
Private Sub cbxPark_AfterUpdate()
    On Error GoTo Err_Handler

    Me.Filter = "Unit_code = " & CorrectText(Me.cbxPark)
    If tglFilterByPark Then
      Me.Filter = "Unit_code = " & CorrectText(Me.cbxPark)
      Me.FilterOn = True
      Me.lblPark.FontBold = True
      If Me!cbxMonitor = "On" Then
        If Not IsNull(Me!cbxPark) Then
          Me.Filter = Me.Filter & " AND Site_Selection = " & -1
        Else
          Me.Filter = "Site_Selection = " & -1
        End If
        Me.FilterOn = True
        Me.lblMonitor.FontBold = True
      Else
        Me!lblMonitor.FontBold = False
      End If
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxPark_AfterUpdate[frm_Data_Gateway form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxMonitor_AfterUpdate
' Description:  selected for monitoring dropdown after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:   John R. Boetsch, June 7, 2006 - for NCPN tools
' Adapted:      -
' Revisions:
'   JRB - 6/7/2006 - initial version
'   BLC - 8/9/2017 - added documentation, error handling
' ---------------------------------
Private Sub cbxMonitor_AfterUpdate()
    On Error GoTo Err_Handler

      Me.FilterOn = True
      Me.lblPark.FontBold = True
      Me.Filter = "Unit_code = " & CorrectText(Me.cbxPark)
      If Me!cbxMonitor = "On" Then
        Me.Filter = Me.Filter & " AND Site_Selection = " & -1
        Me.lblMonitor.FontBold = True
      Else
        Me.lblMonitor.FontBold = False
      End If
 
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxMonitor_AfterUpdate[frm_Data_Gateway form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tglFilterByPark_AfterUpdate
' Description:  filter by park dropdown after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:   John R. Boetsch, June 7, 2006 - for NCPN tools
' Adapted:      -
' Revisions:
'   JRB - 6/7/2006 - initial version
'   BLC - 8/9/2017 - added documentation, error handling
' ---------------------------------
Private Sub tglFilterByPark_AfterUpdate()
    On Error GoTo Err_Handler

    If Me.ActiveControl Then
      If Not IsNull(Me!cbxPark) Then
        Me.Filter = "Unit_code = " & CorrectText(Me.cbxPark)
        Me.FilterOn = True
        Me.lblPark.FontBold = True
      End If
      If Me!selMon = "On" Then
        If Not IsNull(Me!cbxPark) Then
          Me.Filter = Me.Filter & " AND Site_Selection = " & -1
        Else
          Me.Filter = "Site_Selection = " & vbYes
        End If
        Me.FilterOn = True
        Me.lblMonitor.FontBold = True
      Else
        Me!lblMonitor.FontBold = False
      End If
    Else
        Me.FilterOn = False
        Me.lblPark.FontBold = False
        Me!lblMonitor.FontBold = False
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglFilterByPark_AfterUpdate[frm_Data_Gateway form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
'  Sorting Methods
' ---------------------------------

' ---------------------------------
' FUNCTION:     SortRecords
' Description:  Sorts the records by the indicated field
' Parameters:   strFieldName
' Returns:      none
' Throws:       none
' References:   strFieldName, strSortOrder, strSortFieldLabel
'               (form-level variables)
' Source/date:  John R. Boetsch, May 5, 2006
' Revisions:    <name, date, desc - add lines as you go>
'               JRB, 5/5/2006 - initial version
'               BLC, 8/9/2017 - added error handling, documentation
'                               renamed from fxnSortRecords
' ---------------------------------
Private Function SortRecords(ByVal strFieldName As String, _
    Optional ByVal strField2Name As String)
    On Error GoTo Err_Handler

    Dim strOrderBy As String

    ' If already sorting in ascending order by this field, sort descending
    If strFieldName = strSortField And strSortOrder = "" Then
        strSortOrder = " DESC"
    Else: strSortOrder = ""
    End If
    ' Create the order by string and activate the filter
    strOrderBy = strFieldName & strSortOrder
    If strField2Name <> "" Then
        strOrderBy = strField2Name & " DESC, " & strOrderBy
    End If
    strSortField = strFieldName
    Me.Form.OrderBy = strOrderBy
    Me.Form.OrderByOn = True

    ' Change the label format to indicate the sorted field
    Me.Controls.Item(strSortFieldLabel).FontItalic = False
    Me.Controls.Item(strSortFieldLabel).FontBold = False
    strSortFieldLabel = "lbl" & strFieldName
    Me.Controls.Item(strSortFieldLabel).FontItalic = True
    Me.Controls.Item(strSortFieldLabel).FontBold = True

Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SortRecords[frm_Data_Gateway form])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Sub:          lblUnit_Code_DblClick
' Description:  unit code (park) label double click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:   John R. Boetsch, June 7, 2006 - for NCPN tools
' Adapted:      -
' Revisions:
'   JRB - 6/7/2006 - initial version
'   BLC - 8/9/2017 - added documentation, error handling
' ---------------------------------
Private Sub lblUnit_Code_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler

    SortRecords ("Unit_Code")

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblUnit_Code_DblClick[frm_Data_Gateway form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblPlot_ID_DblClick
' Description:  form opening actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:   John R. Boetsch, June 7, 2006 - for NCPN tools
' Adapted:      -
' Revisions:
'   JRB - 6/7/2006 - initial version
'   BLC - 8/9/2017 - added documentation, error handling
' ---------------------------------
Private Sub lblPlot_ID_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler

    SortRecords ("Plot_ID")

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblPlot_ID_DblClick[frm_Data_Gateway form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblUpdated_DblClick
' Description:  entered/updated label double click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:   John R. Boetsch, June 7, 2006 - for NCPN tools
' Adapted:      -
' Revisions:
'   JRB - 6/7/2006 - initial version
'   BLC - 8/9/2017 - added documentation, error handling
' ---------------------------------
Private Sub lblUpdated_Date_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler

    SortRecords ("Updated_Date")

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblUpdated_Date_DblClick[frm_Data_Gateway form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
'   Record Level Methods
' ---------------------------------

' ---------------------------------
' Sub:          btnVisitList_Click
' Description:  visit list button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:   John R. Boetsch, June 7, 2006 - for NCPN tools
' Adapted:      -
' Revisions:
'   JRB - 6/7/2006 - initial version
'   BLC - 8/9/2017 - added documentation, error handling
' ---------------------------------
Private Sub btnVisitList_Click()
On Error GoTo Err_Handler

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Visit_Date"
    
    stLinkCriteria = "[Location_ID]=" & "'" & Me![tbxLocation_ID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnVisitList_Click[frm_Data_Gateway form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnNewVisit_Click
' Description:  new visit click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:   John R. Boetsch, June 7, 2006 - for NCPN tools
' Adapted:      -
' Revisions:
'   JRB - 6/7/2006 - initial version
'   BLC - 8/9/2017 - added documentation, error handling
' ---------------------------------
Private Sub btnNewVisit_Click()
On Error GoTo Err_Handler

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Plot_Revisit"
    
    stLinkCriteria = "[Location_ID]=" & "'" & Me![tbxLocation_ID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
'    DoCmd.Close acForm, "frm_Data_Gateway"

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnNewVisit_Click[frm_Data_Gateway form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnSiteChar_Click
' Description:  site characterization button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:   John R. Boetsch, June 7, 2006 - for NCPN tools
' Adapted:      -
' Revisions:
'   JRB - 6/7/2006 - initial version
'   BLC - 8/9/2017 - added documentation, error handling
' ---------------------------------
Private Sub btnSiteChar_Click()
On Error GoTo Err_Handler

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Locations"
    
    stLinkCriteria = "[Location_ID]=" & "'" & Me![tbxLocation_ID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSiteChar_Click[frm_Data_Gateway form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
'  Overall Methods
' ---------------------------------

' ---------------------------------
' Sub:          btnNewSite_Click
' Description:  new site click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:   John R. Boetsch, June 7, 2006 - for NCPN tools
' Adapted:      -
' Revisions:
'   JRB - 6/7/2006 - initial version
'   BLC - 8/9/2017 - added documentation, error handling
' ---------------------------------
Private Sub btnNewSite_Click()
On Error GoTo Err_Handler

    DoCmd.OpenForm "frm_Locations", , , , acFormAdd, , "New record"

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnNewSite_Click[frm_Data_Gateway form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnRefresh_Click
' Description:  refresh button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:   John R. Boetsch, June 7, 2006 - for NCPN tools
' Adapted:      -
' Revisions:
'   JRB - 6/7/2006 - initial version
'   BLC - 8/9/2017 - added documentation, error handling
' ---------------------------------
Private Sub btnRefresh_Click()
On Error GoTo Err_Handler

    Me.Requery

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnRefresh_Click[frm_Data_Gateway form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnClose_Click
' Description:  close button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:   John R. Boetsch, June 7, 2006 - for NCPN tools
' Adapted:      -
' Revisions:
'   JRB - 6/7/2006 - initial version
'   BLC - 8/9/2017 - added documentation, error handling
' ---------------------------------
Private Sub btnClose_Click()
    On Error GoTo Err_Handler

    DoCmd.Close , , acSaveNo

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnClose_Click[frm_Data_Gateway form])"
    End Select
    Resume Exit_Handler
End Sub
