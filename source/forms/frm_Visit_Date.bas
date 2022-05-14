Version =21
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    FilterOn = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    TabularFamily =127
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4680
    DatasheetFontHeight =9
    ItemSuffix =15
    Left =-32386
    Top =4185
    Right =-27706
    Bottom =7470
    DatasheetGridlinesColor =12632256
    Filter ="[Location_ID]='20081016093629-53504526.6151428'"
    OrderBy ="[Start_Date]"
    RecSrcDt = Begin
        0x5bd611c7ad13e340
    End
    RecordSource ="qfrm_Visit_Date"
    Caption ="Select a Visit"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
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
            Height =1140
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =180
                    Top =900
                    Width =1035
                    Height =240
                    FontWeight =700
                    Name ="Start_Date_Label"
                    Caption ="Visit Date"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =120
                    Top =120
                    Width =540
                    Height =240
                    Name ="Unit_Code_Label"
                    Caption ="Park"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =120
                    Top =420
                    Width =600
                    Height =240
                    Name ="Plot_ID_Label"
                    Caption ="Plot ID"
                    Tag ="DetachedLabel"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3360
                    Top =780
                    Width =1020
                    Height =300
                    Name ="btnClose"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"

                    CursorOnHover =1
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =720
                    Top =120
                    Width =600
                    Height =255
                    ColumnWidth =540
                    ColumnOrder =0
                    TabIndex =1
                    Name ="Unit_Code"
                    ControlSource ="Unit_Code"
                    StatusBarText ="Park Code."

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =780
                    Top =420
                    Width =600
                    Height =255
                    ColumnWidth =600
                    ColumnOrder =1
                    TabIndex =2
                    Name ="Plot_ID"
                    ControlSource ="Plot_ID"
                    StatusBarText ="Plot identifier"

                End
                Begin Label
                    OverlapFlags =85
                    Left =1680
                    Top =240
                    Width =2640
                    Height =360
                    FontSize =12
                    FontWeight =700
                    Name ="lblTitle"
                    Caption ="Select a Visit to Edit"
                End
            End
        End
        Begin Section
            Height =360
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =660
                    Height =255
                    ColumnWidth =2310
                    Name ="Event_ID"
                    ControlSource ="Event_ID"
                    StatusBarText ="M. Event identifier (Event_ID)"

                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =840
                    Top =60
                    Width =690
                    Height =255
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Location_ID"
                    ControlSource ="Location_ID"
                    StatusBarText ="M. Link to tbl_Locations (Loc_ID)"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    BackStyle =0
                    IMESentenceMode =3
                    Left =180
                    Top =60
                    Width =1035
                    Height =255
                    ColumnWidth =1035
                    TabIndex =2
                    Name ="Start_Date"
                    ControlSource ="Start_Date"
                    Format ="Short Date"
                    StatusBarText ="M. Starting date for the event (Start_Date)"

                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1860
                    Top =60
                    Width =1020
                    Height =300
                    TabIndex =3
                    Name ="btnEdit"
                    Caption ="Edit Visit"
                    OnClick ="[Event Procedure]"

                    CursorOnHover =1
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
' Form:         frm_Visit_Date
' Level:        Application form
' Version:      1.03
' Basis:        -
'
' Description:  Visit viewing form object related properties, events, functions & procedures for UI display
'
' Data source:  qfrm_Data_Gateway
' Data access:  view and delete records (delete by cmdDeleteRec)
' Pages:        none
' Functions:    SortRecords
' Source/date:  John R. Boetsch, June 7, 2006
' References:   -
' Revisions:    JRB - 6/7/2006  - 1.00 - initial version
'               BLC - 8/10/2017 - 1.01 - added CallingForm, CallingRecordID properties
'                                        added documentation, error handling
'               BLC - 2/1/2018  - 1.02 - filter by date (yr, mo, day)
'               BLC - 3/29/2018 - 1.03 - revise order by ([Start_Date] vs yr, mo, day)
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
' Source/date:   John R. Boetsch, June 7, 2006 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 8/10/2017 - initial version
'   BLC - 2/1/2018  - filter by date (yr, mo, day)
'   BLC - 3/29/2018 - revise order by
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Err_Handler

    Me.OrderBy = "[Start_Date]" '"Year([Start_Date]) AND Month([Start_Date]) AND Day([Start_Date])"
    Me.OrderByOn = True
    Me.OrderByOnLoad = True

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[frm_Visit_Date form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnEdit_Click
' Description:  edit button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:   John R. Boetsch, June 7, 2006 - for NCPN tools
' Adapted:      -
' Revisions:
'   JRB - 6/7/2006 - initial version
'   BLC - 8/10/2017 - added documentation, error handling
'   BLC - 8/11/2017 - set TempVars for reopening Data Entry w/ PlotCheck
' ---------------------------------
Private Sub btnEdit_Click()
On Error GoTo Err_Handler

    Dim strCriteriaLoc As String
    Dim strCriteriaEvent As String

        strCriteriaLoc = GetCriteriaString("[Location_ID]=", "tbl_Locations", "Location_ID", Me.Name, "Location_ID")
        strCriteriaEvent = GetCriteriaString("[Event_ID]=", "tbl_Events", "Event_ID", Me.Name, "Event_ID")
        
        'set TempVars for re-opening from PlotCheck
        SetTempVar "CriteriaLoc", strCriteriaLoc
        SetTempVar "CriteriaEvent", strCriteriaEvent
        
        ' Filter by location and event
        DoCmd.OpenForm "frm_Data_Entry", , , strCriteriaLoc & " AND " & strCriteriaEvent, , , strCriteriaEvent
        DoCmd.Close acForm, "frm_Visit_Date"
        DoCmd.SelectObject acForm, "frm_Data_Entry"

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnEdit_Click[frm_Visit_Date form])"
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
'   BLC - 8/10/2017 - added documentation, error handling
' ---------------------------------
Private Sub btnClose_Click()
On Error GoTo Err_Handler

    DoCmd.Close

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnClose_Click[frm_Visit_Date form])"
    End Select
    Resume Exit_Handler
End Sub
