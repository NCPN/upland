Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    FilterOn = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =0
    DatasheetFontHeight =11
    Left =585
    Top =1320
    Right =9240
    Bottom =3660
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x96574bf84ef9e440
    End
    RecordSource ="usys_temp_display"
    Caption ="Plot Check Results"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    AllowFormView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    OrderByOnLoad =0
    OrderByOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Section
            Height =0
            Name ="Detail"
            AutoHeight =1
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
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
' Form:         QueryView
' Level:        Application form
' Version:      1.00
' Basis:        Dropdown form
'
' Description:  Plot field check form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, August 7, 2017
' References:   -
' Revisions:    BLC - 8/7/2017 - 1.00 - initial version
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Private m_Title As String
Private m_CallingForm As String
Private m_CallingRecordID As Integer
Private m_CallingSampleDate As Date

'---------------------
' Event Declarations
'---------------------
Public Event InvalidTitle(value As String)
Public Event InvalidCallingForm(value As String)
Public Event InvalidCallingRecordID(value As Integer)
Public Event InvalidCallingSampleDate(value As Date)

'---------------------
' Properties
'---------------------
Public Property Let Title(value As String)
    If Len(value) > 0 Then
        m_Title = value

        'set the form title & caption
'        Me.lblTitle.Caption = m_Title
        Me.Caption = m_Title
    Else
        RaiseEvent InvalidTitle(value)
    End If
End Property

Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Let CallingForm(value As String)
        m_CallingForm = value
End Property

Public Property Get CallingForm() As String
    CallingForm = m_CallingForm
End Property

Public Property Let CallingRecordID(value As Integer)
        m_CallingRecordID = value
End Property

Public Property Get CallingRecordID() As Integer
    CallingRecordID = m_CallingRecordID
End Property

Public Property Let CallingSampleDate(value As Date)
        m_CallingSampleDate = value
End Property

Public Property Get CallingSampleDate() As Date
    CallingSampleDate = m_CallingSampleDate
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
' References:
'   Steve Schapel, September 15, 2008
'   https://www.pcreview.co.uk/threads/switch-focus-to-query-through-vba.3622059/
' Source/date:  Bonnie Campbell, August 7, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 8/7/2017 - initial version
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    'default
    Me.CallingForm = "PlotCheck"
    Me.CallingRecordID = -1
    Me.CallingSampleDate = Date
        
    'If Len(Nz(Me.OpenArgs, "")) > 0 Then Me.CallingForm = Me.OpenArgs

    'minimize calling form
    ToggleForm Me.CallingForm, -1

    'set record
    If Len(Nz(Me.OpenArgs, "")) > 0 Then
        If InStr(Me.OpenArgs, "|") Then
            Dim ary() As String
            ary = Split(Me.OpenArgs, "|")
            Me.CallingForm = ary(0)
            Me.CallingRecordID = ary(1)
            Me.CallingSampleDate = ary(2)
        End If
    End If

    'set the record source to the temp display query (populated from PlotCheck)
    Me.RecordSource = "usys_temp_display"
    
    Me.Caption = "Plot Check Results"
                
    'set underlying data
    'Set Me.Recordset = GetRecords("s_template_num_records")
    
    'set form height <- must be set or detail height = 1 record
    '                   due to setting recordset programmatically
    'normally this would be:
    'Me.InsideHeight = Me.FormHeader.Height + Me.FormFooter.Height + _
    '                    (Me.Detail.Height * 10)
    'but QueryView form doesn't have a form header or footer - only a detail section
    Me.InsideHeight = Me.Detail.Height * 10

    'defaults
    Me.Filter = "[FieldCheck]=" & 1
    'Me.FilterOnLoad = True
    'Me.AllowEdits = True
    Me.AllowFilters = True
    
    'clear num records & run queries
'    RunQueryView
        
    Me.Requery

    'determine if the form should remain open - 0 records just close it
    Dim rs As DAO.Recordset
    
    Set rs = Me.RecordsetClone
    If Not (rs.BOF And rs.EOF) Then rs.MoveLast
        
    'close when there are no records
    If Not (rs.RecordCount > 0) Then DoCmd.Close acForm, Me.Name, acSaveNo
    
Exit_Handler:
    Set rs = Nothing
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case 3048 'Cannot open any more databases
        Resume Next
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[QueryView form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Load
' Description:  form loading actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, August 7, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 8//2017 - initial version
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler

    'eliminate NULLs
    If IsNull(Me.OpenArgs) Then GoTo Exit_Handler

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[QueryView form])"
    End Select
    Resume Exit_Handler
End Sub
