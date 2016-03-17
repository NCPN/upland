Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    TabularFamily =124
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8099
    DatasheetFontHeight =9
    ItemSuffix =13
    Left =2868
    Top =4776
    Right =11268
    Bottom =11628
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xe3687ae93287e340
    End
    RecordSource ="qry_SLI_Gaps"
    Caption ="fsub_SLI_Gaps"
    BeforeInsert ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
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
            Height =360
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =600
                    Top =60
                    Width =1080
                    Height =240
                    FontWeight =700
                    Name ="Species_Label"
                    Caption ="Species"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2400
                    Top =60
                    Width =1020
                    Height =240
                    FontWeight =700
                    Name ="Shrub_Start_Label"
                    Caption ="Start (cm)"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =3660
                    Top =60
                    Width =1019
                    Height =240
                    FontWeight =700
                    Name ="Shrub_End_Label"
                    Caption ="End (cm)"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4920
                    Top =60
                    Width =1260
                    Height =300
                    Name ="ButtonLookup"
                    Caption ="Master Lookup"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6420
                    Top =60
                    Height =300
                    TabIndex =1
                    Name ="ButtonUnknown"
                    Caption ="Unknown Species"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
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
                    Width =360
                    Height =255
                    ColumnWidth =2310
                    Name ="SLI_ID"
                    ControlSource ="SLI_ID"
                    StatusBarText ="Unique record identifier - primary key"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =540
                    Top =60
                    Width =480
                    Height =255
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Transect_ID"
                    ControlSource ="Transect_ID"
                    StatusBarText ="Foreign key to tbl_SL_Transect"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2400
                    Top =60
                    Width =1020
                    Height =255
                    ColumnWidth =900
                    TabIndex =3
                    Name ="Shrub_Start"
                    ControlSource ="Shrub_Start"
                    StatusBarText ="Start of shrub cover to nearest centimeter"
                    BeforeUpdate ="[Event Procedure]"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3660
                    Top =60
                    Width =1019
                    Height =255
                    ColumnWidth =600
                    TabIndex =4
                    Name ="Shrub_End"
                    ControlSource ="Shrub_End"
                    StatusBarText ="End of shrub cover to nearest centimeter"
                    BeforeUpdate ="[Event Procedure]"
                    FontName ="Tahoma"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =6480
                    Left =180
                    Top =60
                    Width =1979
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="Species"
                    ControlSource ="Species"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qryU_Top_Canopy.Master_PLANT_Code, qryU_Top_Canopy.LU_Code, qryU_Top_Cano"
                        "py.Utah_Species FROM qryU_Top_Canopy WHERE (((qryU_Top_Canopy.Utah_Species) Is N"
                        "ot Null)) ORDER BY qryU_Top_Canopy.LU_Code; "
                    ColumnWidths ="0;2160;4320"
                    OnGotFocus ="[Event Procedure]"

                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =5940
                    Top =60
                    Width =705
                    Height =300
                    TabIndex =5
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
' =================================
' MODULE:       fsub_SLI_Gaps
' Level:        Form module
' Version:      1.02
' Description:  data functions & procedures specific to SLI gaps data entry
'
' Source/date:  John R. Boetsch, June 2006
' Adapted:      Bonnie Campbell, 2/3/2016
' Revisions:    RDB - unknown   - 1.00 - initial version
'               BLC - 2/3/2016  - 1.01 - added documentation, adjusted to use transect # overlay
'                                       vs. message box
'               BLC - 2/19/2016 - 1.02 - based on conversation with H. Thomas, this form is
'                                        no longer in use & should no longer be updated,
'                                        however it will remain to handle views of prior data
' =================================

Private Sub Form_BeforeInsert(Cancel As Integer)
    ' Create the GUID primary key value
    If IsNull(Me!SLI_ID) Then
        If GetDataType("tbl_SLI_Gaps", "SLI_ID") = dbText Then
            Me.SLI_ID = fxnGUIDGen
        End If
    End If
End Sub

Private Sub Shrub_End_BeforeUpdate(Cancel As Integer)
    Dim db As DAO.Database
    Dim Points As DAO.Recordset
    Dim strSQL As String
    On Error GoTo Err_Handler
  If Not IsNull(Me!Shrub_End) And Not IsNull(Me!Shrub_Start) Then
    If Me!Shrub_Start <= Me!Shrub_End Then
      MsgBox "Start must be greater than end.", , "Shrub Gaps"
      DoCmd.CancelEvent
      SendKeys "{ESC}"
      Exit Sub
    End If
  End If
  
' Check for species overlap
    If Not IsNull(Me!Shrub_End) Then
      ' Set SQL
      Set db = CurrentDb
      strSQL = "SELECT * FROM [tbl_SLI_Gaps] WHERE [SLI_ID] <> '" & Me!SLI_ID & "' AND [Transect_ID] = '" & Me![Transect_ID] & "' AND [Species] = '" & Me!Species & "' AND " & Me!Shrub_End & " BETWEEN [Shrub_Start] AND [Shrub_End]"
      Set Points = db.OpenRecordset(strSQL)
      If Points.EOF Then
        GoTo Exit_Procedure
      Else
        MsgBox "Species overlap.", , "Shrub Gaps"
        DoCmd.CancelEvent
        SendKeys "{ESC}"
      End If
    End If
Exit_Procedure:
    Points.Close
    Exit Sub
  
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
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
            "Error encountered (#" & Err.Number & " - Species_GotFocus[Form_fsub_SLI_Gaps])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub Shrub_Start_BeforeUpdate(Cancel As Integer)
    Dim db As DAO.Database
    Dim Points As DAO.Recordset
    Dim strSQL As String
    On Error GoTo Err_Handler
  If Not IsNull(Me!Shrub_End) And Not IsNull(Me!Shrub_Start) Then
    If Me!Shrub_Start <= Me!Shrub_End Then
      MsgBox "Start must be greater than end.", , "Shrub Gaps"
      DoCmd.CancelEvent
      SendKeys "{ESC}"
      Exit Sub
    End If
  End If
  
' Check for species overlap
    If Not IsNull(Me!Shrub_Start) Then
      ' Set SQL
      Set db = CurrentDb
      strSQL = "SELECT * FROM [tbl_SLI_Gaps] WHERE [Transect_ID] = '" & Me![Transect_ID] & "' AND [Species] = '" & Me!Species & "' AND " & Me!Shrub_Start & " BETWEEN [Shrub_Start] AND [Shrub_End]"
      Set Points = db.OpenRecordset(strSQL)
      If Points.EOF Then
        GoTo Exit_Procedure
      Else
        MsgBox "Species overlap.", , "Shrub Gaps"
        DoCmd.CancelEvent
        SendKeys "{ESC}"
      End If
    End If
Exit_Procedure:
  Points.Close
  Exit Sub
  
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

Private Sub ButtonLookup_Click()
On Error GoTo Err_ButtonLookup_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Master_Species"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_ButtonLookup_Click:
    Exit Sub

Err_ButtonLookup_Click:
    MsgBox Err.Description
    Resume Exit_ButtonLookup_Click
    
End Sub

Private Sub ButtonUnknown_Click()
On Error GoTo Err_ButtonUnknown_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_List_Unknown"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , acDialog
    Me.Requery
    Me.Refresh

Exit_ButtonUnknown_Click:
    Exit Sub

Err_ButtonUnknown_Click:
    MsgBox Err.Description
    Resume Exit_ButtonUnknown_Click
    
End Sub

Private Sub ButtonDelete_Click()
On Error GoTo Err_ButtonDelete_Click


    DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
    DoCmd.DoMenuItem acFormBar, acEditMenu, 6, , acMenuVer70

Exit_ButtonDelete_Click:
    Exit Sub

Err_ButtonDelete_Click:
    MsgBox Err.Description
    Resume Exit_ButtonDelete_Click
    
End Sub
