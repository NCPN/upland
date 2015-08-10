Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11880
    DatasheetFontHeight =9
    ItemSuffix =28
    Left =1440
    Top =2325
    Right =13605
    Bottom =5880
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x9aa5143d6c56e340
    End
    RecordSource ="tbl_LP_Shrub"
    Caption ="fsub_LP_Belt_Shrub"
    BeforeInsert ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
        End
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
        End
        Begin Tab
            BackStyle =0
        End
        Begin FormHeader
            Height =960
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =240
                    Top =720
                    Width =1320
                    Height =240
                    FontWeight =700
                    Name ="Species_Label"
                    Caption ="Shrub Species"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    Left =2520
                    Top =720
                    Width =720
                    Height =240
                    FontWeight =700
                    Name ="Alive_Label"
                    Caption ="Alive?"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =223
                    TextAlign =2
                    Left =3060
                    Top =720
                    Width =1080
                    Height =240
                    FontWeight =700
                    Name ="HC10_Label"
                    Caption ="0-10cm"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =4140
                    Top =720
                    Width =1080
                    Height =240
                    FontWeight =700
                    Name ="HC25_Label"
                    Caption ="10.1-25cm"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =5220
                    Top =720
                    Width =1080
                    Height =240
                    FontWeight =700
                    Name ="HC50_Label"
                    Caption ="25.1-50cm"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =6300
                    Top =720
                    Width =1080
                    Height =240
                    FontWeight =700
                    Name ="HC100_Label"
                    Caption ="50.1-100cm"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =7380
                    Top =720
                    Width =1080
                    Height =240
                    FontWeight =700
                    Name ="HC2m_Label"
                    Caption ="1.01-2m"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =8460
                    Top =720
                    Width =1080
                    Height =240
                    FontWeight =700
                    Name ="HCGT2_Label"
                    Caption =">2.01m"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    Left =3060
                    Top =480
                    Width =6480
                    Height =240
                    FontWeight =700
                    Name ="Label22"
                    Caption ="Height Class Totals"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2460
                    Top =60
                    Width =6480
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label23"
                    Caption ="Number of Live Shrubs Rooted in 1 Meter Belt Transect"
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
                    Name ="Shrub_ID"
                    ControlSource ="Shrub_ID"
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
                    TabIndex =1
                    Name ="Transect_ID"
                    ControlSource ="Transect_ID"
                    StatusBarText ="Foreign key to tbl_Canopy_Transect"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3300
                    Top =60
                    Width =540
                    Height =255
                    ColumnWidth =600
                    TabIndex =4
                    Name ="HC10"
                    ControlSource ="HC10"
                    StatusBarText ="0-10cm height class total"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4380
                    Top =60
                    Width =540
                    Height =255
                    ColumnWidth =600
                    TabIndex =5
                    Name ="HC25"
                    ControlSource ="HC25"
                    StatusBarText ="10.1-25cm height class total"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5520
                    Top =60
                    Width =540
                    Height =255
                    ColumnWidth =600
                    TabIndex =6
                    Name ="HC50"
                    ControlSource ="HC50"
                    StatusBarText ="25.1-50cm height class total"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6540
                    Top =60
                    Width =540
                    Height =255
                    ColumnWidth =600
                    TabIndex =7
                    Name ="HC100"
                    ControlSource ="HC100"
                    StatusBarText ="50.1-100cm height class total"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7680
                    Top =60
                    Width =540
                    Height =255
                    ColumnWidth =600
                    TabIndex =8
                    Name ="HC2m"
                    ControlSource ="HC2m"
                    StatusBarText ="1.01-2m height class total"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8760
                    Top =60
                    Width =540
                    Height =255
                    ColumnWidth =600
                    TabIndex =9
                    Name ="HCGT2"
                    ControlSource ="HCGT2"
                    StatusBarText =">2.01m height class total"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =375
                    Left =2520
                    Top =60
                    Width =780
                    TabIndex =3
                    Name ="Alive"
                    ControlSource ="Alive"
                    RowSourceType ="Value List"
                    RowSource ="-1;\"Yes\";0;\"No\""
                    ColumnWidths ="0;375"
                    DefaultValue ="-1"
                End
                Begin ComboBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =6480
                    Left =120
                    Top =60
                    Width =2304
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="Species"
                    ControlSource ="Species"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qryU_Top_Canopy.Master_PLANT_Code, qryU_Top_Canopy.LU_Code, qryU_Top_Cano"
                        "py.Utah_Species FROM qryU_Top_Canopy WHERE (((qryU_Top_Canopy.Utah_Species) Is N"
                        "ot Null)) ORDER BY qryU_Top_Canopy.LU_Code; "
                    ColumnWidths ="0;2160;4320"
                    BeforeUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10080
                    Top =60
                    Width =1275
                    Height =300
                    TabIndex =10
                    ForeColor =255
                    Name ="ButtonDelete"
                    Caption ="Delete Record"
                    OnClick ="[Event Procedure]"
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
                    Left =4560
                    Top =60
                    Width =606
                    Height =288
                    Name ="ButtonA1"
                    Caption ="+ 1"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5280
                    Top =60
                    Width =606
                    Height =288
                    TabIndex =1
                    Name ="ButtonA5"
                    Caption ="+ 5"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6000
                    Top =60
                    Width =606
                    Height =288
                    TabIndex =2
                    Name ="ButtonS1"
                    Caption ="- 1"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6720
                    Top =60
                    Width =606
                    Height =288
                    TabIndex =3
                    Name ="ButtonS5"
                    Caption ="- 5"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7440
                    Top =60
                    Width =606
                    Height =288
                    TabIndex =4
                    Name ="ButtonZero"
                    Caption ="0"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
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

Private Sub Button_Master_Species_Click()
On Error GoTo Err_Button_Master_Species_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim strOpenArg As String

    strOpenArg = "fsub_LP_Belt_Shrub"
    stDocName = "frm_Master_Species"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , , strOpenArg

Exit_Button_Master_Species_Click:
    Exit Sub

Err_Button_Master_Species_Click:
    MsgBox Err.Description
    Resume Exit_Button_Master_Species_Click
 
End Sub

Private Sub ButtonUnknown_Click()
    Dim stDocName As String
    Dim stLinkCriteria As String

    DoCmd.OpenForm "frm_List_Unknown", , , , , acDialog
    Me.Refresh
    
 '   stLinkCriteria = "[Species_ID]=" & "'" & Me![Shrub_ID] & "'"
 '   DoCmd.OpenForm stDocName, , , stLinkCriteria, , , Me![Shrub_ID]
End Sub

Private Sub ButtonA1_Click()

  If Screen.PreviousControl.name <> "Species" And Not IsNull(Me!Species) Then
    If IsNull(Screen.PreviousControl.Value) Then
      Screen.PreviousControl.Value = 1
    Else
      Screen.PreviousControl.Value = Screen.PreviousControl.Value + 1
    End If
  End If
  Screen.PreviousControl.SetFocus
End Sub

Private Sub ButtonA5_Click()
  If Screen.PreviousControl.name <> "Species" And Not IsNull(Me!Species) Then
    If IsNull(Screen.PreviousControl.Value) Then
      Screen.PreviousControl.Value = 5
    Else
      Screen.PreviousControl.Value = Screen.PreviousControl.Value + 5
    End If
  End If
  Screen.PreviousControl.SetFocus
End Sub

Private Sub ButtonS1_Click()
  If Screen.PreviousControl.name <> "Species" And Not IsNull(Me!Species) Then
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
End Sub

Private Sub ButtonS5_Click()
  If Screen.PreviousControl.name <> "Species" And Not IsNull(Me!Species) Then
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
End Sub

Private Sub ButtonZero_Click()
  If Screen.PreviousControl.name <> "Species" And Not IsNull(Me!Species) Then
      Screen.PreviousControl.Value = 0
  End If
  Screen.PreviousControl.SetFocus
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)

  '  If IsNull(Me.Parent!Observer) And IsNull(Me.Parent!Recorder) Then
  '    MsgBox "You must enter an observer or recorder first."
  '    DoCmd.CancelEvent
  '    SendKeys "{ESC}"
  '    GoTo Exit_Procedure
  '  End If
    ' Create the GUID primary key value
    If IsNull(Me!Shrub_ID) Then
        If GetDataType("tbl_LP_Shrub", "Shrub_ID") = dbText Then
            Me.Shrub_ID = fxnGUIDGen
        End If
    End If
Exit_Procedure:
End Sub

Private Sub Species_BeforeUpdate(Cancel As Integer)

    If Not IsNull(DLookup("[Shrub_ID]", "tbl_LP_Shrub", "[Transect_ID] = '" & Me!Transect_ID & "' AND [Species] = '" & Me!Species & "'")) Then
      MsgBox "This species is already recorded for this transect."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
    End If

End Sub

Private Sub Species_GotFocus()

    If IsNull(Me.Parent!Visit_Date) Then    ' If they didn't bother to enter a date, default to event date.
      Me.Parent!Visit_Date = Me.Parent.Parent!Start_Date
      Me.Parent.Refresh   ' Force save of transect record
    End If
   
End Sub
Private Sub ButtonDelete_Click()
On Error GoTo Err_ButtonDelete_Click
  Dim intReply As Integer
  
  intReply = MsgBox("Are you sure you want to delete this record?", vbYesNo, "Delete Record")
    If intReply = vbYes Then
      DoCmd.SetWarnings False
      DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
      DoCmd.DoMenuItem acFormBar, acEditMenu, 6, , acMenuVer70
      DoCmd.SetWarnings True
      Me.Requery
    End If
Exit_ButtonDelete_Click:
    Exit Sub

Err_ButtonDelete_Click:
    MsgBox Err.Description
    Resume Exit_ButtonDelete_Click
    
End Sub
