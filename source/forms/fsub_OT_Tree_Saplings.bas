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
    Left =1260
    Top =2010
    Right =12930
    Bottom =5520
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x384b3f359387e340
    End
    RecordSource ="tbl_OT_Tree_Saplings"
    Caption ="fsub_LP_Belt_Shrub"
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
                    Name ="ButtonMaster"
                    Caption ="Master Lookup"
                    OnClick ="[Event Procedure]"

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
                    Name ="ButtonUnknown"
                    Caption ="Unknown Species"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =7500
                    LayoutCachedTop =540
                    LayoutCachedWidth =9045
                    LayoutCachedHeight =840
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Left =9180
                    Top =240
                    Width =2100
                    Height =480
                    BackColor =6750207
                    Name ="rctNoData"
                    OnClick ="[Event Procedure]"
                    LayoutCachedLeft =9180
                    LayoutCachedTop =240
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =720
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =9300
                    Top =390
                    Width =300
                    TabIndex =2
                    Name ="cbxNoData"
                    ControlTipText ="No tree saplings found"

                    LayoutCachedLeft =9300
                    LayoutCachedTop =390
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =630
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =9530
                            Top =360
                            Width =1650
                            Height =240
                            FontWeight =600
                            Name ="lblNoData"
                            Caption ="No Species Found"
                            ControlTipText ="No tree saplings found"
                            LayoutCachedLeft =9530
                            LayoutCachedTop =360
                            LayoutCachedWidth =11180
                            LayoutCachedHeight =600
                        End
                    End
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
                    IMESentenceMode =3
                    Left =3855
                    Top =60
                    Width =540
                    Height =255
                    ColumnWidth =600
                    TabIndex =4
                    Name ="HC25"
                    ControlSource ="D25"
                    StatusBarText ="10.1-25cm height class total"

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
                    TabIndex =5
                    Name ="HC50"
                    ControlSource ="D51"
                    StatusBarText ="25.1-50cm height class total"

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
                    TabIndex =6
                    Name ="HC100"
                    ControlSource ="D101"
                    StatusBarText ="50.1-100cm height class total"

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
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =6480
                    Left =60
                    Top =60
                    Width =2304
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="Species"
                    ControlSource ="Species"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qryU_Top_Canopy.Master_PLANT_Code, qryU_Top_Canopy.LU_Code, qryU_Top_Cano"
                        "py.Utah_Species, Lifeform FROM qryU_Top_Canopy WHERE (((qryU_Top_Canopy.Utah_Spe"
                        "cies) Is Not Null)) AND Lifeform = 'Tree' ORDER BY qryU_Top_Canopy.LU_Code;"
                    ColumnWidths ="0;2160;4320"
                    BeforeUpdate ="[Event Procedure]"

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
                    Name ="ButtonA1"
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
                    Name ="ButtonA5"
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
                    Name ="ButtonS1"
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
                    Name ="ButtonS5"
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
                    Name ="ButtonZero"
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

Private Sub Alive_BeforeUpdate(Cancel As Integer)
    If Not IsNull(DLookup("[TS_ID]", "tbl_OT_Tree_Saplings", "[Event_ID] = '" & Me!Event_ID & "' AND [Species] = '" & Me!Species & "' AND [Alive] = " & Me!Alive)) Then
      MsgBox "This species is already recorded for this transect."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
    End If
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

Private Sub ButtonUnknown_Click()
    Dim stDocName As String
    Dim stLinkCriteria As String

    DoCmd.OpenForm "frm_List_Unknown", , , , , acDialog
    Me.Refresh
End Sub

Private Sub ButtonZero_Click()
  If Screen.PreviousControl.name <> "Species" And Not IsNull(Me!Species) Then
      Screen.PreviousControl.Value = 0
  End If
  Screen.PreviousControl.SetFocus
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)

    On Error GoTo Err_Handler

    ' Make sure there is an events record
    If IsNull(Me.Parent!Start_Date) Then
      MsgBox "Missing site visit date."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
      GoTo Exit_Procedure
    End If
    ' Create the GUID primary key value
    If IsNull(Me!TS_ID) Then
        If GetDataType("tbl_OT_Tree_Saplings", "TS_ID") = dbText Then
            Me.TS_ID = fxnGUIDGen
        End If
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

Private Sub Species_BeforeUpdate(Cancel As Integer)
    Dim Reply As Integer

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
