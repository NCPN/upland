Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11520
    DatasheetFontHeight =9
    ItemSuffix =36
    Left =120
    Top =1245
    Right =11895
    Bottom =3675
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x43f03470521ee340
    End
    RecordSource ="tbl_Quadrat_Species"
    Caption ="fsub_Species"
    BeforeInsert ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
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
            FontWeight =700
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
                    Left =3480
                    Top =240
                    Width =660
                    Height =240
                    Name ="Alive_Label"
                    Caption ="Alive ?"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =4320
                    Width =2940
                    Height =240
                    Name ="Nested_Quad_Label"
                    Caption ="Nested quadrats (m2)"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =7440
                    Top =60
                    Width =840
                    Height =420
                    Name ="Percent_Cover_Label"
                    Caption ="Percent Cover"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    Left =8280
                    Top =60
                    Width =960
                    Height =420
                    Name ="Rooted_Outside_Label"
                    Caption ="Rooted Outside ?"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =60
                    Top =60
                    Width =2820
                    Height =420
                    Name ="Label14"
                    Caption ="Species- Select by State PLANT code or Species Name"
                End
                Begin Label
                    OverlapFlags =87
                    Left =4440
                    Top =240
                    Width =420
                    Height =240
                    Name ="Label23"
                    Caption ="0.01"
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    Left =5160
                    Top =240
                    Width =420
                    Height =240
                    Name ="Label24"
                    Caption ="0.1"
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    Left =6000
                    Top =240
                    Width =420
                    Height =240
                    Name ="Label25"
                    Caption ="1"
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    Left =6600
                    Top =240
                    Width =420
                    Height =240
                    Name ="Label26"
                    Caption ="10"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =2580
                    Top =60
                    Width =300
                    Name ="State_Code"

                End
            End
        End
        Begin Section
            Height =660
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =420
                    Height =255
                    ColumnWidth =2310
                    Name ="Species_ID"
                    ControlSource ="Species_ID"
                    StatusBarText ="Unique record identifier - primary key"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =600
                    Top =60
                    Width =360
                    Height =255
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Quadrat_ID"
                    ControlSource ="Quadrat_ID"
                    StatusBarText ="Foreign key to tbl_Quadrat_Transect"

                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7620
                    Top =60
                    Width =540
                    Height =255
                    ColumnWidth =465
                    TabIndex =9
                    Name ="Percent_Cover"
                    ControlSource ="Percent_Cover"
                    StatusBarText ="Percent cover in 10 m2 quadrat"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =8700
                    Top =120
                    TabIndex =10
                    Name ="Rooted_Outside"
                    ControlSource ="Rooted_Outside"
                    StatusBarText ="Is plant rooted outside quadrat?"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =4560
                    Top =60
                    Width =240
                    TabIndex =5
                    Name ="Q1"
                    ControlSource ="Q1"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =5340
                    Top =60
                    Width =240
                    Height =180
                    TabIndex =6
                    Name ="Q2"
                    ControlSource ="Q2"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =6120
                    Top =60
                    Width =300
                    TabIndex =7
                    Name ="Q3"
                    ControlSource ="Q3"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =6780
                    Top =60
                    Width =300
                    TabIndex =8
                    Name ="Q4"
                    ControlSource ="Q4"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =7020
                    Top =60
                    Width =300
                    TabIndex =11
                    Name ="Nested_Quad"
                    ControlSource ="Nested_Quad"

                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =10680
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
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9180
                    Top =60
                    Height =300
                    TabIndex =13
                    Name ="ButtonUnknown"
                    Caption ="Unknown Species"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =4
                    ListWidth =10800
                    Left =60
                    Top =60
                    Width =1260
                    TabIndex =2
                    BoundColumn =2
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"40\""
                    Name ="cbo_Code"
                    ControlSource ="Plant_Code"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_NCPN_Plants.Utah_PLANT_Code, tlu_NCPN_Plants.Utah_Species, tlu_NCPN_P"
                        "lants.Master_PLANT_Code, tlu_NCPN_Plants.Master_Species FROM tlu_NCPN_Plants WHE"
                        "RE (((tlu_NCPN_Plants.Utah_PLANT_Code) Is Not Null)) ORDER BY tlu_NCPN_Plants.Ut"
                        "ah_PLANT_Code; "
                    ColumnWidths ="1800;3600;1800;3600"
                    OnGotFocus ="[Event Procedure]"
                    ControlTipText ="Select using State PLANTS Code here or State Species Code below."

                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    ColumnCount =4
                    ListWidth =10800
                    Left =60
                    Top =360
                    Width =2640
                    TabIndex =3
                    BoundColumn =2
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="cbo_Species"
                    ControlSource ="Plant_Code"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_NCPN_Plants.Utah_Species, tlu_NCPN_Plants.Utah_PLANT_Code, tlu_NCPN_P"
                        "lants.Master_PLANT_Code, tlu_NCPN_Plants.Master_Species FROM tlu_NCPN_Plants WHE"
                        "RE (((tlu_NCPN_Plants.Utah_Species) Is Not Null And (tlu_NCPN_Plants.Utah_Specie"
                        "s)<>\" \")) ORDER BY tlu_NCPN_Plants.Utah_Species; "
                    ColumnWidths ="3600;1800;1800;3600"
                    OnGotFocus ="[Event Procedure]"
                    ControlTipText ="Select using State Species Code here or State PLANTS Code above."

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =87
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4020
                    Top =360
                    Width =780
                    TabIndex =14
                    ForeColor =16711680
                    Name ="Master_Code"
                    ControlSource ="Plant_Code"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3060
                            Top =360
                            Width =960
                            Height =240
                            FontWeight =400
                            Name ="Label30"
                            Caption ="Master Code"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =87
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6060
                    Top =360
                    Width =2820
                    TabIndex =15
                    ForeColor =16711680
                    Name ="Text31"
                    ControlSource ="=DLookUp(\"[Master_Species]\",\"tlu_NCPN_Plants\",\"[Master_PLANT_Code] = '\" & "
                        "[Master_Code] & \"'\")"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4860
                            Top =360
                            Width =1200
                            Height =240
                            FontWeight =400
                            Name ="Label32"
                            Caption ="Master Species"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =1440
                    Top =60
                    Width =1335
                    Height =300
                    TabIndex =16
                    Name ="Button_Master_Species"
                    Caption ="Master Lookup"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =224
                    Left =3480
                    Top =60
                    Width =660
                    TabIndex =4
                    Name ="Alive"
                    ControlSource ="Alive"
                    RowSourceType ="Value List"
                    RowSource ="\"Yes\";\"No\""
                    ColumnWidths ="224"

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

Private Sub cbo_Code_GotFocus()
  If Me!State_Code = "UT" Then
    Me!cbo_Code.RowSource = "Select tlu_NCPN_Plants.Utah_PLANT_Code, tlu_NCPN_Plants.Utah_Species, tlu_NCPN_Plants.Master_PLANT_Code, tlu_NCPN_Plants.Master_Species FROM tlu_NCPN_Plants WHERE (((tlu_NCPN_Plants.Utah_PLANT_Code) Is Not Null)) ORDER BY tlu_NCPN_Plants.Utah_PLANT_Code"
    Me!cbo_Species.RowSource = "SELECT tlu_NCPN_Plants.Utah_Species, tlu_NCPN_Plants.Utah_PLANT_Code, tlu_NCPN_Plants.Master_PLANT_Code, tlu_NCPN_Plants.Master_Species FROM tlu_NCPN_Plants WHERE (((tlu_NCPN_Plants.Utah_Species) Is Not Null And (tlu_NCPN_Plants.Utah_Species)<>' ')) ORDER BY tlu_NCPN_Plants.Utah_Species"
  ElseIf Me!State_Code = "CO" Then
    Me!cbo_Code.RowSource = "Select tlu_NCPN_Plants.Co_PLANT_Code, tlu_NCPN_Plants.Co_Species, tlu_NCPN_Plants.Master_PLANT_Code, tlu_NCPN_Plants.Master_Species FROM tlu_NCPN_Plants WHERE (((tlu_NCPN_Plants.Co_PLANT_Code) Is Not Null)) ORDER BY tlu_NCPN_Plants.Co_PLANT_Code"
    Me!cbo_Species.RowSource = "SELECT tlu_NCPN_Plants.Co_Species, tlu_NCPN_Plants.Co_PLANT_Code, tlu_NCPN_Plants.Master_PLANT_Code, tlu_NCPN_Plants.Master_Species FROM tlu_NCPN_Plants WHERE (((tlu_NCPN_Plants.Co_Species) Is Not Null And (tlu_NCPN_Plants.Co_Species)<>' ')) ORDER BY tlu_NCPN_Plants.Co_Species"
  Else
    Me!cbo_Code.RowSource = "Select tlu_NCPN_Plants.Wy_PLANT_Code, tlu_NCPN_Plants.Wy_Species, tlu_NCPN_Plants.Master_PLANT_Code, tlu_NCPN_Plants.Master_Species FROM tlu_NCPN_Plants WHERE (((tlu_NCPN_Plants.Wy_PLANT_Code) Is Not Null)) ORDER BY tlu_NCPN_Plants.Wy_PLANT_Code"
    Me!cbo_Species.RowSource = "SELECT tlu_NCPN_Plants.Wy_Species, tlu_NCPN_Plants.Wy_PLANT_Code, tlu_NCPN_Plants.Master_PLANT_Code, tlu_NCPN_Plants.Master_Species FROM tlu_NCPN_Plants WHERE (((tlu_NCPN_Plants.Wy_Species) Is Not Null And (tlu_NCPN_Plants.Wy_Species)<>' ')) ORDER BY tlu_NCPN_Plants.Wy_Species"
  End If
End Sub

Private Sub cbo_Species_GotFocus()
  If Me!State_Code = "UT" Then
    Me!cbo_Code.RowSource = "Select tlu_NCPN_Plants.Utah_PLANT_Code, tlu_NCPN_Plants.Utah_Species, tlu_NCPN_Plants.Master_PLANT_Code, tlu_NCPN_Plants.Master_Species FROM tlu_NCPN_Plants WHERE (((tlu_NCPN_Plants.Utah_PLANT_Code) Is Not Null)) ORDER BY tlu_NCPN_Plants.Utah_PLANT_Code"
    Me!cbo_Species.RowSource = "SELECT tlu_NCPN_Plants.Utah_Species, tlu_NCPN_Plants.Utah_PLANT_Code, tlu_NCPN_Plants.Master_PLANT_Code, tlu_NCPN_Plants.Master_Species FROM tlu_NCPN_Plants WHERE (((tlu_NCPN_Plants.Utah_Species) Is Not Null And (tlu_NCPN_Plants.Utah_Species)<>' ')) ORDER BY tlu_NCPN_Plants.Utah_Species"
  ElseIf Me!State_Code = "CO" Then
    Me!cbo_Code.RowSource = "Select tlu_NCPN_Plants.Co_PLANT_Code, tlu_NCPN_Plants.Co_Species, tlu_NCPN_Plants.Master_PLANT_Code, tlu_NCPN_Plants.Master_Species FROM tlu_NCPN_Plants WHERE (((tlu_NCPN_Plants.Co_PLANT_Code) Is Not Null)) ORDER BY tlu_NCPN_Plants.Co_PLANT_Code"
    Me!cbo_Species.RowSource = "SELECT tlu_NCPN_Plants.Co_Species, tlu_NCPN_Plants.Co_PLANT_Code, tlu_NCPN_Plants.Master_PLANT_Code, tlu_NCPN_Plants.Master_Species FROM tlu_NCPN_Plants WHERE (((tlu_NCPN_Plants.Co_Species) Is Not Null And (tlu_NCPN_Plants.Co_Species)<>' ')) ORDER BY tlu_NCPN_Plants.Co_Species"
  Else
    Me!cbo_Code.RowSource = "Select tlu_NCPN_Plants.Wy_PLANT_Code, tlu_NCPN_Plants.Wy_Species, tlu_NCPN_Plants.Master_PLANT_Code, tlu_NCPN_Plants.Master_Species FROM tlu_NCPN_Plants WHERE (((tlu_NCPN_Plants.Wy_PLANT_Code) Is Not Null)) ORDER BY tlu_NCPN_Plants.Wy_PLANT_Code"
    Me!cbo_Species.RowSource = "SELECT tlu_NCPN_Plants.Wy_Species, tlu_NCPN_Plants.Wy_PLANT_Code, tlu_NCPN_Plants.Master_PLANT_Code, tlu_NCPN_Plants.Master_Species FROM tlu_NCPN_Plants WHERE (((tlu_NCPN_Plants.Wy_Species) Is Not Null And (tlu_NCPN_Plants.Wy_Species)<>' ')) ORDER BY tlu_NCPN_Plants.Wy_Species"
  End If
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    On Error GoTo Err_Handler

    If IsNull(Me.Parent!Recorder) And IsNull(Me.Parent!Observer) Then
      MsgBox "You must enter Observer or Recorder first."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
      GoTo Exit_Procedure
    End If
    ' Create the GUID primary key value
    If IsNull(Me!Species_ID) Then
        If GetDataType("tbl_Quadrat_Species", "species_ID") = dbText Then
            Me.Species_ID = fxnGUIDGen
        End If
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

Private Sub Plant_Code_NotInList(NewData As String, Response As Integer)
  Dim IReply As Integer
  
  IReply = MsgBox("Species Does Not Exist.  Add(Yes/No)?", vbYesNo + vbQuestion)
  If IReply = vbYes Then
  ' Add observer form must be opened with acDialog option to make VB wait for
  ' add observer form to close before continuing on to next instruction.
    DoCmd.OpenForm "frm_add_Species", , , , , acDialog, NewData
    Response = acDataErrAdded

  Else
    Response = acDataErrContinue
  End If

End Sub

Private Sub ButtonDelete_Click()
On Error GoTo Err_ButtonDelete_Click
  Dim Reply As Integer
  Reply = MsgBox("Are you sure you want to delete this record?", vbYesNo, "Species Delete")
  If Reply = 6 Then
    DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
    DoCmd.DoMenuItem acFormBar, acEditMenu, 6, , acMenuVer70
  End If
Exit_ButtonDelete_Click:
    Exit Sub

Err_ButtonDelete_Click:
    MsgBox Err.Description
    Resume Exit_ButtonDelete_Click
    
End Sub
Private Sub ButtonUnknown_Click()
On Error GoTo Err_ButtonUnknown_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_LP_Unknown_Species"
    
    stLinkCriteria = "[Species_ID]=" & "'" & Me![Species_ID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , , Me![Species_ID]

Exit_ButtonUnknown_Click:
    Exit Sub

Err_ButtonUnknown_Click:
    MsgBox Err.Description
    Resume Exit_ButtonUnknown_Click
    
End Sub
Private Sub Button_Master_Species_Click()
On Error GoTo Err_Button_Master_Species_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    
    strOpenArg = "fsub_Species"
    stDocName = "frm_Master_Species"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , , strOpenArg

Exit_Button_Master_Species_Click:
    Exit Sub

Err_Button_Master_Species_Click:
    MsgBox Err.Description
    Resume Exit_Button_Master_Species_Click
    
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    
        Dim db As DAO.Database
        Dim Species As DAO.Recordset
        Dim strSQL As String
        
    On Error GoTo Err_Handler
    
    ' Check for duplicate species
    Set db = CurrentDb
    strSQL = "SELECT [Species_ID] FROM [tbl_Quadrat_Species] WHERE Quadrat_ID = '" & Me!Quadrat_ID & "' AND Plant_Code = '" & Me!Master_Code & "' AND Alive = '" & Me!Alive & "'"
    Set Species = db.OpenRecordset(strSQL)
    If Not Species.EOF Then
      MsgBox "This species has already been recorded for this quadrat."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
    End If

Exit_Procedure:
    Species.Close
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
    
End Sub

Private Sub Percent_Cover_AfterUpdate()
  If Not IsNull(Me!Percent_Cover) Then
    If Not IsNumeric(Me!Percent_Cover) Then
      If Me!Percent_Cover <> "T" Then
        MsgBox "Percent cover must be between 1 and 100 or = 'T'"
        DoCmd.CancelEvent
        SendKeys "{ESC}"
        GoTo Exit_Procedure
      End If
    Else
      If (Me!Percent_Cover < 1) Or (Me!Percent_Cover > 100) Then
        MsgBox "Percent cover must be between 1 and 100 or = 'T'"
        DoCmd.CancelEvent
        SendKeys "{ESC}"
        GoTo Exit_Procedure
      End If
    End If
  End If
Exit_Procedure:
End Sub

Private Sub Q1_AfterUpdate()
  If Me!Q1 = -1 Then
    Me!Q2 = 0
    Me!Q3 = 0
    Me!Q4 = 0
  End If
  
End Sub

Private Sub Q2_AfterUpdate()
  If Me!Q2 = -1 Then
    Me!Q1 = 0
    Me!Q3 = 0
    Me!Q4 = 0
  End If
End Sub

Private Sub Q3_AfterUpdate()
  If Me!Q3 = -1 Then
    Me!Q2 = 0
    Me!Q1 = 0
    Me!Q4 = 0
  End If
End Sub

Private Sub Q4_AfterUpdate()
  If Me!Q4 = -1 Then
    Me!Q2 = 0
    Me!Q3 = 0
    Me!Q1 = 0
  End If
End Sub
