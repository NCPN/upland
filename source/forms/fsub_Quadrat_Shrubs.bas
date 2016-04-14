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
    Cycle =1
    GridX =24
    GridY =24
    Width =9180
    DatasheetFontHeight =9
    ItemSuffix =34
    Left =-2175
    Top =10965
    Right =6780
    Bottom =12945
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xf02518bf4b1ee340
    End
    RecordSource ="tbl_Quadrat_Shrubs"
    Caption ="fsub_Quadrat_Shrubs"
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
            OldBorderStyle =1
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
            Height =540
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    Left =60
                    Top =60
                    Width =2640
                    Height =480
                    FontWeight =700
                    Name ="Plant_Code_Label"
                    Caption ="Species- Select by State PLANT code or Species Name"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    BorderWidth =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =2760
                    Top =300
                    Width =900
                    Height =240
                    FontWeight =700
                    Name ="0cm_Label"
                    Caption ="0-10 cm"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl0cm_Label"
                End
                Begin Label
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =3660
                    Top =300
                    Width =900
                    Height =240
                    FontWeight =700
                    Name ="10cm_Label"
                    Caption ="10-25 cm"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl10cm_Label"
                End
                Begin Label
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =4560
                    Top =300
                    Width =900
                    Height =240
                    FontWeight =700
                    Name ="25cm_Label"
                    Caption ="25-50 cm"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl25cm_Label"
                End
                Begin Label
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =5460
                    Top =300
                    Width =900
                    Height =240
                    FontWeight =700
                    Name ="50cm_Label"
                    Caption ="50-100 cm"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl50cm_Label"
                End
                Begin Label
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =6360
                    Top =300
                    Width =900
                    Height =240
                    FontWeight =700
                    Name ="100cm_Label"
                    Caption ="1-2 m"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl100cm_Label"
                End
                Begin Label
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =7260
                    Top =300
                    Width =900
                    Height =240
                    FontWeight =700
                    Name ="200cm_Label"
                    Caption =">2 m"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Ctl200cm_Label"
                End
                Begin Label
                    BorderWidth =1
                    OverlapFlags =87
                    TextAlign =2
                    Left =2760
                    Top =60
                    Width =5400
                    Height =240
                    FontWeight =700
                    Name ="Label21"
                    Caption ="Height Classes"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =2040
                    Top =120
                    Width =360
                    Height =300
                    Name ="State_Code"

                End
            End
        End
        Begin Section
            Height =720
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
                    Name ="Shrub_ID"
                    ControlSource ="Shrub_ID"
                    StatusBarText ="Unique record identifier - primary key"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =540
                    Top =60
                    Width =420
                    Height =255
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Quadrat_ID"
                    ControlSource ="Quadrat_ID"
                    StatusBarText ="Foreign key to tbl_Quadrat_Transect"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2940
                    Top =60
                    Width =479
                    Height =255
                    ColumnWidth =600
                    TabIndex =4
                    Name ="0cm"
                    ControlSource ="0cm"
                    StatusBarText ="number of shrubs 0-10 cm"
                    EventProcPrefix ="Ctl0cm"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3840
                    Top =60
                    Width =479
                    Height =255
                    ColumnWidth =600
                    TabIndex =5
                    Name ="10cm"
                    ControlSource ="10cm"
                    StatusBarText ="number of shrubs 10-25 cm"
                    EventProcPrefix ="Ctl10cm"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4740
                    Top =60
                    Width =479
                    Height =255
                    ColumnWidth =600
                    TabIndex =6
                    Name ="25cm"
                    ControlSource ="25cm"
                    StatusBarText ="number of shrubs 25-50 cm"
                    EventProcPrefix ="Ctl25cm"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5640
                    Top =60
                    Width =479
                    Height =255
                    ColumnWidth =600
                    TabIndex =7
                    Name ="50cm"
                    ControlSource ="50cm"
                    StatusBarText ="number of shrubs 50-100 cm"
                    EventProcPrefix ="Ctl50cm"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6540
                    Top =60
                    Width =479
                    Height =255
                    ColumnWidth =600
                    TabIndex =8
                    Name ="100cm"
                    ControlSource ="100cm"
                    StatusBarText ="number of shrubs 1-2 m"
                    EventProcPrefix ="Ctl100cm"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7440
                    Top =60
                    Width =479
                    Height =255
                    ColumnWidth =600
                    TabIndex =9
                    Name ="200cm"
                    ControlSource ="200cm"
                    StatusBarText ="number of shrubs >2 m"
                    EventProcPrefix ="Ctl200cm"

                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =8220
                    Top =60
                    Width =705
                    Height =300
                    TabIndex =10
                    ForeColor =255
                    Name ="ButtonDelete"
                    Caption ="Delete"
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
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    ListWidth =10800
                    Left =60
                    Top =420
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
                    Left =3900
                    Top =420
                    Width =780
                    TabIndex =11
                    ForeColor =16711680
                    Name ="Master_Code"
                    ControlSource ="Plant_Code"

                    Begin
                        Begin Label
                            OldBorderStyle =0
                            OverlapFlags =93
                            Left =2940
                            Top =420
                            Width =960
                            Height =240
                            ForeColor =0
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
                    Left =5940
                    Top =420
                    Width =2820
                    TabIndex =12
                    ForeColor =16711680
                    Name ="Text31"
                    ControlSource ="=DLookUp(\"[Master_Species]\",\"tlu_NCPN_Plants\",\"[Master_PLANT_Code] = '\" & "
                        "[Master_Code] & \"'\")"

                    Begin
                        Begin Label
                            OldBorderStyle =0
                            OverlapFlags =93
                            Left =4740
                            Top =420
                            Width =1200
                            Height =240
                            ForeColor =0
                            Name ="Label32"
                            Caption ="Master Species"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1440
                    Top =60
                    Width =1320
                    Height =300
                    FontSize =6
                    TabIndex =13
                    Name ="Button_Master_Species"
                    Caption ="Master Lookup"
                    OnClick ="[Event Procedure]"

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
    If IsNull(Me!Shrub_ID) Then
        If GetDataType("tbl_Quadrat_Shrubs", "Shrub_ID") = dbText Then
            Me.Shrub_ID = fxnGUIDGen
        End If
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

Private Sub Button_Master_Species_Click()
On Error GoTo Err_Button_Master_Species_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim strOpenArg As String

    strOpenArg = "fsub_Quadrat_Shrubs"
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
    strSQL = "SELECT [Shrub_ID] FROM [tbl_Quadrat_Shrubs] WHERE Quadrat_ID = '" & Me!Quadrat_ID & "' AND Plant_Code = '" & Me!Master_Code & "'"
    Set Species = db.OpenRecordset(strSQL)
    If Not Species.EOF Then
      MsgBox "This shrub has already been recorded for this quadrat."
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
