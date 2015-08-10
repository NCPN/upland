Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    TabularFamily =124
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10080
    DatasheetFontHeight =9
    ItemSuffix =22
    Left =2070
    Top =300
    Right =12885
    Bottom =2535
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xfac8c09cb286e340
    End
    RecordSource ="tbl_Monument"
    Caption ="fsub_FW_Monument"
    DatasheetFontName ="Arial"
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
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
            Height =300
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =120
                    Top =60
                    Width =960
                    Height =240
                    FontWeight =700
                    Name ="Monument_Code_Label"
                    Caption ="Monument"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1200
                    Top =60
                    Width =660
                    Height =240
                    FontWeight =700
                    Name ="Tag_No_Label"
                    Caption ="Tag #"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1980
                    Top =60
                    Width =840
                    Height =240
                    FontWeight =700
                    Name ="Species_Label"
                    Caption ="Species"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =4920
                    Top =60
                    Width =975
                    Height =240
                    FontWeight =700
                    Name ="DBH_Label"
                    Caption ="DBH (cm)"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =5940
                    Top =60
                    Width =1335
                    Height =240
                    FontWeight =700
                    Name ="Bearing_Label"
                    Caption ="Bearing (deg)"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =7260
                    Top =60
                    Width =2040
                    Height =240
                    FontWeight =700
                    Name ="Rebar_Distance_Label"
                    Caption ="Distance to Rebar (m)"
                    Tag ="DetachedLabel"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2940
                    Width =780
                    Height =300
                    ForeColor =32768
                    Name ="ButtonLookup"
                    Caption ="Lookup"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3780
                    Width =930
                    Height =300
                    TabIndex =1
                    ForeColor =32768
                    Name ="ButtonUnknown"
                    Caption ="Unknown"
                    OnClick ="[Event Procedure]"
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
                    Name ="Monument_ID"
                    ControlSource ="Monument_ID"
                    StatusBarText ="Master identifier"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1320
                    Top =60
                    Width =480
                    Height =255
                    ColumnWidth =600
                    TabIndex =2
                    Name ="Tag_No"
                    ControlSource ="Tag_No"
                    StatusBarText ="F/W Tag number of monument tree"
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5175
                    Top =60
                    Width =480
                    Height =255
                    ColumnWidth =2310
                    TabIndex =4
                    Name ="DBH"
                    ControlSource ="DBH"
                    StatusBarText ="F/W Diameter at breast height iin centimeters"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6360
                    Top =60
                    Width =540
                    Height =255
                    ColumnWidth =600
                    TabIndex =5
                    Name ="Bearing"
                    ControlSource ="Bearing"
                    StatusBarText ="F/W Bearing from monument tree to plot corner in degrees"
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7980
                    Top =60
                    Width =540
                    Height =255
                    ColumnWidth =2310
                    TabIndex =6
                    Name ="Rebar_Distance"
                    ControlSource ="Rebar_Distance"
                    StatusBarText ="F/W Distance from center point of minument tree to plot corner in meters"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    ListWidth =615
                    Left =180
                    Top =60
                    Width =960
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"10\";\"10\""
                    Name ="Monument_Code"
                    ControlSource ="Monument_Code"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Monument_Code.Monument_Code FROM tlu_Monument_Code; "
                    ColumnWidths ="615"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =6480
                    Left =2040
                    Top =60
                    Width =1979
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="Species"
                    ControlSource ="Species"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qryU_Top_Canopy.Master_PLANT_Code, qryU_Top_Canopy.LU_Code, qryU_Top_Cano"
                        "py.Utah_Species FROM qryU_Top_Canopy WHERE (((qryU_Top_Canopy.Utah_Species) Is N"
                        "ot Null)) ORDER BY qryU_Top_Canopy.LU_Code; "
                    ColumnWidths ="0;2160;4320"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4320
                    Top =60
                    Width =480
                    TabIndex =7
                    Name ="Location_ID"
                    ControlSource ="Location_ID"
                    StatusBarText ="Foreign key to tbl_Locations"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9180
                    Top =60
                    Width =705
                    Height =300
                    TabIndex =8
                    ForeColor =255
                    Name ="ButtonDelete"
                    Caption ="Delete"
                    OnClick ="[Event Procedure]"
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

    DoCmd.OpenForm "frm_List_Unknown", , , , , acDialog
    Me!Species.Requery
'    Me.Refresh

Exit_ButtonUnknown_Click:
    Exit Sub

Err_ButtonUnknown_Click:
    MsgBox Err.Description
    Resume Exit_ButtonUnknown_Click
    
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
