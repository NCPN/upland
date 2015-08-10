Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    TabularFamily =124
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10080
    DatasheetFontHeight =9
    ItemSuffix =22
    Left =1575
    Top =300
    Right =11640
    Bottom =1800
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x996fa5868b0fe340
    End
    RecordSource ="tbl_Soil_Profile"
    Caption ="fsub_Soil_Profile"
    BeforeInsert ="[Event Procedure]"
    DatasheetFontName ="Arial"
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
            Height =240
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2280
                    Width =900
                    Height =240
                    FontWeight =700
                    Name ="Texture_Label"
                    Caption ="Texture"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =4440
                    Width =1320
                    Height =240
                    FontWeight =700
                    Name ="Rock_fragments_Label"
                    Caption ="Rock frag qty"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =7500
                    Width =1380
                    Height =240
                    FontWeight =700
                    Name ="Effervescence_Label"
                    Caption ="Effervescence"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5940
                    Width =1320
                    Height =240
                    FontWeight =700
                    Name ="Label20"
                    Caption ="Rock frag size"
                    FontName ="Tahoma"
                End
            End
        End
        Begin Section
            Height =465
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =600
                    Height =255
                    ColumnWidth =2310
                    Name ="Soil_Profile_ID"
                    ControlSource ="Soil_Profile_ID"
                    StatusBarText ="Unique record identifier"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =720
                    Top =60
                    Width =540
                    Height =255
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Location_ID"
                    ControlSource ="Location_ID"
                    StatusBarText ="M. Link to tbl_Locations (Loc_ID)"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1440
                    Top =60
                    Width =2760
                    Height =239
                    ColumnWidth =3000
                    TabIndex =3
                    Name ="Texture"
                    ControlSource ="Texture"
                    StatusBarText ="Soil texture"
                    FontName ="Tahoma"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    ListWidth =840
                    Left =180
                    Top =60
                    Width =1140
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"10\";\"18\""
                    Name ="Depth"
                    ControlSource ="Depth"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Profile_Depth.Depth FROM tlu_Profile_Depth; "
                    ColumnWidths ="840"
                    FontName ="Tahoma"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2085
                    Left =4500
                    Top =60
                    Width =1200
                    TabIndex =4
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"18\""
                    Name ="Rock_Fragment_Qty"
                    ControlSource ="Rock_Fragment_Qty"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Rock_Frag_Q.Quantity, tlu_Rock_Frag_Q.Descriptor FROM tlu_Rock_Frag_Q"
                        "; "
                    ColumnWidths ="900;1185"
                    FontName ="Tahoma"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =885
                    Left =7500
                    Top =60
                    Width =1380
                    Height =239
                    TabIndex =6
                    ColumnInfo ="\"\";\"\";\"10\";\"22\""
                    Name ="Effervescence"
                    ControlSource ="Effervescence"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Effervescence.Effervescence FROM tlu_Effervescence; "
                    ColumnWidths ="885"
                    FontName ="Tahoma"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =795
                    Left =6000
                    Top =60
                    Width =1200
                    TabIndex =5
                    ColumnInfo ="\"\";\"\";\"10\";\"16\""
                    Name ="Rock_Fragment_Size"
                    ControlSource ="Rock_Fragment_Size"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Rock_Frag_Size.Size FROM tlu_Rock_Frag_Size; "
                    ColumnWidths ="795"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9180
                    Top =60
                    Width =705
                    Height =300
                    TabIndex =7
                    ForeColor =255
                    Name ="ButtonDelete"
                    Caption ="Delete"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
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

Private Sub Form_BeforeInsert(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Create the GUID primary key value
    If IsNull(Me!Soil_Profile_ID) Then
        If GetDataType("tbl_Soil_Profile", "Soil_Profile_ID") = dbText Then
            Me!Soil_Profile_ID = fxnGUIDGen
        End If
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
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
