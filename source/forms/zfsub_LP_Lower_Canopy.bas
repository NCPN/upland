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
    Width =6240
    DatasheetFontHeight =9
    ItemSuffix =21
    Left =1230
    Top =75
    Right =7215
    Bottom =2445
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xd0cb4fa45156e340
    End
    RecordSource ="qry_LP_Lower_Canopy"
    Caption ="fsub_LP_Lower_Canopy"
    OnCurrent ="[Event Procedure]"
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
            Height =600
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2460
                    Top =60
                    Width =1500
                    Height =240
                    FontWeight =700
                    Name ="Label13"
                    Caption ="Lower Canopy"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1140
                    Top =360
                    Width =780
                    Height =240
                    Name ="Label14"
                    Caption ="Species"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2640
                    Top =360
                    Width =720
                    Height =240
                    Name ="Label15"
                    Caption ="Alive?"
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    Left =4440
                    Top =180
                    Width =540
                    Height =240
                    FontWeight =700
                    Name ="Label16"
                    Caption ="Point"
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4980
                    Top =180
                    Width =480
                    FontWeight =700
                    Name ="Point"
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
                    Width =300
                    Height =239
                    ColumnWidth =2310
                    Name ="LC_ID"
                    ControlSource ="LC_ID"
                    StatusBarText ="Unique record identifier - primary key"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =660
                    Top =60
                    Width =360
                    Height =239
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Intercept_ID"
                    ControlSource ="Intercept_ID"
                    StatusBarText ="Foreign key to tbl_LP_Intercept"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Top =60
                    Width =420
                    Height =239
                    ColumnWidth =600
                    TabIndex =2
                    Name ="Sequence"
                    ControlSource ="Sequence"
                    StatusBarText ="Lower canopy sequence number"
                End
                Begin ComboBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2880
                    Left =660
                    Top =60
                    Width =1800
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="Species"
                    ControlSource ="Species"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qryU_LP_Canopy.Master_Plant_Code, qryU_LP_Canopy.Utah_Species FROM qryU_L"
                        "P_Canopy WHERE (((qryU_LP_Canopy.Utah_Species) Is Not Null)); "
                    ColumnWidths ="0;2160"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =375
                    Left =2640
                    Top =60
                    Width =720
                    TabIndex =4
                    Name ="Alive"
                    ControlSource ="Alive"
                    RowSourceType ="Value List"
                    RowSource ="-1;\"Yes\";0;\"No\""
                    ColumnWidths ="0;375"
                    DefaultValue ="-1"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3540
                    Top =60
                    Width =1200
                    Height =300
                    TabIndex =5
                    Name ="ButtonLookup"
                    Caption ="Master Lookup"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4920
                    Top =60
                    Width =1215
                    Height =300
                    TabIndex =6
                    Name ="ButtonUnknown"
                    Caption ="Unknown Sp."
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
On Error GoTo Err_Button_Master_Species_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim strOpenArg As String

    strOpenArg = "fsub_LP_Lower_Canopy"
    stDocName = "frm_Master_Species"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , , strOpenArg

Exit_Button_Master_Species_Click:
    Exit Sub

Err_Button_Master_Species_Click:
    MsgBox Err.Description
    Resume Exit_Button_Master_Species_Click
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)

    Dim db As DAO.Database
    Dim Canopy As DAO.Recordset
    Dim strSQL As String
        
    On Error GoTo Err_Handler
    
    ' Set point number
    Set db = CurrentDb
    strSQL = "SELECT [Sequence] FROM [tbl_LP_Lower_Canopy] WHERE Intercept_ID = '" & Me!Intercept_ID & "' ORDER BY [Sequence] DESC"
    Set Canopy = db.OpenRecordset(strSQL)
    
    If Canopy.EOF Then
      Me![Sequence] = 1  ' First sequence will be one - duh.
    Else
      Canopy.MoveFirst
      Me![Sequence] = Canopy![Sequence] + 1
    End If
    Canopy.Close
    
    ' Create the GUID primary key value
    If IsNull(Me!LC_ID) Then
        If GetDataType("tbl_LP_Lower_Canopy", "LC_ID") = dbText Then
            Me.LC_ID = fxnGUIDGen
        End If
    End If

Exit_Procedure:

    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

Private Sub Form_Current()
      If Not IsNull(Me!Species) Or Me!Species <> "" Then
          If IsNull(DLookup("[Surface_Code]", "tlu_LP_Soil_Surface", "[Surface_Code] = '" & Me!Species & "'")) Then
            Me!Alive.Enabled = True
          Else
            Me!Alive.Enabled = False
          End If
      Else
        Me!Alive.Enabled = False
      End If
End Sub

Private Sub Species_AfterUpdate()
      If IsNull(Me!Species) Or Me!Species = "" Then
        Me!Alive = 0
        Me!Alive.Enabled = False
      Else
          If IsNull(DLookup("[Surface_Code]", "tlu_LP_Soil_Surface", "[Surface_Code] = '" & Me!Species & "'")) Then
            Me!Alive.Enabled = True
            Me!Alive = -1
          Else
            Me!Alive = 0
            Me!Alive.Enabled = False
          End If
      End If
End Sub

Private Sub Species_BeforeUpdate(Cancel As Integer)
    If Not IsNull(DLookup("[LC_ID]", "tbl_LP_Lower_Canopy", "[Intercept_ID] = '" & Me!Intercept_ID & "' AND [Species] = '" & Me!Species & "'")) Then
      MsgBox "This species is already recorded for this point."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
    End If
    ' Check for duplicate species in Soil Surface field.
    If Not IsNull(DLookup("[Intercept_ID]", "tbl_LP_Intercept", "[Intercept_ID] = '" & Me!Intercept_ID & "' AND [Top] = '" & Me!Species & "'")) Then
      MsgBox "This species is already recorded for this point."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
    End If
End Sub
Private Sub ButtonUnknown_Click()
On Error GoTo Err_ButtonUnknown_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_LP_Unknown_Species"
    
    stLinkCriteria = "[Species_ID]=" & "'" & Me![LC_ID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , , Me![LC_ID]

Exit_ButtonUnknown_Click:
    Exit Sub

Err_ButtonUnknown_Click:
    MsgBox Err.Description
    Resume Exit_ButtonUnknown_Click
    
End Sub
