Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    TabularFamily =126
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11520
    DatasheetFontHeight =9
    ItemSuffix =34
    Left =645
    Top =345
    Right =12165
    Bottom =7650
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xb0c0f4149355e340
    End
    RecordSource ="qry_LP_Intercept"
    Caption ="fsub_LP_Intercept"
    OnCurrent ="[Event Procedure]"
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
            Height =360
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =60
                    Top =60
                    Width =840
                    Height =240
                    FontWeight =700
                    Name ="Point_Label"
                    Caption ="Point (m)"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1380
                    Top =60
                    Width =1020
                    Height =240
                    FontWeight =700
                    Name ="Top_Label"
                    Caption ="Top Canopy"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =4140
                    Top =60
                    Width =660
                    Height =240
                    FontWeight =700
                    Name ="Alive_Label"
                    Caption ="Alive?"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =8220
                    Top =60
                    Width =1200
                    Height =240
                    FontWeight =700
                    Name ="Surface_Label"
                    Caption ="Soil surface"
                    Tag ="DetachedLabel"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6360
                    Top =60
                    Height =300
                    ForeColor =8421376
                    Name ="ButtonInitialize"
                    Caption ="Initialize Form"
                    OnClick ="[Event Procedure]"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =10560
                    Top =60
                    Width =660
                    Height =240
                    FontWeight =700
                    Name ="Label33"
                    Caption ="Alive?"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =360
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =600
                    Top =60
                    Width =420
                    Height =255
                    ColumnWidth =2310
                    Name ="Intercept_ID"
                    ControlSource ="Intercept_ID"
                    StatusBarText ="Unique record identifier - primary key"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =255
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =840
                    Height =255
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Point"
                    ControlSource ="Point"
                    Format ="General Number"
                    StatusBarText ="Intercept point - increments of .5m up to 50.0"
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =5760
                    Left =8220
                    Top =60
                    Width =2100
                    TabIndex =6
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="Surface"
                    ControlSource ="Surface"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qryU_LP_Soil_Surface.Master_Plant_Code, qryU_LP_Soil_Surface.Utah_Plant_C"
                        "ode, qryU_LP_Soil_Surface.Utah_Species FROM qryU_LP_Soil_Surface; "
                    ColumnWidths ="0;3312;2448"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin ComboBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2880
                    Left =1080
                    Top =60
                    Width =2880
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Top"
                    ControlSource ="Top"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_NCPN_Plants.Master_PLANT_Code, tlu_NCPN_Plants.Utah_Species FROM tlu_"
                        "NCPN_Plants WHERE (((tlu_NCPN_Plants.Utah_Species) Is Not Null And (tlu_NCPN_Pla"
                        "nts.Utah_Species)<>\" \" And (tlu_NCPN_Plants.Utah_Species)<>\"\")) ORDER BY tlu"
                        "_NCPN_Plants.Utah_Species; "
                    ColumnWidths ="0;2880"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =840
                    Top =60
                    Width =540
                    TabIndex =8
                    Name ="Transect_ID"
                    ControlSource ="Transect_ID"
                    StatusBarText ="Foreign key to tbl_Canopy_Transect"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =375
                    Left =4140
                    Top =60
                    Width =780
                    TabIndex =3
                    Name ="Alive"
                    ControlSource ="Alive"
                    RowSourceType ="Value List"
                    RowSource ="-1;\"Yes\";0;\"No\""
                    ColumnWidths ="0;375"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5100
                    Top =60
                    Width =1305
                    Height =300
                    TabIndex =4
                    Name ="ButtonLookup"
                    Caption ="Master Lookup"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6540
                    Top =60
                    Width =1545
                    Height =300
                    TabIndex =5
                    Name ="ButtonUnknown"
                    Caption ="Unknown Species"
                    OnClick ="[Event Procedure]"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =375
                    Left =10500
                    Top =60
                    Width =780
                    TabIndex =7
                    Name ="Surface_Alive"
                    ControlSource ="Surface_Alive"
                    RowSourceType ="Value List"
                    RowSource ="-1;\"Yes\";0;\"No\""
                    ColumnWidths ="0;375"
                End
            End
        End
        Begin FormFooter
            CanGrow = NotDefault
            Height =2280
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin Subform
                    OverlapFlags =85
                    Left =180
                    Top =120
                    Width =6465
                    Height =2040
                    Name ="fsub_LP_Lower_Canopy"
                    SourceObject ="Form.zfsub_LP_Lower_Canopy"
                    LinkChildFields ="Intercept_ID"
                    LinkMasterFields ="Intercept_ID"
                End
                Begin Subform
                    OverlapFlags =85
                    Left =6960
                    Top =120
                    Width =4005
                    Height =2040
                    TabIndex =1
                    Name ="fsub_LP_Disturbance"
                    SourceObject ="Form.zfsub_LP_Disturbance"
                    LinkChildFields ="Intercept_ID"
                    LinkMasterFields ="Intercept_ID"
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

Private Sub ButtonInitialize_Click()

    Dim db As DAO.Database
    Dim Points As DAO.Recordset
    Dim PointCount As Single
        
    On Error GoTo Err_Handler
    
    If Me!ButtonInitialize.ForeColor = 255 Then
      GoTo Exit_Procedure        ' Already initialized
    End If
    
    If IsNull(Me.Parent!Recorder) And IsNull(Me.Parent!Observer) Then
      MsgBox "You must enter Observer or Recorder first."
      GoTo Exit_Procedure
    End If
    
    ' Set point number
    Set db = CurrentDb
    Set Points = db.OpenRecordset("tbl_LP_Intercept")
    PointCount = 0.5
    Do Until PointCount > 50
      Points.AddNew
      Points!Intercept_ID = fxnGUIDGen  ' Generate an ID for it
      Points!Transect_ID = Forms!frm_Data_Entry!frm_LP_Transect.Form!Transect_ID
      Points!Point = PointCount
      Points!Alive = -1
      Points!Surface_Alive = 0
      Points.Update  ' write the record
      PointCount = PointCount + 0.5
    Loop

    Points.Close
    Me!ButtonInitialize.ForeColor = 255
    Me.Requery
    Me!fsub_LP_Lower_Canopy.Requery

Exit_Procedure:

    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

Private Sub Form_Current()
    Dim db As DAO.Database
    Dim Points As DAO.Recordset
    Dim strSQL As String
        
    On Error GoTo Err_Handler
    If IsNull(Me!Transect_ID) Then
      Me!ButtonInitialize.ForeColor = 8421376
      GoTo Exit_Procedure
    End If
    
    ' Set SQL
    Set db = CurrentDb
    strSQL = "SELECT [Point] FROM [tbl_LP_Intercept] WHERE [Transect_ID] = '" & Me![Transect_ID] & "'"
    Set Points = db.OpenRecordset(strSQL)
    
    If Points.EOF Then
      Me!ButtonInitialize.ForeColor = 8421376
    Else
      Me!ButtonInitialize.ForeColor = 255
      Me!fsub_LP_Lower_Canopy.Requery
      If IsNull(Me!Top) Then
        Me!Alive.Enabled = False
      Else
        Me!Alive.Enabled = True
      End If  ' End if for top canopy test
      If IsNull(Me!Surface) Or Me!Surface = "" Then
        Me!Surface_Alive.Enabled = False
      Else
          If IsNull(DLookup("[Surface_Code]", "tlu_LP_Soil_Surface", "[Surface_Code] = '" & Me!Surface & "'")) Then
            Me!Surface_Alive.Enabled = True
          Else
            Me!Surface_Alive.Enabled = False
          End If
      End If  ' End if for soil surface test
    End If  ' End if for points eof test
    Points.Close

Exit_Procedure:

    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

Private Sub Surface_AfterUpdate()
      If IsNull(Me!Surface) Or Me!Surface = "" Then
        Me!Surface_Alive = 0
        Me!Surface_Alive.Enabled = False
      Else
          If IsNull(DLookup("[Surface_Code]", "tlu_LP_Soil_Surface", "[Surface_Code] = '" & Me!Surface & "'")) Then
            Me!Surface_Alive.Enabled = True
            Me!Surface_Alive = -1
          Else
            Me!Surface_Alive = 0
            Me!Surface_Alive.Enabled = False
          End If
      End If
End Sub

Private Sub Top_AfterUpdate()
      If IsNull(Me!Top) Or Me!Top = "" Then
        Me!Alive.Enabled = False
      Else
        Me!Alive.Enabled = True
        Me!Alive = vbYes
      End If
End Sub

Private Sub Top_BeforeUpdate(Cancel As Integer)
    
    On Error GoTo Err_Handler
    
    ' Check for duplicate species in Soil Surface field.
    If Not IsNull(DLookup("[LC_ID]", "tbl_LP_Lower_Canopy", "[Intercept_ID] = '" & Me!Intercept_ID & "' AND [Species] = '" & Me!Top & "'")) Then
      MsgBox "This species is already recorded for this point."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
 
End Sub
Private Sub ButtonLookup_Click()
On Error GoTo Err_Button_Master_Species_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim strOpenArg As String

    strOpenArg = "fsub_LP_Intercept"
    stDocName = "frm_Master_Species"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , , strOpenArg

Exit_Button_Master_Species_Click:
    Exit Sub

Err_Button_Master_Species_Click:
    MsgBox Err.Description
    Resume Exit_Button_Master_Species_Click
     
End Sub
Private Sub ButtonUnknown_Click()

On Error GoTo Err_ButtonUnknown_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_LP_Unknown_Species"
    
    stLinkCriteria = "[Species_ID]=" & "'" & Me![Intercept_ID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , , Me![Intercept_ID]

Exit_ButtonUnknown_Click:
    Exit Sub

Err_ButtonUnknown_Click:
    MsgBox Err.Description
    Resume Exit_ButtonUnknown_Click
    
End Sub
