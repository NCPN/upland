Version =21
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    ShortcutMenu = NotDefault
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
    Width =13980
    DatasheetFontHeight =9
    ItemSuffix =107
    Right =11880
    Bottom =8370
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
            Height =720
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =60
                    Top =480
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
                    Left =1080
                    Top =480
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
                    Left =2520
                    Top =480
                    Width =480
                    Height =240
                    FontWeight =700
                    Name ="Alive_Label"
                    Caption ="A/D"
                    Tag ="DetachedLabel"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2220
                    Top =60
                    Width =1500
                    Height =300
                    ForeColor =2012742
                    Name ="ButtonInitialize"
                    Caption ="Initialize Form"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =2220
                    LayoutCachedTop =60
                    LayoutCachedWidth =3720
                    LayoutCachedHeight =360
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4080
                    Top =60
                    Width =1500
                    Height =300
                    TabIndex =1
                    Name ="ButtonLookup"
                    Caption ="Master Lookup"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5940
                    Top =60
                    Width =1500
                    Height =300
                    TabIndex =2
                    Name ="ButtonUnknown"
                    Caption ="Unknown Species"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =3120
                    Top =480
                    Width =1920
                    Height =240
                    FontWeight =700
                    Name ="Label99"
                    Caption ="Lower Canopy One"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5760
                    Top =480
                    Width =1919
                    Height =240
                    FontWeight =700
                    Name ="Label100"
                    Caption ="Lower Canopy Two"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =8400
                    Top =480
                    Width =1920
                    Height =240
                    FontWeight =700
                    Name ="Label101"
                    Caption ="Lower Canopy Three"
                    Tag ="DetachedLabel"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7800
                    Top =60
                    Width =1500
                    Height =300
                    TabIndex =3
                    Name ="ButtonLC"
                    Caption ="More LC"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9660
                    Top =60
                    Width =1500
                    Height =300
                    TabIndex =4
                    Name ="ButtonDisturbance"
                    Caption ="More Disturbance"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =12780
                    Top =480
                    Width =659
                    Height =240
                    FontWeight =700
                    Name ="Label89"
                    Caption ="Dist"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =11040
                    Top =480
                    Width =1140
                    Height =240
                    FontWeight =700
                    Name ="Surface_Label"
                    Caption ="Soil Surface"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5100
                    Top =480
                    Width =480
                    Height =240
                    FontWeight =700
                    Name ="Label103"
                    Caption ="A/D"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =7800
                    Top =480
                    Width =480
                    Height =240
                    FontWeight =700
                    Name ="Label104"
                    Caption ="A/D"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =10440
                    Top =480
                    Width =480
                    Height =240
                    FontWeight =700
                    Name ="Label105"
                    Caption ="A/D"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    Left =12180
                    Top =480
                    Width =480
                    Height =240
                    FontWeight =700
                    Name ="Label106"
                    Caption ="A/D"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =420
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
                    Top =60
                    Width =666
                    Height =255
                    ColumnWidth =2310
                    FontSize =6
                    FontWeight =700
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
                    ListWidth =6480
                    Left =11040
                    Top =60
                    Width =1080
                    TabIndex =11
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="Surface"
                    ControlSource ="Surface"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qryU_LP_Soil_Surface.Master_Plant_Code, qryU_LP_Soil_Surface.LU_Code, qry"
                        "U_LP_Soil_Surface.Utah_Species FROM qryU_LP_Soil_Surface WHERE (((qryU_LP_Soil_S"
                        "urface.Utah_Species) Is Not Null)); "
                    ColumnWidths ="0;2160;4320"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin ComboBox
                    OverlapFlags =255
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =6480
                    Left =660
                    Top =60
                    Width =1799
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="Top"
                    ControlSource ="Top"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qryU_Top_Canopy.Master_PLANT_Code, qryU_Top_Canopy.LU_Code, qryU_Top_Cano"
                        "py.Utah_Species FROM qryU_Top_Canopy WHERE (((qryU_Top_Canopy.Utah_Species) Is N"
                        "ot Null)) ORDER BY qryU_Top_Canopy.LU_Code;"
                    ColumnWidths ="0;2160;4320"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =840
                    Top =60
                    Width =540
                    TabIndex =4
                    Name ="Transect_ID"
                    ControlSource ="Transect_ID"
                    StatusBarText ="Foreign key to tbl_Canopy_Transect"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =375
                    Left =2520
                    Top =60
                    Width =479
                    TabIndex =3
                    Name ="Alive"
                    ControlSource ="Alive"
                    RowSourceType ="Value List"
                    RowSource ="-1;\"Y\";0;\"N\""
                    ColumnWidths ="0;375"
                    BeforeUpdate ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =375
                    Left =12180
                    Top =60
                    Width =478
                    TabIndex =12
                    Name ="Surface_Alive"
                    ControlSource ="Surface_Alive"
                    RowSourceType ="Value List"
                    RowSource ="-1;\"Y\";0;\"N\""
                    ColumnWidths ="0;375"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =6480
                    Left =3120
                    Top =60
                    Width =1980
                    TabIndex =5
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="LCS1"
                    ControlSource ="LCS1"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qryU_LP_Canopy.Master_Plant_Code, qryU_LP_Canopy.LU_Code, qryU_LP_Canopy."
                        "Utah_Species FROM qryU_LP_Canopy WHERE (((qryU_LP_Canopy.Utah_Species) Is Not Nu"
                        "ll)); "
                    ColumnWidths ="0;2160;4320"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =375
                    Left =5160
                    Top =60
                    Width =480
                    TabIndex =6
                    Name ="LCA1"
                    ControlSource ="LCA1"
                    RowSourceType ="Value List"
                    RowSource ="-1;\"Y\";0;\"N\""
                    ColumnWidths ="0;375"
                    BeforeUpdate ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =6480
                    Left =5760
                    Top =60
                    Width =1980
                    TabIndex =7
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="LCS2"
                    ControlSource ="LCS2"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qryU_LP_Canopy.Master_Plant_Code, qryU_LP_Canopy.LU_Code, qryU_LP_Canopy."
                        "Utah_Species FROM qryU_LP_Canopy WHERE (((qryU_LP_Canopy.Utah_Species) Is Not Nu"
                        "ll)); "
                    ColumnWidths ="0;2160;4320"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =375
                    Left =7800
                    Top =60
                    Width =480
                    TabIndex =8
                    Name ="LCA2"
                    ControlSource ="LCA2"
                    RowSourceType ="Value List"
                    RowSource ="-1;\"Y\";0;\"N\""
                    ColumnWidths ="0;375"
                    BeforeUpdate ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =6480
                    Left =8400
                    Top =60
                    Width =1980
                    TabIndex =9
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="LCS3"
                    ControlSource ="LCS3"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qryU_LP_Canopy.Master_Plant_Code, qryU_LP_Canopy.LU_Code, qryU_LP_Canopy."
                        "Utah_Species FROM qryU_LP_Canopy WHERE (((qryU_LP_Canopy.Utah_Species) Is Not Nu"
                        "ll)); "
                    ColumnWidths ="0;2160;4320"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =375
                    Left =10440
                    Top =60
                    Width =480
                    TabIndex =10
                    Name ="LCA3"
                    ControlSource ="LCA3"
                    RowSourceType ="Value List"
                    RowSource ="-1;\"Y\";0;\"N\""
                    ColumnWidths ="0;375"
                    BeforeUpdate ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2745
                    Left =12780
                    Top =60
                    Width =660
                    TabIndex =13
                    ColumnInfo ="\"Disturbance code\";\"\";\"disturbance description\";\"\";\"10\";\"10\""
                    Name ="D1"
                    ControlSource ="D1"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_LP_Disturbance.Dist_Code, tlu_LP_Disturbance.Disturbance FROM tlu_LP_"
                        "Disturbance; "
                    ColumnWidths ="495;2250"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"

                End
            End
        End
        Begin FormFooter
            CanGrow = NotDefault
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
Option Explicit

' =================================
' MODULE:       Form_fsub_LP_Intercept
' Level:        Form module
' Version:      1.02
' Description:  data functions & procedures specific to LP intercept monitoring
'
' Source/date:  Bonnie Campbell, 2/09/2016
' Revisions:    RDB - unknown  - 1.00 - initial version
'               BLC - 2/9/2016 - 1.01 - added documentation, checkbox for no species found
'               BLC - 8/17/2017 - 1.02 - switched from long to constant colors for readability
'                                        Son initialize fore color
' =================================

' ---------------------------------
'  Declarations
' ---------------------------------
Public CurrentPointID As String

' ---------------------------------
' SUB:          Form_Current
' Description:  Handles form current actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Russ DenBleyker, unknown
' Adapted:      Bonnie Campbell, February 9, 2016 - for NCPN tools
' Revisions:
'   RDB, unknown   - initial version
'   BLC, 2/9/2016  - added error handling, updated documentation
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim Points As DAO.Recordset
    Dim strSQL As String
        
    On Error GoTo Err_Handler
    If IsNull(Me!Transect_ID) Then
      Me!ButtonInitialize.ForeColor = lngDkBrtGrn '8421376
      GoTo Exit_Handler
    End If
    CurrentPointID = Me!Transect_ID
    ' Set SQL
    Set db = CurrentDb
    strSQL = "SELECT [Point] FROM [tbl_LP_Intercept] WHERE [Transect_ID] = '" & Me![Transect_ID] & "'"
    Set Points = db.OpenRecordset(strSQL)
    
    If Points.EOF Or IsNull(Points!Point) Then
      Me!ButtonInitialize.ForeColor = lngDkBrtGrn '8421376
    Else
      Me!ButtonInitialize.ForeColor = lngRed '255
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
    If IsNull(Me!LCS1) Or Me!LCS1 = "" Then
      Me!LCA1.Enabled = False
    Else
      Me!LCA1.Enabled = True
      Select Case Me!LCS1  ' If it's surface crud, its dead
        Case "L", "SL", "SW", "WD"
          Me!LCA1 = 0
          Me!LCA1.Enabled = False
      End Select
    End If
    If IsNull(Me!LCS2) Or Me!LCS2 = "" Then
      Me!LCA2.Enabled = False
    Else
      Me!LCA2.Enabled = True
      Select Case Me!LCS2  ' If it's surface crud, its dead
        Case "L", "SL", "SW", "WD"
          Me!LCA2 = 0
          Me!LCA2.Enabled = False
      End Select
    End If
    If IsNull(Me!LCS3) Or Me!LCS3 = "" Then
      Me!LCA3.Enabled = False
    Else
      Me!LCA3.Enabled = True
      Select Case Me!LCS3  ' If it's surface crud, its dead
        Case "L", "SL", "SW", "WD"
          Me!LCA3 = 0
          Me!LCA3.Enabled = False
      End Select
    End If
    
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[Form_fsub_LP_Intercept])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub ButtonInitialize_Click()

    Dim db As DAO.Database
    Dim Points As DAO.Recordset
    Dim PointCount As Single
    Dim PointIncrement As Single
    Dim PointLimit As Integer
    Dim Veg_Type As Variant
        
    On Error GoTo Err_Handler
    
    If Me!ButtonInitialize.ForeColor = 255 Then
      GoTo Exit_Procedure        ' Already initialized
    End If
    
    ' Disabled 3/19/2009 as per ecologist demand - RD
    ' If IsNull(Me.Parent!Recorder) And IsNull(Me.Parent!Observer) Then
    '   MsgBox "You must enter Observer or Recorder first."
    '   GoTo Exit_Procedure
    ' End If
    
    If IsNull(Me.Parent!Visit_Date) Then    ' If they didn't bother to enter a date, default to event date.
      Me.Parent!Visit_Date = Me.Parent.Parent!Start_Date
      Me.Parent.Refresh   ' Force save of transect record
    End If
    
    ' Set point number
    Set db = CurrentDb
    Set Points = db.OpenRecordset("tbl_LP_Intercept")
    Veg_Type = DLookup("[Vegetation_Type]", "tbl_Locations", "[Location_ID] = '" & Me.Parent.Parent!Location_ID & "'")
    If Not IsNull(Veg_Type) And Veg_Type = "oak scrub" Then
      PointCount = 0.25
      PointIncrement = 0.25
      PointLimit = 20
    Else
      PointCount = 0.5
      PointIncrement = 0.5
      PointLimit = 50
    End If
    Do Until PointCount > PointLimit
      Points.AddNew
      Points!Intercept_ID = fxnGUIDGen  ' Generate an ID for it
      Points!Transect_ID = Forms!frm_Data_Entry!frm_LP_Transect.Form!Transect_ID
      Points!Point = PointCount
      Points!Alive = -1
      Points!Surface_Alive = 0
      Points.Update  ' write the record
      PointCount = PointCount + PointIncrement
    Loop

    Points.Close
    Me!ButtonInitialize.ForeColor = 255
    Me.Requery

Exit_Procedure:

    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

'---------------
' Top/LCS (species)
'---------------
' ---------------------------------
' SUB:          Top_GotFocus
' Description:  Handles top species actions when control has focus
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Russ DenBleyker, unknown
' Adapted:      Bonnie Campbell, February 9, 2016 - for NCPN tools
' Revisions:
'   RDB, unknown  - initial version
'   BLC, 2/9/2016 - added error handling, documentation, refresh list to catch unknowns
' ---------------------------------
Private Sub Top_GotFocus()
On Error GoTo Err_Handler

    'update the data to ensure new unknowns are added
    Me.ActiveControl.Requery
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Top_GotFocus[Form_fsub_LP_Intercept])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub Top_BeforeUpdate(Cancel As Integer)
    Dim LCIndex As Integer
    Dim SpeciesColumn As String
    Dim AliveColumn As String
    Dim AliveValue As Boolean
    
    On Error GoTo Err_Handler
     
    LCIndex = 1
    SpeciesColumn = "LCS" & LCIndex
    Do Until IsNull(Me(SpeciesColumn))    ' Check for duplicate species in Lower Canopy.
      If Me(SpeciesColumn) = Me!Top Then
        If Me!Alive.Enabled = False Then
          AliveValue = vbYes  ' Top is going to default to alive if this is a new entry
        Else
          AliveValue = Me!Alive
        End If
        AliveColumn = "LCA" & LCIndex
        If Me(AliveColumn) = AliveValue Then
          MsgBox "This species is already recorded for this point."
          DoCmd.CancelEvent
          SendKeys "{ESC}"
          GoTo Exit_Procedure
        End If
      End If
      LCIndex = LCIndex + 1
      If LCIndex > 10 Then  ' Do not go past the end
        GoTo Exit_Procedure
      End If
      SpeciesColumn = "LCS" & LCIndex
    Loop
Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
 
End Sub

Private Sub Top_AfterUpdate()
      If IsNull(Me!Top) Or Me!Top = "" Then
        Me!Alive.Enabled = False
      Else
        Me!Alive.Enabled = True
        Me!Alive = vbYes
      End If
End Sub

' ---------------------------------
' SUB:          LCS1_GotFocus
' Description:  Handles lower canopy 1 species actions when control has focus
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Russ DenBleyker, unknown
' Adapted:      Bonnie Campbell, February 9, 2016 - for NCPN tools
' Revisions:
'   RDB, unknown  - initial version
'   BLC, 2/9/2016 - added error handling, documentation, refresh list to catch unknowns
' ---------------------------------
Private Sub LCS1_GotFocus()
On Error GoTo Err_Handler

    'update the data to ensure new unknowns are added
    Me.ActiveControl.Requery
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - LCS1_GotFocus[Form_fsub_LP_Intercept])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub LCS1_BeforeUpdate(Cancel As Integer)
  Dim Reply As Integer
  Dim AorD As Boolean
  Dim TextMsg As String
  If Not IsNull(Me!LCS1) Then
   Me!LCA1.Enabled = True
   Select Case Me!LCS1
     Case "L", "SL", "SW", "WD"
       Me!LCA1 = 0
       Me!LCA1.Enabled = False
    '   Me.Refresh
   End Select
   
   AorD = Me!LCA1
   If TestDuplicateSpecies([LCS1], 1, AorD) Then
     Select Case Me!LCS1
       Case "L", "SL", "SW", "WD"
         MsgBox "This surface is already recorded for this point."
         DoCmd.CancelEvent
         SendKeys "{ESC}"
         GoTo Exit_Sub
     End Select
     ' The code below was commented to bypass the message requesting user input. [HT, 3-24-15]
     ' -- Begin commented code [HT, 3-24-15]
     ' If AorD Then
     '   TextMsg = "This species already exists as alive on this point.  Would you like to set it to dead?"
     ' Else
     '   TextMsg = "This species already exists as dead on this point.  Would you like to set it to alive?"
     ' End If
     ' Reply = MsgBox(TextMsg, vbYesNo, "Species Verification")
     ' If Reply = vbYes Then
     ' -- End commented code [HT, 3-24-15]
       AorD = IIf(AorD = True, False, True)
       If TestDuplicateSpecies([LCS1], 1, AorD) Then
         MsgBox "This species is already recorded for this point."
         DoCmd.CancelEvent
         SendKeys "{ESC}"
         GoTo Exit_Sub
       End If
     ' -- Begin commented code [HT, 3-24-15]
     ' Else
'       MsgBox "This species is already recorded for this point."
      ' DoCmd.CancelEvent
      ' SendKeys "{ESC}"
      ' GoTo Exit_Sub
     ' End If  ' End if for reply test
     ' -- End commented code [HT, 3-24-15]
   End If  '  End if for duplicate species test
   Me!LCA1 = AorD  ' Make sure alive or dead field is correct
  Else
    Me!LCA1.Enabled = False
  End If   ' End if for null field test
Exit_Sub:
End Sub

Private Sub LCS1_AfterUpdate()
  Dim ResultFlag As Boolean
  Dim lngPosition As Long
  
  lngPosition = Me.CurrentRecord ' capture index position of record currently selected
  If lngPosition > 1 Then
    lngPosition = lngPosition - 1
  End If
  If IsNull(Me!LCS1) Then
    ResultFlag = ClearLCGaps(1)
    Me!LCS1.SetFocus   ' Reset focus
    Me.Form.Recordset.Move lngPosition ' navigate back to original record position
  End If
  
End Sub

' ---------------------------------
' SUB:          LCS2_GotFocus
' Description:  Handles lower canopy 2 species actions when control has focus
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Russ DenBleyker, unknown
' Adapted:      Bonnie Campbell, February 9, 2016 - for NCPN tools
' Revisions:
'   RDB, unknown  - initial version
'   BLC, 2/9/2016 - added error handling, documentation, refresh list to catch unknowns
' ---------------------------------
Private Sub LCS2_GotFocus()
On Error GoTo Err_Handler

    'update the data to ensure new unknowns are added
    Me.ActiveControl.Requery
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - LCS2_GotFocus[Form_fsub_LP_Intercept])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub LCS2_BeforeUpdate(Cancel As Integer)
  Dim Reply As Integer
  Dim GapColumn As Integer
  Dim AorD As Boolean
  Dim TextMsg As String
  
  If Not IsNull(Me!LCS2) Then
   Me!LCA2.Enabled = True
   GapColumn = TestGaps(2)
   If GapColumn > 0 Then  ' First check to see if they're making gaps
     MsgBox "You cannot create gaps in LC.  LC" & GapColumn & " is available."
     DoCmd.CancelEvent
     SendKeys "{ESC}"
     GoTo Exit_Sub
   End If
   Select Case Me!LCS2  ' If it's surface crud, its dead
     Case "L", "SL", "SW", "WD"
       Me!LCA2 = 0
       Me!LCA2.Enabled = False
   End Select
   AorD = Me!LCA2  ' Now check for duplicate species
   If TestDuplicateSpecies([LCS2], 2, AorD) Then
     Select Case Me!LCS2
       Case "L", "SL", "SW", "WD"
         MsgBox "This surface is already recorded for this point."
         DoCmd.CancelEvent
         SendKeys "{ESC}"
         GoTo Exit_Sub
     End Select
     ' The code below was commented to bypass the message requesting user input. [HT, 3-24-15]
     ' -- Begin commented code [HT, 3-24-15]
     ' If AorD Then
     '   TextMsg = "This species already exists as alive on this point.  Would you like to set it to dead?"
     ' Else
     '  TextMsg = "This species already exists as dead on this point.  Would you like to set it to alive?"
     ' End If
     ' Reply = MsgBox(TextMsg, vbYesNo, "Species Verification")
     ' If Reply = vbYes Then
     ' -- End commented code [HT, 3-24-15]
       AorD = IIf(AorD = True, False, True)
       If TestDuplicateSpecies([LCS2], 2, AorD) Then
         MsgBox "This species is already recorded for this point."
         DoCmd.CancelEvent
         SendKeys "{ESC}"
         GoTo Exit_Sub
       End If
     ' -- Begin commented code [HT, 3-24-15]
     ' Else
'       MsgBox "This species is already recorded for this point."
     '  DoCmd.CancelEvent
     '  SendKeys "{ESC}"
     '  GoTo Exit_Sub
     ' End If
     ' -- End commented code [HT, 3-24-15]
   End If
   Me!LCA2 = AorD  ' Make sure alive or dead field is correct
  Else
   Me!LCA2.Enabled = False
  End If
Exit_Sub:
End Sub

Private Sub LCS2_AfterUpdate()
  Dim ResultFlag As Boolean
  Dim lngPosition As Long
  
  lngPosition = Me.CurrentRecord ' capture index position of record currently selected
  If lngPosition > 1 Then
    lngPosition = lngPosition - 1
  End If
  If IsNull(Me!LCS2) Then
    ResultFlag = ClearLCGaps(2)
    Me!LCS2.SetFocus   ' Reset focus
    Me.Form.Recordset.Move lngPosition ' navigate back to original record position
  End If
  
End Sub

' ---------------------------------
' SUB:          LCS3_GotFocus
' Description:  Handles lower canopy 3 species actions when control has focus
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Russ DenBleyker, unknown
' Adapted:      Bonnie Campbell, February 9, 2016 - for NCPN tools
' Revisions:
'   RDB, unknown  - initial version
'   BLC, 2/9/2016 - added error handling, documentation, refresh list to catch unknowns
' ---------------------------------
Private Sub LCS3_GotFocus()
On Error GoTo Err_Handler

    'update the data to ensure new unknowns are added
    Me.ActiveControl.Requery
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - LCS3_GotFocus[Form_fsub_LP_Intercept])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub LCS3_BeforeUpdate(Cancel As Integer)
  Dim Reply As Integer
  Dim GapColumn As Integer
  Dim AorD As Boolean
  Dim TextMsg As String
  
  If Not IsNull(Me!LCS3) Then
   Me!LCA3.Enabled = True
   GapColumn = TestGaps(3)
   If GapColumn > 0 Then  ' First check to see if they're making gaps
     MsgBox "You cannot create gaps in LC.  LC" & GapColumn & " is available."
     DoCmd.CancelEvent
     SendKeys "{ESC}"
     GoTo Exit_Sub
   End If
   Select Case Me!LCS3  ' If it's surface crud, its dead
     Case "L", "SL", "SW", "WD"
       Me!LCA3 = 0
       Me!LCA3.Enabled = False
   End Select
   AorD = Me!LCA3
   If TestDuplicateSpecies([LCS3], 3, AorD) Then
     Select Case Me!LCS3
       Case "L", "SL", "SW", "WD"
         MsgBox "This surface is already recorded for this point."
         DoCmd.CancelEvent
         SendKeys "{ESC}"
         GoTo Exit_Sub
     End Select
     ' The code below was commented to bypass the message requesting user input. [HT, 3-24-15]
     ' -- Begin commented code [HT, 3-24-15]
     ' If AorD Then
     '   TextMsg = "This species already exists as alive on this point.  Would you like to set it to dead?"
     ' Else
     '   TextMsg = "This species already exists as dead on this point.  Would you like to set it to alive?"
     ' End If
     ' Reply = MsgBox(TextMsg, vbYesNo, "Species Verification")
     ' If Reply = vbYes Then
     ' -- End commented code [HT, 3-24-15]
       AorD = IIf(AorD = True, False, True)
       If TestDuplicateSpecies([LCS3], 3, AorD) Then
         MsgBox "This species is already recorded for this point."
         DoCmd.CancelEvent
         SendKeys "{ESC}"
         GoTo Exit_Sub
       End If
     ' -- Begin commented code [HT, 3-24-15]
     ' Else
'       MsgBox "This species is already recorded for this point."
     '   DoCmd.CancelEvent
     '   SendKeys "{ESC}"
     '   GoTo Exit_Sub
     ' End If
     ' -- End commented code [HT, 3-24-15]
   End If
   Me!LCA3 = AorD  ' Make sure alive or dead field is correct
  Else
   Me!LCA3.Enabled = False
  End If
Exit_Sub:
End Sub

Private Sub LCS3_AfterUpdate()
  Dim ResultFlag As Boolean
  Dim lngPosition As Long
  
  lngPosition = Me.CurrentRecord ' capture index position of record currently selected
  If lngPosition > 1 Then
    lngPosition = lngPosition - 1
  End If
  If IsNull(Me!LCS3) Then
    ResultFlag = ClearLCGaps(3)
    Me!LCS3.SetFocus   ' Reset focus
    Me.Form.Recordset.Move lngPosition ' navigate back to original record position
  End If
  
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

'---------------
' LCA (Alive or Dead)
'---------------

Private Sub LCA1_BeforeUpdate(Cancel As Integer)
  Dim AorD As Boolean
  If IsNull(Me!LCS1) Then
     MsgBox "Species cannot be null."
     DoCmd.CancelEvent
     SendKeys "{ESC}"
     GoTo Exit_Sub
   End If
   AorD = Me!LCA1
   If TestDuplicateSpecies([LCS1], 1, AorD) Then
     MsgBox "This species is already recorded for this point."
     DoCmd.CancelEvent
     SendKeys "{ESC}"
   End If
Exit_Sub:
End Sub

Private Sub LCA2_BeforeUpdate(Cancel As Integer)
  Dim AorD As Boolean
  If IsNull(Me!LCS2) Then
     MsgBox "Species cannot be null."
     DoCmd.CancelEvent
     SendKeys "{ESC}"
     GoTo Exit_Sub
   End If
   AorD = Me!LCA2
   If TestDuplicateSpecies([LCS2], 2, AorD) Then
     MsgBox "This species is already recorded for this point."
     DoCmd.CancelEvent
     SendKeys "{ESC}"
   End If
Exit_Sub:
End Sub

Private Sub LCA3_BeforeUpdate(Cancel As Integer)
  Dim AorD As Boolean
  If IsNull(Me!LCS3) Then
     MsgBox "Species cannot be null."
     DoCmd.CancelEvent
     SendKeys "{ESC}"
     GoTo Exit_Sub
   End If
   AorD = Me!LCA3
   If TestDuplicateSpecies([LCS3], 3, AorD) Then
     MsgBox "This species is already recorded for this point."
     DoCmd.CancelEvent
     SendKeys "{ESC}"
   End If
Exit_Sub:
End Sub

Private Sub Alive_BeforeUpdate(Cancel As Integer)
    Dim LCIndex As Integer
    Dim SpeciesColumn As String
    Dim AliveColumn As String
    
    On Error GoTo Err_Handler
     
    LCIndex = 1
    SpeciesColumn = "LCS" & LCIndex
    Do Until IsNull(Me(SpeciesColumn))    ' Check for duplicate species in Lower Canopy.
      If Me(SpeciesColumn) = Me!Top Then
        AliveColumn = "LCA" & LCIndex
        If Me(AliveColumn) = Me!Alive Then
          MsgBox "This species is already recorded for this point."
          DoCmd.CancelEvent
          SendKeys "{ESC}"
          GoTo Exit_Procedure
        End If
      End If
      LCIndex = LCIndex + 1
      If LCIndex > 10 Then  ' Do not go past the end
        GoTo Exit_Procedure
      End If
      SpeciesColumn = "LCS" & LCIndex
    Loop
Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
 
End Sub

Private Sub D1_BeforeUpdate(Cancel As Integer)
  
  If Not IsNull(Me!D1) Then
    If TestDuplicateDist([D1], 1) Then
      MsgBox "This disturbance is already recorded for this point."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
    End If
  End If
  
End Sub

Private Sub D1_AfterUpdate()
    Dim GapIndex As Integer
    Dim NextIndex As Integer
    Dim SpeciesColumn As String
    Dim NextColumn As String
    
    On Error GoTo Err_Handler
  If IsNull(Me!D1) Then   ' If they cleared it, we need to eliminate any gaps.
    GapIndex = 1
    NextIndex = 2
    Do Until GapIndex > 4
      NextColumn = "D" & NextIndex
      If IsNull(Me(NextColumn)) Then    ' Check for disturbance in next entry.
        GoTo Exit_Procedure   ' Nope - we are finished
      Else
        SpeciesColumn = "D" & GapIndex
        Me(SpeciesColumn) = Me(NextColumn)   ' move the next column down.
        Me(NextColumn) = Null                ' clear the old column
      End If
      GapIndex = GapIndex + 1
      NextIndex = NextIndex + 1
    Loop
  End If
Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (ClearDisturbanceGaps)"
            Resume Exit_Procedure
    End Select

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

    DoCmd.OpenForm "frm_List_Unknown", , , , , acDialog
    Me.Refresh
    
Exit_ButtonUnknown_Click:
    Exit Sub

Err_ButtonUnknown_Click:
    MsgBox Err.Description
    Resume Exit_ButtonUnknown_Click
    
End Sub

Private Sub ButtonLC_Click()
On Error GoTo Err_ButtonLC_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
   
    If IsNull(Me!Intercept_ID) Then
      MsgBox "No point has been slected.", , "More Lower Canopy"
    ElseIf (IsNull(Me!LCS1) + IsNull(Me!LCS2) + IsNull(Me!LCS3)) <> 0 Then
      MsgBox "There is space available in LC1, LC2, or LC3.", , "More Lower Canopy"
    Else
      DoCmd.RunCommand acCmdSaveRecord  ' Force a save so frm_More_LC can see any recent updates.
      stDocName = "frm_More_LC"
      stLinkCriteria = "[Intercept_ID]=" & "'" & Me![Intercept_ID] & "'"
      DoCmd.OpenForm stDocName, , , stLinkCriteria
    End If
Exit_ButtonLC_Click:
    Exit Sub

Err_ButtonLC_Click:
    MsgBox Err.Description
    Resume Exit_ButtonLC_Click
    
End Sub

Private Sub ButtonDisturbance_Click()
On Error GoTo Err_ButtonDisturbance_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    If IsNull(Me!Intercept_ID) Then
      MsgBox "No point has been slected.", , "More Disturbance"
    ElseIf IsNull(Me!D1) Then
      MsgBox "D1 is available.", , "More Disturbance"
    Else
      DoCmd.RunCommand acCmdSaveRecord  ' Force a save so frm_More_Disturbance can see D1 if it was just updated.
      stDocName = "frm_More_Disturbance"
      stLinkCriteria = "[Intercept_ID]=" & "'" & Me![Intercept_ID] & "'"
      DoCmd.OpenForm stDocName, , , stLinkCriteria
    End If

Exit_ButtonDisturbance_Click:
    Exit Sub

Err_ButtonDisturbance_Click:
    MsgBox Err.Description
    Resume Exit_ButtonDisturbance_Click
    
End Sub


'---------------
' Functions
'---------------

Public Function ClearLCGaps(SpeciesIndex As Integer) As Boolean
' Clear gaps in lower canopy - 2/27/2009 - Russ DenBleyker
' Northern Colorado Plateau Network
' Called from lower canopy updates to clear gaps caused by nulling of an LC column
' SpeciesIndex = Index of the calling field
' Returns true if operation was successful

    Dim GapIndex As Integer
    Dim NextIndex As Integer
    Dim SpeciesColumn As String
    Dim NextColumn As String
    Dim AliveColumn As String
    
    On Error GoTo Err_Handler
    ClearLCGaps = True   ' Assume AOK
    GapIndex = SpeciesIndex
    NextIndex = GapIndex + 1
    Do Until GapIndex > 9
      NextColumn = "LCS" & NextIndex
      If IsNull(Me(NextColumn)) Then    ' Check for species in next entry.
        GoTo Exit_Procedure_CG   ' Nope - we are finished
      Else
        SpeciesColumn = "LCS" & GapIndex
        Me(SpeciesColumn) = Me(NextColumn)   ' move the next column down.
        Me(NextColumn) = Null                ' clear the old column
        SpeciesColumn = "LCA" & GapIndex
        NextColumn = "LCA" & NextIndex
        Me(SpeciesColumn) = Me(NextColumn)   ' get the a/d flag.
        Me(NextColumn) = False            ' set old column a/d to default
      End If
      GapIndex = GapIndex + 1
      NextIndex = NextIndex + 1
    Loop
    
Exit_Procedure_CG:
    Me.Requery     ' Necessary to force frm_More_LC to reflect this update.
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (TestGaps)"
                ClearLCGaps = False
            Resume Exit_Procedure_CG
    End Select

End Function

Public Function TestGaps(SpeciesIndex As Integer) As Integer
' Test for gaps in lower canopy - 2/27/2009 - Russ DenBleyker
' Northern Colorado Plateau Network
' Called from lower canopy updates to check for gaps in entries
' SpeciesIndex = Index of the calling field
' Returns zero if no gaps or the number of an available field

    Dim GapIndex As Integer
    Dim SpeciesColumn As String
    
    On Error GoTo Err_Handler
    TestGaps = 0  ' Assume it is not a duplicate
    GapIndex = SpeciesIndex
    Do Until GapIndex < 2
      GapIndex = GapIndex - 1
      SpeciesColumn = "LCS" & GapIndex
      If IsNull(Me(SpeciesColumn)) Then    ' Check for duplicate species in Lower Canopy.
        TestGaps = GapIndex  ' Flag available column
        GoTo Exit_Procedure_TG
      End If
    Loop
    
Exit_Procedure_TG:
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (TestGaps)"
            Resume Exit_Procedure_TG
    End Select

End Function

Public Function TestDuplicateSpecies(Species As String, SpeciesIndex As Integer, AnimationState As Boolean) As Boolean
' Test for duplicate species in a point - 2/26/2009 - Russ DenBleyker
' Northern Colorado Plateau Network
' Called from lower canopy updates to check for duplication of species
' Species = Species code to test
' SpeciesIndex = Index of the calling field
' Animation State = Alive (-1) or Dead (0)
' Returns true if species exists and animation state is equal

    Dim LCIndex As Integer
    Dim SpeciesColumn As String
    Dim AliveColumn As String
    
    On Error GoTo Err_Handler
    TestDuplicateSpecies = False  ' Assume it is not a duplicate
    LCIndex = 1
    SpeciesColumn = "LCS" & LCIndex
    Do Until IsNull(Me(SpeciesColumn))    ' Check for duplicate species in Lower Canopy.
      If LCIndex <> SpeciesIndex Then     ' Do not test calling field
        If Me(SpeciesColumn) = Species Then
          AliveColumn = "LCA" & LCIndex
          If Me(AliveColumn) = AnimationState Then
            TestDuplicateSpecies = True
            GoTo Exit_Procedure_TDS
          End If
        End If
      End If
      LCIndex = LCIndex + 1
      If LCIndex > 10 Then  ' Do not go past the end
        GoTo Exit_Procedure_TDS
      End If
      SpeciesColumn = "LCS" & LCIndex
    Loop
    If Me!Top = Species And Me!Alive = AnimationState Then  ' Test top canopy
      TestDuplicateSpecies = True
    End If

Exit_Procedure_TDS:
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (TestDuplicatespecies)"
            Resume Exit_Procedure_TDS
    End Select

End Function

Public Function TestDuplicateDist(Disturbance As String, DistIndex As Integer) As Boolean
' Test for duplicate disturbance in a point - 3/18/2010 - Russ DenBleyker
' Northern Colorado Plateau Network
' Called from disturbance updates to check for duplicates
' Disturbance = Disturbance code to test
' distIndex = Index of the calling field
' Returns true if disturbance exists

    Dim DIndex As Integer
    Dim DistColumn As String
    
    On Error GoTo Err_Handler
    TestDuplicateDist = False  ' Assume it is not a duplicate
    DIndex = 1
    DistColumn = "D" & DIndex
    Do Until IsNull(Me(DistColumn))    ' Check for duplicate disturbances.
      If DIndex <> DistIndex Then     ' Do not test calling field
        If Me(DistColumn) = Disturbance Then
          TestDuplicateDist = True
          GoTo Exit_Procedure_TDD
        End If
      End If
      DIndex = DIndex + 1
      If DIndex > 5 Then  ' Do not go past the end
        GoTo Exit_Procedure_TDD
      End If
      DistColumn = "D" & DIndex
    Loop
Exit_Procedure_TDD:
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (TestDuplicateDist)"
            Resume Exit_Procedure_TDD
    End Select

End Function
