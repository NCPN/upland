Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =127
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =6915
    DatasheetFontHeight =9
    ItemSuffix =59
    Left =5340
    Top =3735
    Right =10125
    Bottom =8625
    DatasheetGridlinesColor =12632256
    Filter ="[Intercept_ID]='{9EF6C165-5AFC-42C5-BB3F-B92C765F933B}'"
    RecSrcDt = Begin
        0xa80917b9b277e340
    End
    RecordSource ="tbl_LP_Intercept"
    Caption ="frm_More_LC"
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
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =4500
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Width =840
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
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4020
                    Top =540
                    Width =720
                    Height =255
                    ColumnWidth =2310
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="Point"
                    ControlSource ="Point"
                    Format ="General Number"
                    StatusBarText ="Intercept point - increments of .5m up to 50.0"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =2400
                            Top =540
                            Width =1560
                            Height =255
                            FontSize =10
                            FontWeight =700
                            Name ="Point_Label"
                            Caption ="Point Number"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1140
                    Top =120
                    Width =4800
                    Height =360
                    FontSize =12
                    FontWeight =700
                    Name ="Label32"
                    Caption ="Add Lower Canopy Levels 4-10"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =375
                    Left =3780
                    Top =1320
                    Width =480
                    TabIndex =3
                    Name ="LCA4"
                    ControlSource ="LCA4"
                    RowSourceType ="Value List"
                    RowSource ="-1;\"Y\";0;\"N\""
                    ColumnWidths ="0;375"
                    BeforeUpdate ="[Event Procedure]"

                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3120
                    Top =4020
                    Width =1020
                    Height =300
                    TabIndex =16
                    Name ="ButtonClose"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =6480
                    Left =720
                    Top =1320
                    Width =2880
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="LCS4"
                    ControlSource ="LCS4"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qryU_LP_Canopy.Master_Plant_Code, qryU_LP_Canopy.LU_Code, qryU_LP_Canopy."
                        "Utah_Species FROM qryU_LP_Canopy WHERE (((qryU_LP_Canopy.Utah_Species) Is Not Nu"
                        "ll)); "
                    ColumnWidths ="0;2160;4320"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =180
                    Top =1320
                    Width =420
                    Height =240
                    Name ="Label36"
                    Caption ="LC 4"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1680
                    Top =1020
                    Width =1200
                    Height =240
                    FontWeight =700
                    Name ="Label37"
                    Caption ="Species"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =3660
                    Top =1020
                    Width =780
                    Height =240
                    FontWeight =700
                    Name ="Label38"
                    Caption ="Alive?"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =375
                    Left =3780
                    Top =1680
                    Width =480
                    TabIndex =5
                    Name ="LCA5"
                    ControlSource ="LCA5"
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
                    Left =720
                    Top =1680
                    Width =2880
                    TabIndex =4
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="LCS5"
                    ControlSource ="LCS5"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qryU_LP_Canopy.Master_Plant_Code, qryU_LP_Canopy.LU_Code, qryU_LP_Canopy."
                        "Utah_Species FROM qryU_LP_Canopy WHERE (((qryU_LP_Canopy.Utah_Species) Is Not Nu"
                        "ll)); "
                    ColumnWidths ="0;2160;4320"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =180
                    Top =1680
                    Width =420
                    Height =240
                    Name ="Label41"
                    Caption ="LC 5"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =375
                    Left =3780
                    Top =2040
                    Width =480
                    TabIndex =7
                    Name ="LCA6"
                    ControlSource ="LCA6"
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
                    Left =720
                    Top =2040
                    Width =2880
                    TabIndex =6
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="LCS6"
                    ControlSource ="LCS6"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qryU_LP_Canopy.Master_Plant_Code, qryU_LP_Canopy.LU_Code, qryU_LP_Canopy."
                        "Utah_Species FROM qryU_LP_Canopy WHERE (((qryU_LP_Canopy.Utah_Species) Is Not Nu"
                        "ll)); "
                    ColumnWidths ="0;2160;4320"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =180
                    Top =2040
                    Width =420
                    Height =240
                    Name ="Label44"
                    Caption ="LC 6"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =375
                    Left =3780
                    Top =2400
                    Width =480
                    TabIndex =9
                    Name ="LCA7"
                    ControlSource ="LCA7"
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
                    Left =720
                    Top =2400
                    Width =2880
                    TabIndex =8
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="LCS7"
                    ControlSource ="LCS7"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qryU_LP_Canopy.Master_Plant_Code, qryU_LP_Canopy.LU_Code, qryU_LP_Canopy."
                        "Utah_Species FROM qryU_LP_Canopy WHERE (((qryU_LP_Canopy.Utah_Species) Is Not Nu"
                        "ll)); "
                    ColumnWidths ="0;2160;4320"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =180
                    Top =2400
                    Width =420
                    Height =240
                    Name ="Label47"
                    Caption ="LC 7"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =375
                    Left =3780
                    Top =2760
                    Width =480
                    TabIndex =11
                    Name ="LCA8"
                    ControlSource ="LCA8"
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
                    Left =720
                    Top =2760
                    Width =2880
                    TabIndex =10
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="LCS8"
                    ControlSource ="LCS8"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qryU_LP_Canopy.Master_Plant_Code, qryU_LP_Canopy.LU_Code, qryU_LP_Canopy."
                        "Utah_Species FROM qryU_LP_Canopy WHERE (((qryU_LP_Canopy.Utah_Species) Is Not Nu"
                        "ll)); "
                    ColumnWidths ="0;2160;4320"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =180
                    Top =2760
                    Width =420
                    Height =240
                    Name ="Label50"
                    Caption ="LC 8"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =375
                    Left =3780
                    Top =3120
                    Width =480
                    TabIndex =13
                    Name ="LCA9"
                    ControlSource ="LCA9"
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
                    Left =720
                    Top =3120
                    Width =2880
                    TabIndex =12
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="LCS9"
                    ControlSource ="LCS9"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qryU_LP_Canopy.Master_Plant_Code, qryU_LP_Canopy.LU_Code, qryU_LP_Canopy."
                        "Utah_Species FROM qryU_LP_Canopy WHERE (((qryU_LP_Canopy.Utah_Species) Is Not Nu"
                        "ll)); "
                    ColumnWidths ="0;2160;4320"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =180
                    Top =3120
                    Width =420
                    Height =240
                    Name ="Label53"
                    Caption ="LC 9"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =375
                    Left =3780
                    Top =3480
                    Width =480
                    TabIndex =15
                    Name ="LCA10"
                    ControlSource ="LCA10"
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
                    Left =720
                    Top =3480
                    Width =2880
                    TabIndex =14
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="LCS10"
                    ControlSource ="LCS10"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qryU_LP_Canopy.Master_Plant_Code, qryU_LP_Canopy.LU_Code, qryU_LP_Canopy."
                        "Utah_Species FROM qryU_LP_Canopy WHERE (((qryU_LP_Canopy.Utah_Species) Is Not Nu"
                        "ll)); "
                    ColumnWidths ="0;2160;4320"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =120
                    Top =3480
                    Width =480
                    Height =240
                    Name ="Label56"
                    Caption ="LC 10"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4680
                    Top =1500
                    Width =1560
                    Height =300
                    TabIndex =17
                    Name ="ButtonMaster"
                    Caption ="Master Species"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4680
                    Top =1980
                    Width =1545
                    Height =300
                    TabIndex =18
                    Name ="ButtonUnknown"
                    Caption ="Unknown Species"
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
Public Function ClearMoreLCGaps(SpeciesIndex As Integer) As Boolean
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
    ClearMoreLCGaps = True   ' Assume AOK
    GapIndex = SpeciesIndex
    NextIndex = GapIndex + 1
    Do Until GapIndex > 9
      NextColumn = "LCS" & NextIndex
      If IsNull(Me(NextColumn)) Then    ' Check for species in next entry.
        GoTo Exit_Procedure_CMG   ' Nope - we are finished
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
    
Exit_Procedure_CMG:
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (ClearMoreGaps)"
                ClearMoreLCGaps = False
            Resume Exit_Procedure_CMG
    End Select

End Function
Public Function TestMoreGaps(SpeciesIndex As Integer) As Integer
' Test for gaps in lower canopy - 2/27/2009 - Russ DenBleyker
' Northern Colorado Plateau Network
' Called from lower canopy updates to check for gaps in entries
' SpeciesIndex = Index of the calling field
' Returns zero if no gaps or the number of an available field

    Dim GapIndex As Integer
    Dim SpeciesColumn As String
    
    On Error GoTo Err_Handler
    TestMoreGaps = 0  ' Assume no available gap
    GapIndex = SpeciesIndex
    Do Until GapIndex < 2
      GapIndex = GapIndex - 1
      SpeciesColumn = "LCS" & GapIndex
      If IsNull(Me(SpeciesColumn)) Then    ' Check for available gap in Lower Canopy.
        TestMoreGaps = GapIndex  ' Flag available column
        GoTo Exit_Procedure_TMG
      End If
    Loop
    
Exit_Procedure_TMG:
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (TestGaps)"
            Resume Exit_Procedure_TMG
    End Select

End Function

' ---------------------------------
' FUNCTION:     TestMoreDuplicateSpecies
' Description:  Test for duplicate species in a point
'               Called from lower canopy updates to check for duplication of species
' Assumptions:  -
' Parameters:   Species - species to check (string)
'               SpeciesIndex - index of the calling field (integer)
'               AnimationState - Alive (-1) or Dead (0) (boolean)
' Returns:      Returns true if species exists and animation state is equal, false if not
' Throws:       none
' References:   -
' Source/date:  Russ DenBleyker, February 26, 2009, Northern Colorado Plateau Network
' Adapted:      -
' Revisions:
'   RD  - 2/26/2009 - initial version
'   BLC - 5/27/2015 - fixed bug causing top canopy check (exiting procedure vs. loop),
'                     updated error handling & added documentation
' ---------------------------------
Public Function TestMoreDuplicateSpecies(Species As String, SpeciesIndex As Integer, AnimationState As Boolean) As Boolean
    Dim LCIndex As Integer
    Dim SpeciesColumn As String
    Dim AliveColumn As String
    
    On Error GoTo Err_Handler
    TestMoreDuplicateSpecies = False  ' Assume it is not a duplicate
    LCIndex = 1
    SpeciesColumn = "LCS" & LCIndex
    
    '-------------------------
    ' lower canopy duplicates?
    '-------------------------
    Do Until IsNull(Me(SpeciesColumn))    ' Check for duplicate species in Lower Canopy.
      If LCIndex <> SpeciesIndex Then     ' Do not test calling field
        If Me(SpeciesColumn) = Species Then
          AliveColumn = "LCA" & LCIndex
          If Me(AliveColumn) = AnimationState Then
            TestMoreDuplicateSpecies = True
            GoTo Exit_Function 'Exit_Procedure_TMDS
          End If
        End If
      End If
      LCIndex = LCIndex + 1
      If LCIndex > 10 Then  ' Do not go past the end
        'GoTo Exit_Procedure_TMDS
        'exit loop vs. procedure to avoid missing comparison to top canopy species
        Exit Do
      End If
      SpeciesColumn = "LCS" & LCIndex
    Loop
    
    '-------------------------
    ' top canopy duplicates?
    '-------------------------
    If Me!Top = Species And Me!Alive = AnimationState Then  ' Test top canopy
      TestMoreDuplicateSpecies = True
    End If

'Exit_Procedure_TMDS:
Exit_Function:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - TestMoreDuplicatespecies[frm_More_LC])"
    End Select
    Resume Exit_Function
End Function

Private Sub ButtonClose_Click()
On Error GoTo Err_ButtonClose_Click


    DoCmd.Close

Exit_ButtonClose_Click:
    Exit Sub

Err_ButtonClose_Click:
    MsgBox Err.Description
    Resume Exit_ButtonClose_Click
    
End Sub
Private Sub ButtonMaster_Click()
On Error GoTo Err_ButtonMaster_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    strOpenArg = "fsub_LP_Intercept"
    stDocName = "frm_Master_Species"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , , strOpenArg

Exit_ButtonMaster_Click:
    Exit Sub

Err_ButtonMaster_Click:
    MsgBox Err.Description
    Resume Exit_ButtonMaster_Click
    
End Sub
Private Sub ButtonUnknown_Click()
On Error GoTo Err_ButtonUnknown_Click

    DoCmd.OpenForm "frm_List_Unknown", , , , , acDialog
    Me.Refresh

Exit_ButtonUnknown_Click:
    Exit Sub

Err_ButtonUnknown_Click:
    MsgBox Err.Description
    Resume Exit_ButtonUnknown_Click
    
End Sub

Private Sub LCA10_BeforeUpdate(Cancel As Integer)
  Dim AorD As Boolean
  If IsNull(Me!LCS10) Then
     MsgBox "Species cannot be null."
     DoCmd.CancelEvent
     SendKeys "{ESC}"
     GoTo Exit_Sub
   End If
   AorD = Me!LCA10
   If TestMoreDuplicateSpecies([LCS10], 10, AorD) Then
     MsgBox "This species is already recorded for this point."
     DoCmd.CancelEvent
     SendKeys "{ESC}"
   End If
Exit_Sub:
End Sub

Private Sub LCA4_BeforeUpdate(Cancel As Integer)
  Dim AorD As Boolean
  If IsNull(Me!LCS4) Then
     MsgBox "Species cannot be null."
     DoCmd.CancelEvent
     SendKeys "{ESC}"
     GoTo Exit_Sub
   End If
   AorD = Me!LCA4
   If TestMoreDuplicateSpecies([LCS4], 4, AorD) Then
     MsgBox "This species is already recorded for this point."
     DoCmd.CancelEvent
     SendKeys "{ESC}"
   End If
Exit_Sub:
End Sub

Private Sub LCA5_BeforeUpdate(Cancel As Integer)
  Dim AorD As Boolean
  If IsNull(Me!LCS5) Then
     MsgBox "Species cannot be null."
     DoCmd.CancelEvent
     SendKeys "{ESC}"
     GoTo Exit_Sub
   End If
   AorD = Me!LCA5
   If TestMoreDuplicateSpecies([LCS5], 5, AorD) Then
     MsgBox "This species is already recorded for this point."
     DoCmd.CancelEvent
     SendKeys "{ESC}"
   End If
Exit_Sub:
End Sub

Private Sub LCA6_BeforeUpdate(Cancel As Integer)
  Dim AorD As Boolean
  If IsNull(Me!LCS6) Then
     MsgBox "Species cannot be null."
     DoCmd.CancelEvent
     SendKeys "{ESC}"
     GoTo Exit_Sub
   End If
   AorD = Me!LCA6
   If TestMoreDuplicateSpecies([LCS6], 6, AorD) Then
     MsgBox "This species is already recorded for this point."
     DoCmd.CancelEvent
     SendKeys "{ESC}"
   End If
Exit_Sub:
End Sub

Private Sub LCA7_BeforeUpdate(Cancel As Integer)
  Dim AorD As Boolean
  If IsNull(Me!LCS7) Then
     MsgBox "Species cannot be null."
     DoCmd.CancelEvent
     SendKeys "{ESC}"
     GoTo Exit_Sub
   End If
   AorD = Me!LCA7
   If TestMoreDuplicateSpecies([LCS7], 7, AorD) Then
     MsgBox "This species is already recorded for this point."
     DoCmd.CancelEvent
     SendKeys "{ESC}"
   End If
Exit_Sub:
End Sub

Private Sub LCA8_BeforeUpdate(Cancel As Integer)
  Dim AorD As Boolean
  If IsNull(Me!LCS8) Then
     MsgBox "Species cannot be null."
     DoCmd.CancelEvent
     SendKeys "{ESC}"
     GoTo Exit_Sub
   End If
   AorD = Me!LCA8
   If TestMoreDuplicateSpecies([LCS8], 8, AorD) Then
     MsgBox "This species is already recorded for this point."
     DoCmd.CancelEvent
     SendKeys "{ESC}"
   End If
Exit_Sub:
End Sub

Private Sub LCA9_BeforeUpdate(Cancel As Integer)
  Dim AorD As Boolean
  If IsNull(Me!LCS9) Then
     MsgBox "Species cannot be null."
     DoCmd.CancelEvent
     SendKeys "{ESC}"
     GoTo Exit_Sub
   End If
   AorD = Me!LCA9
   If TestMoreDuplicateSpecies([LCS9], 9, AorD) Then
     MsgBox "This species is already recorded for this point."
     DoCmd.CancelEvent
     SendKeys "{ESC}"
   End If
Exit_Sub:
End Sub

Private Sub LCS10_AfterUpdate()
  Dim ResultFlag As Boolean
  
  If IsNull(Me!LCS10) Then
    Me!LCA10 = 0  ' Set default A/D flag.
  End If
  
End Sub

Private Sub LCS10_BeforeUpdate(Cancel As Integer)
  Dim Reply As Integer
  Dim AorD As Boolean
  Dim TextMsg As String
  Dim GapColumn As Integer
  
   GapColumn = TestMoreGaps(10)
   If GapColumn > 0 Then  ' First check to see if they're making gaps
     MsgBox "You cannot create gaps in LC.  LC" & GapColumn & " is available."
     DoCmd.CancelEvent
     SendKeys "{ESC}"
     GoTo Exit_Sub
   End If
   Select Case Me!LCS10  ' If it's surface crud, its dead
     Case "L", "SL", "SW", "WD"
       Me!LCA10 = 0
   End Select
  If Not IsNull(Me!LCS10) Then
   AorD = Me!LCA10
   If TestMoreDuplicateSpecies([LCS10], 10, AorD) Then
     Select Case Me!LCS10
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
       If TestMoreDuplicateSpecies([LCS10], 10, AorD) Then
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
   Me!LCA10 = AorD  ' Make sure alive or dead field is correct
  End If
Exit_Sub:
End Sub

Private Sub LCS4_AfterUpdate()
  Dim ResultFlag As Boolean
  
  If IsNull(Me!LCS4) Then
    ResultFlag = ClearMoreLCGaps(4)    '  eliminate the gap if they deleted the entry
  End If
  
End Sub

Private Sub LCS4_BeforeUpdate(Cancel As Integer)
  Dim Reply As Integer
  Dim AorD As Boolean
  Dim TextMsg As String
  Dim GapColumn As String
  
   GapColumn = TestMoreGaps(4)
   If GapColumn > 0 Then  ' First check to see if they're making gaps
     MsgBox "You cannot create gaps in LC.  LC" & GapColumn & " is available."
     DoCmd.CancelEvent
     SendKeys "{ESC}"
     Exit Sub
   End If
   Select Case Me!LCS4  ' If it's surface crud, its dead
     Case "L", "SL", "SW", "WD"
       Me!LCA4 = 0
   End Select
  If Not IsNull(Me!LCS4) Then
   AorD = Me!LCA4
   If TestMoreDuplicateSpecies([LCS4], 4, AorD) Then
     Select Case Me!LCS4
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
       If TestMoreDuplicateSpecies([LCS4], 4, AorD) Then
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
   Me!LCA4 = AorD  ' Make sure alive or dead field is correct
  End If
Exit_Sub:
End Sub

Private Sub LCS5_AfterUpdate()
  Dim ResultFlag As Boolean
  
  If IsNull(Me!LCS5) Then
    ResultFlag = ClearMoreLCGaps(5)    '  eliminate the gap if they deleted the entry
  End If
  
End Sub

Private Sub LCS5_BeforeUpdate(Cancel As Integer)
  Dim Reply As Integer
  Dim AorD As Boolean
  Dim TextMsg As String
  Dim GapColumn As Integer
  
   GapColumn = TestMoreGaps(5)
   If GapColumn > 0 Then  ' First check to see if they're making gaps
     MsgBox "You cannot create gaps in LC.  LC" & GapColumn & " is available."
     DoCmd.CancelEvent
     SendKeys "{ESC}"
     GoTo Exit_Sub
   End If
   Select Case Me!LCS5  ' If it's surface crud, its dead
     Case "L", "SL", "SW", "WD"
       Me!LCA5 = 0
   End Select
  If Not IsNull(Me!LCS5) Then
   AorD = Me!LCA5
   If TestMoreDuplicateSpecies([LCS5], 5, AorD) Then
     Select Case Me!LCS5
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
       If TestMoreDuplicateSpecies([LCS5], 5, AorD) Then
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
   Me!LCA5 = AorD  ' Make sure alive or dead field is correct
  End If
Exit_Sub:
End Sub

Private Sub LCS6_AfterUpdate()
  Dim ResultFlag As Boolean
  
  If IsNull(Me!LCS6) Then
    ResultFlag = ClearMoreLCGaps(6)    '  eliminate the gap if they deleted the entry
  End If
  
End Sub

Private Sub LCS6_BeforeUpdate(Cancel As Integer)
  Dim Reply As Integer
  Dim AorD As Boolean
  Dim TextMsg As String
  Dim GapColumn As Integer
  
   GapColumn = TestMoreGaps(6)
   If GapColumn > 0 Then  ' First check to see if they're making gaps
     MsgBox "You cannot create gaps in LC.  LC" & GapColumn & " is available."
     DoCmd.CancelEvent
     SendKeys "{ESC}"
     GoTo Exit_Sub
   End If
   Select Case Me!LCS6  ' If it's surface crud, its dead
     Case "L", "SL", "SW", "WD"
       Me!LCA6 = 0
   End Select
  If Not IsNull(Me!LCS6) Then
   AorD = Me!LCA6
   If TestMoreDuplicateSpecies([LCS6], 6, AorD) Then
     Select Case Me!LCS6
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
       If TestMoreDuplicateSpecies([LCS6], 6, AorD) Then
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
   Me!LCA6 = AorD  ' Make sure alive or dead field is correct
  End If
Exit_Sub:
End Sub

Private Sub LCS7_AfterUpdate()
  Dim ResultFlag As Boolean
  
  If IsNull(Me!LCS7) Then
    ResultFlag = ClearMoreLCGaps(7)    '  eliminate the gap if they deleted the entry
  End If
  
End Sub

Private Sub LCS7_BeforeUpdate(Cancel As Integer)
  Dim Reply As Integer
  Dim AorD As Boolean
  Dim TextMsg As String
  Dim GapColumn As Integer
  
   GapColumn = TestMoreGaps(7)
   If GapColumn > 0 Then  ' First check to see if they're making gaps
     MsgBox "You cannot create gaps in LC.  LC" & GapColumn & " is available."
     DoCmd.CancelEvent
     SendKeys "{ESC}"
     GoTo Exit_Sub
   End If
   Select Case Me!LCS7  ' If it's surface crud, its dead
     Case "L", "SL", "SW", "WD"
       Me!LCA7 = 0
   End Select
  If Not IsNull(Me!LCS7) Then
   AorD = Me!LCA7
   If TestMoreDuplicateSpecies([LCS7], 7, AorD) Then
     Select Case Me!LCS7
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
       If TestMoreDuplicateSpecies([LCS7], 7, AorD) Then
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
   Me!LCA7 = AorD  ' Make sure alive or dead field is correct
  End If
Exit_Sub:
End Sub

Private Sub LCS8_AfterUpdate()
  Dim ResultFlag As Boolean
  
  If IsNull(Me!LCS8) Then
    ResultFlag = ClearMoreLCGaps(8)    '  eliminate the gap if they deleted the entry
  End If
  
End Sub

Private Sub LCS8_BeforeUpdate(Cancel As Integer)
  Dim Reply As Integer
  Dim AorD As Boolean
  Dim TextMsg As String
  Dim GapColumn As Integer
  
   GapColumn = TestMoreGaps(8)
   If GapColumn > 0 Then  ' First check to see if they're making gaps
     MsgBox "You cannot create gaps in LC.  LC" & GapColumn & " is available."
     DoCmd.CancelEvent
     SendKeys "{ESC}"
     GoTo Exit_Sub
   End If
   Select Case Me!LCS8  ' If it's surface crud, its dead
     Case "L", "SL", "SW", "WD"
       Me!LCA8 = 0
   End Select
  If Not IsNull(Me!LCS8) Then
   AorD = Me!LCA8
   If TestMoreDuplicateSpecies([LCS8], 8, AorD) Then
     Select Case Me!LCS8
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
       If TestMoreDuplicateSpecies([LCS8], 8, AorD) Then
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
   Me!LCA8 = AorD  ' Make sure alive or dead field is correct
  End If
Exit_Sub:
End Sub

Private Sub LCS9_AfterUpdate()
  Dim ResultFlag As Boolean
  
  If IsNull(Me!LCS9) Then
    ResultFlag = ClearMoreLCGaps(9)    '  eliminate the gap if they deleted the entry
  End If
  
End Sub

Private Sub LCS9_BeforeUpdate(Cancel As Integer)
  Dim Reply As Integer
  Dim AorD As Boolean
  Dim TextMsg As String
  Dim GapColumn As Integer
  
   GapColumn = TestMoreGaps(9)
   If GapColumn > 0 Then  ' First check to see if they're making gaps
     MsgBox "You cannot create gaps in LC.  LC" & GapColumn & " is available."
     DoCmd.CancelEvent
     SendKeys "{ESC}"
     GoTo Exit_Sub
   End If
   Select Case Me!LCS9  ' If it's surface crud, its dead
     Case "L", "SL", "SW", "WD"
       Me!LCA9 = 0
   End Select
  If Not IsNull(Me!LCS9) Then
   AorD = Me!LCA9
   If TestMoreDuplicateSpecies([LCS9], 9, AorD) Then
     Select Case Me!LCS9
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
       If TestMoreDuplicateSpecies([LCS9], 9, AorD) Then
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
   Me!LCA9 = AorD  ' Make sure alive or dead field is correct
  End If
Exit_Sub:
End Sub
