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
    ItemSuffix =61
    Left =2760
    Top =2415
    Right =7545
    Bottom =6000
    DatasheetGridlinesColor =12632256
    Filter ="[Intercept_ID]='{20317BF4-825E-49D7-BE79-60E75A5A86B2}'"
    RecSrcDt = Begin
        0xa80917b9b277e340
    End
    RecordSource ="tbl_LP_Intercept"
    Caption ="frm_More_LC"
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
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =3600
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
                    Left =1125
                    Top =120
                    Width =4815
                    Height =360
                    FontSize =12
                    FontWeight =700
                    Name ="Label32"
                    Caption ="Add Disturbance 2-5"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3120
                    Top =2880
                    Width =1020
                    Height =300
                    TabIndex =6
                    Name ="ButtonClose"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =180
                    Top =1320
                    Width =420
                    Height =240
                    Name ="Label36"
                    Caption ="D 2"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =180
                    Top =1680
                    Width =420
                    Height =240
                    Name ="Label41"
                    Caption ="D 3"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =180
                    Top =2040
                    Width =420
                    Height =240
                    Name ="Label44"
                    Caption ="D 4"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =180
                    Top =2400
                    Width =420
                    Height =240
                    Name ="Label47"
                    Caption ="D 5"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2745
                    Left =840
                    Top =1320
                    Width =2340
                    TabIndex =2
                    ColumnInfo ="\"Disturbance code\";\"\";\"disturbance description\";\"\";\"10\";\"10\""
                    Name ="D2"
                    ControlSource ="D2"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_LP_Disturbance.Dist_Code, tlu_LP_Disturbance.Disturbance FROM tlu_LP_"
                        "Disturbance; "
                    ColumnWidths ="495;2250"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2745
                    Left =840
                    Top =1680
                    Width =2340
                    TabIndex =3
                    ColumnInfo ="\"Disturbance code\";\"\";\"disturbance description\";\"\";\"10\";\"10\""
                    Name ="D3"
                    ControlSource ="D3"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_LP_Disturbance.Dist_Code, tlu_LP_Disturbance.Disturbance FROM tlu_LP_"
                        "Disturbance; "
                    ColumnWidths ="495;2250"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2745
                    Left =840
                    Top =2040
                    Width =2340
                    TabIndex =4
                    ColumnInfo ="\"Disturbance code\";\"\";\"disturbance description\";\"\";\"10\";\"10\""
                    Name ="D4"
                    ControlSource ="D4"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_LP_Disturbance.Dist_Code, tlu_LP_Disturbance.Disturbance FROM tlu_LP_"
                        "Disturbance; "
                    ColumnWidths ="495;2250"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2745
                    Left =840
                    Top =2400
                    Width =2340
                    TabIndex =5
                    ColumnInfo ="\"Disturbance code\";\"\";\"disturbance description\";\"\";\"10\";\"10\""
                    Name ="D5"
                    ControlSource ="D5"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_LP_Disturbance.Dist_Code, tlu_LP_Disturbance.Disturbance FROM tlu_LP_"
                        "Disturbance; "
                    ColumnWidths ="495;2250"
                    BeforeUpdate ="[Event Procedure]"
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
Public Function ClearDisturbanceGaps(DisturbanceIndex As Integer) As Boolean
' Clear gaps in lower canopy - 3/4/2009 - Russ DenBleyker
' Northern Colorado Plateau Network
' Called from disturbance updates to clear gaps caused by nulling of a column
' DisturbanceIndex = Index of the calling field
' Returns true if operation was successful

    Dim GapIndex As Integer
    Dim NextIndex As Integer
    Dim SpeciesColumn As String
    Dim NextColumn As String
    
    On Error GoTo Err_Handler
    ClearDisturbanceGaps = True   ' Assume AOK
    GapIndex = DisturbanceIndex
    NextIndex = GapIndex + 1
    Do Until GapIndex > 4
      NextColumn = "D" & NextIndex
      If IsNull(Me(NextColumn)) Then    ' Check for disturbance in next entry.
        GoTo Exit_Procedure_CDG   ' Nope - we are finished
      Else
        SpeciesColumn = "D" & GapIndex
        Me(SpeciesColumn) = Me(NextColumn)   ' move the next column down.
        Me(NextColumn) = Null                ' clear the old column
      End If
      GapIndex = GapIndex + 1
      NextIndex = NextIndex + 1
    Loop
    
Exit_Procedure_CDG:
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (ClearDisturbanceGaps)"
                ClearDisturbanceGaps = False
            Resume Exit_Procedure_CDG
    End Select

End Function
Public Function TestDisturbanceGaps(DistIndex As Integer) As Integer
' Test for gaps in disturbances - 3/3/2009 - Russ DenBleyker
' Northern Colorado Plateau Network
' Called from disturbance update to check for gaps in entries
' GapIndex = Index of the calling field
' Returns zero if no gaps or the number of an available field

    Dim GapIndex As Integer
    Dim DistColumn As String
    
    On Error GoTo Err_Handler
    TestDisturbanceGaps = 0  ' Assume it is not a duplicate
    GapIndex = DistIndex
    Do Until GapIndex < 2
      GapIndex = GapIndex - 1
      DistColumn = "D" & GapIndex
      If IsNull(Me(DistColumn)) Then    ' Check for available spot.
        TestDisturbanceGaps = GapIndex  ' Flag available column
        GoTo Exit_Procedure_TD
      End If
    Loop
    
Exit_Procedure_TD:
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (TestDisturbanceGaps)"
            Resume Exit_Procedure_TD
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
Private Sub ButtonClose_Click()
On Error GoTo Err_ButtonClose_Click

    DoCmd.Close

Exit_ButtonClose_Click:
    Exit Sub

Err_ButtonClose_Click:
    MsgBox Err.Description
    Resume Exit_ButtonClose_Click
    
End Sub

Private Sub D2_AfterUpdate()
  Dim ResultFlag As Boolean
  
  If IsNull(Me!D2) Then
    ResultFlag = ClearDisturbanceGaps(2)    '  eliminate the gap if they deleted the entry
  End If
End Sub

Private Sub D2_BeforeUpdate(Cancel As Integer)
  If Not IsNull(Me!D2) Then
    If TestDuplicateDist([D2], 2) Then
      MsgBox "This disturbance is already recorded for this point."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
    End If
  End If
End Sub

Private Sub D3_AfterUpdate()
  Dim ResultFlag As Boolean
  
  If IsNull(Me!D3) Then
    ResultFlag = ClearDisturbanceGaps(3)    '  eliminate the gap if they deleted the entry
  End If
End Sub

Private Sub D3_BeforeUpdate(Cancel As Integer)
  Dim GapColumn As Integer
  
  GapColumn = TestDisturbanceGaps(3)
  If GapColumn > 0 Then
     MsgBox "You cannot create gaps in Disturbances.  D" & GapColumn & " is available."
     DoCmd.CancelEvent
     SendKeys "{ESC}"
  End If
  If Not IsNull(Me!D3) Then
    If TestDuplicateDist([D3], 3) Then
      MsgBox "This disturbance is already recorded for this point."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
    End If
  End If
End Sub

Private Sub D4_AfterUpdate()
  Dim ResultFlag As Boolean
  
  If IsNull(Me!D4) Then
    ResultFlag = ClearDisturbanceGaps(4)    '  eliminate the gap if they deleted the entry
  End If
End Sub

Private Sub D4_BeforeUpdate(Cancel As Integer)
  Dim GapColumn As Integer
  
  GapColumn = TestDisturbanceGaps(4)
  If GapColumn > 0 Then
     MsgBox "You cannot create gaps in Disturbances.  D" & GapColumn & " is available."
     DoCmd.CancelEvent
     SendKeys "{ESC}"
  End If
  If Not IsNull(Me!D4) Then
    If TestDuplicateDist([D4], 4) Then
      MsgBox "This disturbance is already recorded for this point."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
    End If
  End If
End Sub

Private Sub D5_BeforeUpdate(Cancel As Integer)
  Dim GapColumn As Integer
  
  GapColumn = TestDisturbanceGaps(5)
  If GapColumn > 0 Then
     MsgBox "You cannot create gaps in Disturbances.  D" & GapColumn & " is available."
     DoCmd.CancelEvent
     SendKeys "{ESC}"
  End If
  If Not IsNull(Me!D5) Then
    If TestDuplicateDist([D5], 5) Then
      MsgBox "This disturbance is already recorded for this point."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
    End If
  End If
End Sub
