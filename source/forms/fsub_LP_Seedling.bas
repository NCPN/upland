Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5760
    DatasheetFontHeight =9
    ItemSuffix =28
    Top =5736
    Right =4968
    Bottom =8292
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x73776174d27ce340
    End
    RecordSource ="tbl_LP_Seedling"
    Caption ="fsub_LP_Seedling"
    BeforeInsert ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    FetchDefaults =0
    FetchDefaults =0
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
                    Left =240
                    Top =420
                    Width =1335
                    Height =240
                    FontWeight =700
                    Name ="Species_Label"
                    Caption ="Species Name"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =2775
                    Top =420
                    Width =540
                    Height =240
                    FontWeight =700
                    Name ="Total_Label"
                    Caption ="Total"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =480
                    Top =60
                    Width =3045
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label23"
                    Caption ="Tree Seedlings"
                End
            End
        End
        Begin Section
            Height =420
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
                    Name ="Shrub_ID"
                    ControlSource ="Seedling_ID"
                    StatusBarText ="Unique record identifier - primary key"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =300
                    Top =60
                    Width =300
                    Height =255
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Transect_ID"
                    ControlSource ="Transect_ID"
                    StatusBarText ="Foreign key to tbl_Canopy_Transect"

                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =2760
                    Top =60
                    Width =540
                    Height =255
                    ColumnWidth =600
                    TabIndex =3
                    BackColor =65535
                    Name ="SeedTotal"
                    ControlSource ="Total"
                    StatusBarText ="0-10cm height class total"
                    DefaultValue ="Null"
                    OnGotFocus ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000ac000000030000000000000006000000000000000200000001000000 ,
                        0x00000000ffffff00000000000500000003000000050000000100000000000000 ,
                        0xffff0000010000000000000006000000250000000100000000000000ffffff00 ,
                        0x3000000000003000000000005b0050006100720065006e0074005d002e005b00 ,
                        0x6300620078004e006f0053006500650064006c0069006e00670073005d003d00 ,
                        0x540072007500650000000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000300000000000000060000000100000000000000ffffff00010000003000 ,
                        0x0000000000000000000000000000000000000000000000000005000000010000 ,
                        0x0000000000ffff00000100000030000000000000000000000000000000000000 ,
                        0x0000000001000000000000000100000000000000ffffff001e0000005b005000 ,
                        0x6100720065006e0074005d002e005b006300620078004e006f00530065006500 ,
                        0x64006c0069006e00670073005d003d0054007200750065000000000000000000 ,
                        0x00000000000000000000000000
                    End
                End
                Begin ComboBox
                    OverlapFlags =247
                    TextFontFamily =0
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =6480
                    Left =180
                    Top =60
                    Width =2304
                    TabIndex =2
                    BackColor =65535
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    ConditionalFormat = Begin
                        0x01000000ec000000030000000100000000000000000000001100000001000000 ,
                        0x00000000ffffff00010000000000000012000000250000000100000000000000 ,
                        0xffff0000010000000000000026000000450000000100000000000000ffffff00 ,
                        0x4c0065006e0028005b0053007000650063006900650073005d0029003e003000 ,
                        0x00000000490073004e0075006d00650072006900630028005b0054006f007400 ,
                        0x61006c005d002900000000005b0050006100720065006e0074005d002e005b00 ,
                        0x6300620078004e006f0053006500650064006c0069006e00670073005d003d00 ,
                        0x540072007500650000000000
                    End
                    Name ="Species"
                    ControlSource ="Species"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT qryU_Top_Canopy.Master_PLANT_Code, qryU_Top_Canopy.LU_Code, qryU"
                        "_Top_Canopy.Utah_Species, qryU_Top_Canopy.Lifeform FROM qryU_Top_Canopy WHERE (("
                        "(qryU_Top_Canopy.Utah_Species) Is Not Null) AND ((qryU_Top_Canopy.[Lifeform])='T"
                        "ree')) ORDER BY qryU_Top_Canopy.LU_Code  UNION  (SELECT DISTINCT tbl_Unknown_Spe"
                        "cies.Unknown_Code, tbl_Unknown_Species.Unknown_Code, tbl_Unknown_Species.Plant_T"
                        "ype + \" - \" + tbl_Unknown_Species.Plant_Description, tbl_Unknown_Species.Plant"
                        "_Type AS Lifeform FROM tbl_Unknown_Species WHERE tbl_Unknown_Species.Plant_Type "
                        "IN ('Tree','Other') OR tbl_Unknown_Species.Plant_Type IS NULL ORDER BY tbl_Unkno"
                        "wn_Species.Unknown_Code);"
                    ColumnWidths ="0;2160;4320"
                    BeforeUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    ConditionalFormat14 = Begin
                        0x01000300000001000000000000000100000000000000ffffff00100000004c00 ,
                        0x65006e0028005b0053007000650063006900650073005d0029003e0030000000 ,
                        0x0000000000000000000000000000000000000001000000000000000100000000 ,
                        0x000000ffff000012000000490073004e0075006d00650072006900630028005b ,
                        0x0054006f00740061006c005d0029000000000000000000000000000000000000 ,
                        0x0000000001000000000000000100000000000000ffffff001e0000005b005000 ,
                        0x6100720065006e0074005d002e005b006300620078004e006f00530065006500 ,
                        0x64006c0069006e00670073005d003d0054007200750065000000000000000000 ,
                        0x00000000000000000000000000
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =3600
                    Top =60
                    Width =705
                    Height =300
                    TabIndex =4
                    ForeColor =255
                    Name ="ButtonDelete"
                    Caption ="Delete"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
            End
        End
        Begin FormFooter
            Height =480
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =720
                    Top =60
                    Width =606
                    Height =288
                    Name ="TallyA1"
                    Caption ="+ 1"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =1440
                    Top =60
                    Width =606
                    Height =288
                    TabIndex =1
                    Name ="TallyA5"
                    Caption ="+ 5"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =2160
                    Top =60
                    Width =606
                    Height =288
                    TabIndex =2
                    Name ="TallyS1"
                    Caption ="- 1"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =2880
                    Top =60
                    Width =606
                    Height =288
                    TabIndex =3
                    Name ="TallyS5"
                    Caption ="- 5"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =3600
                    Top =60
                    Width =606
                    Height =288
                    TabIndex =4
                    Name ="TallyA0"
                    Caption ="0"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
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
Option Explicit

' =================================
' MODULE:       Form_fsub_LP_Seedling
' Level:        Form module
' Version:      1.02
' Description:  data functions & procedures specific to LP seedling monitoring
'
' Source/date:  Bonnie Campbell, 2/2/2016
' Revisions:    RDB - unknown  - 1.00 - initial version
'               BLC - 2/2/2016 - 1.01 - added documentation, checkbox for no species found
'               BLC - 4/1/2016 - 1.02 - added clearing of SeedTotal 0 (table default) to set to NULL
'                                       ensuring that field crew enter data (vs data being defaulted)
' =================================

' ---------------------------------
' SUB:          Form_Load
' Description:  Handles form loading actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 2, 2016 - for NCPN tools
' Revisions:
'   BLC, 2/2/2016  - initial version
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler

    'reset tally buttons (disable)
    DisableTallyButtons Me, "Tally"
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[Form_fsub_LP_Seedling])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_BeforeInsert
' Description:  Handles form pre-insert actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Russ DenBleyker, unknown
' Adapted:      Bonnie Campbell, February 11, 2016 - for NCPN tools
' Revisions:
'   RDB, unknown    - initial version
'   BLC, 2/11/2016  - added no data collected info updates
'   BLC, 4/1/2016   - added setting SeedTotal to NULL to clear 0 (table default)
' ---------------------------------
Private Sub Form_BeforeInsert(Cancel As Integer)
On Error GoTo Err_Handler
 
   ' If IsNull(Me.Parent!Observer) And IsNull(Me.Parent!Recorder) Then
   '   MsgBox "You must enter an observer or recorder first."
   '   DoCmd.CancelEvent
   '   SendKeys "{ESC}"
   '   GoTo Exit_Procedure
   ' End If
    ' Create the GUID primary key value
    If IsNull(Me!Seedling_ID) Then
        If GetDataType("tbl_LP_Seedling", "Seedling_ID") = dbText Then
            Me.Seedling_ID = fxnGUIDGen
        End If
    End If
    
    'clear seed total value to get rid of 0 table default
    Me.SeedTotal = Null
    
    '-----------------------------------
    ' update the NoDataCollected info
    '-----------------------------------
    Dim NoData As Scripting.Dictionary
    
    'remove the no data collected record
    Set NoData = SetNoDataCollected(Me.Parent.Form.Controls("Transect_ID"), "T", "1mBelt-TreeSeedling", 0)
        
    'update checkbox/rectangle
    Me.Parent.Form.Controls("cbxNoSeedlings") = 0
    Me.Parent.Form.Controls("cbxNoSeedlings").Enabled = False
    Me.Parent.Form.Controls("rctNoSeedlings").Visible = False

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeInsert[Form_fsub_Dist_Exotic])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxNoSpecies_Click
' Description:  Handles checkbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 2, 2016 - for NCPN tools
' Revisions:
'   BLC, 2/2/2016  - initial version
' ---------------------------------
Private Sub cbxNoSpecies_Click()
On Error GoTo Err_Handler


Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxNoSpecies_Click[Form_fsub_LP_Seedling])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Species_GotFocus
' Description:  Handles species actions when control has focus
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
'   BLC, 4/1/2016 - revised based on use of AddTallyValue & tally
'                   buttons disabling when tally values should not be available
' ---------------------------------
Private Sub Species_GotFocus()
On Error GoTo Err_Handler

    If IsNull(Me.Parent!Visit_Date) Then    ' If they didn't bother to enter a date, default to event date.
      Me.Parent!Visit_Date = Me.Parent.Parent!Start_Date
      Me.Parent.Refresh   ' Force save of transect record
    End If

    'update the data to ensure new unknowns are added
    Me.ActiveControl.Requery
    
    'reset tally buttons (disable)
    DisableTallyButtons Me, "Tally"
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Species_GotFocus[Form_fsub_LP_Seedling])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Species_Change
' Description:  Handles species actions when control is changed
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, March 29, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC, 3/29/2016  - initial version
' ---------------------------------
Private Sub Species_Change()
On Error GoTo Err_Handler

    'clear seed total value to get rid of 0
'    Me.SeedTotal = Null
'    Me.Refresh

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Species_Change[Form_fsub_LP_Seedling])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Species_BeforeUpdate
' Description:  Handles species actions when control is updated
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Russ DenBleyker, unknown
' Adapted:      Bonnie Campbell, February 9, 2016 - for NCPN tools
' Revisions:
'   RDB, unknown  - initial version
'   BLC, 2/9/2016 - added error handling, documentation
' ---------------------------------
Private Sub Species_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

    If Not IsNull(DLookup("[Seedling_ID]", "tbl_LP_Seedling", "[Transect_ID] = '" & Me!Transect_ID & "' AND [Species] = '" & Me!Species & "'")) Then
      MsgBox "This species is already recorded for this transect."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
      Me.Undo
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Species_BeforeUpdate[Form_fsub_LP_Seedling])"
    End Select
    Resume Exit_Handler
End Sub

'==================================
'   Tally Buttons
'==================================

' ---------------------------------
' SUB:          SeedTotal_Click
' Description:  Handles actions when control is clicked
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, April 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC, 4/1/2016  - initial version
' ---------------------------------
Private Sub SeedTotal_Click()
On Error GoTo Err_Handler
  
  'default none are enabled
  DisableTallyButtons Me, "Tally"
  
  'disable tallies that drive seed total < 0
  Select Case Nz(SeedTotal.Value, "Null")
    Case "Null"
        TallyA0.Enabled = True
        TallyA1.Enabled = True
        TallyA5.Enabled = True
    Case Is < 1
        TallyA0.Enabled = True
        TallyA1.Enabled = True
        TallyA5.Enabled = True
    Case Is < 5
        TallyS1.Enabled = True
        TallyA0.Enabled = True
        TallyA1.Enabled = True
        TallyA5.Enabled = True
    Case Is >= 5
        TallyS5.Enabled = True
        TallyS1.Enabled = True
        TallyA0.Enabled = True
        TallyA1.Enabled = True
        TallyA5.Enabled = True
  End Select
  
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SeedTotal_Click[Form_fsub_LP_Seedling])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          SeedTotal_GotFocus
' Description:  Handles actions when after control is the focus
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, April 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC, 4/1/2016  - initial version
' ---------------------------------
Private Sub SeedTotal_GotFocus()
On Error GoTo Err_Handler
  
  'default none are enabled
  DisableTallyButtons Me, "Tally"
  
  'disable tallies that drive seed total < 0
  Select Case Nz(SeedTotal.Value, "Null")
    Case "Null"
        TallyA0.Enabled = True
        TallyA1.Enabled = True
        TallyA5.Enabled = True
    Case Is < 1
        TallyA0.Enabled = True
        TallyA1.Enabled = True
        TallyA5.Enabled = True
    Case Is < 5
        TallyS1.Enabled = True
        TallyA0.Enabled = True
        TallyA1.Enabled = True
        TallyA5.Enabled = True
    Case Is >= 5
        TallyS5.Enabled = True
        TallyS1.Enabled = True
        TallyA0.Enabled = True
        TallyA1.Enabled = True
        TallyA5.Enabled = True
  End Select
  
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SeedTotal_GotFocus[Form_fsub_LP_Seedling])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          TallyA1_Click
' Description:  Handles actions when control is clicked
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Russ DenBleyker, unknown
' Adapted:      Bonnie Campbell, April 1, 2016 - for NCPN tools
' Revisions:
'   RDB, unknown  - initial version
'   BLC, 4/1/2016 - added error handling, documentation, revised based on use of AddTallyValue & tally
'                   buttons disabling when tally values should not be available
' ---------------------------------
Private Sub TallyA1_Click()
On Error GoTo Err_Handler
  
  AddTallyValue Screen.PreviousControl, 1
'  'If Screen.PreviousControl.name = "SeedTotal" And Not IsNull(Me!Species) Then
'  If Screen.PreviousControl.name = "SeedTotal" Then
'    Screen.PreviousControl.Value = Screen.PreviousControl.Value + 1
'  End If
'  Screen.PreviousControl.SetFocus
  
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - TallyA1_Click[Form_fsub_LP_Seedling])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          TallyA5_Click
' Description:  Handles actions when control is clicked
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Russ DenBleyker, unknown
' Adapted:      Bonnie Campbell, April 1, 2016 - for NCPN tools
' Revisions:
'   RDB, unknown  - initial version
'   BLC, 4/1/2016 - added error handling, documentation, revised based on use of AddTallyValue & tally
'                   buttons disabling when tally values should not be available
' ---------------------------------
Private Sub TallyA5_Click()
On Error GoTo Err_Handler
  
'  'If Screen.PreviousControl.name = "SeedTotal" And Not IsNull(Me!Species) Then
'  If Screen.PreviousControl.name = "SeedTotal" Then
'    Screen.PreviousControl.Value = Screen.PreviousControl.Value + 5
'  End If
'  Screen.PreviousControl.SetFocus

    AddTallyValue Screen.PreviousControl, 5

  
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - TallyA5_Click[Form_fsub_LP_Seedling])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          TallyS1_Click
' Description:  Handles actions when control is clicked
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Russ DenBleyker, unknown
' Adapted:      Bonnie Campbell, April 1, 2016 - for NCPN tools
' Revisions:
'   RDB, unknown  - initial version
'   BLC, 4/1/2016 - added error handling, documentation, revised based on use of AddTallyValue & tally
'                   buttons disabling when tally values should not be available
' ---------------------------------
Private Sub TallyS1_Click()
On Error GoTo Err_Handler
  
'  'If Screen.PreviousControl.name = "SeedTotal" And Not IsNull(Me!Species) Then
'  If Screen.PreviousControl.name = "SeedTotal" Then
''    If Screen.PreviousControl.Value - 1 < 0 Then
''      MsgBox "Total cannot be negative.", , "Tree Seedlings"
''      Exit Sub
''    Else
'      Screen.PreviousControl.Value = Screen.PreviousControl.Value - 1
''    End If
'  End If
'  Screen.PreviousControl.SetFocus
  
    AddTallyValue Screen.PreviousControl, -1
  
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - TallyS1_Click[Form_fsub_LP_Seedling])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          TallyS5_Click
' Description:  Handles actions when control is clicked
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Russ DenBleyker, unknown
' Adapted:      Bonnie Campbell, April 1, 2016 - for NCPN tools
' Revisions:
'   RDB, unknown  - initial version
'   BLC, 4/1/2016 - added error handling, documentation, revised based on use of AddTallyValue & tally
'                   buttons disabling when tally values should not be available
' ---------------------------------
Private Sub TallyS5_Click()
On Error GoTo Err_Handler
  
'  If Screen.PreviousControl.name = "SeedTotal" And Not IsNull(Me!Species) Then
'    If Screen.PreviousControl.Value - 5 < 0 Then
'      MsgBox "Total cannot be negative.", , "Tree Seedlings"
'      Exit Sub
'    Else
'      Screen.PreviousControl.Value = Screen.PreviousControl.Value - 5
'    End If
'  End If
'  Screen.PreviousControl.SetFocus
  
      AddTallyValue Screen.PreviousControl, -5
  
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - TallyS5_Click[Form_fsub_LP_Seedling])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          TallyA0_Click
' Description:  Handles actions when control is clicked
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Russ DenBleyker, unknown
' Adapted:      Bonnie Campbell, April 1, 2016 - for NCPN tools
' Revisions:
'   RDB, unknown  - initial version
'   BLC, 4/1/2016 - added error handling, documentation, revised based on use of AddTallyValue & tally
'                   buttons disabling when tally values should not be available
' ---------------------------------
Private Sub TallyA0_Click()
On Error GoTo Err_Handler

    AddTallyValue Screen.PreviousControl, 0
    
'  If Screen.PreviousControl.name = "SeedTotal" And Not IsNull(Me!Species) Then
'    Screen.PreviousControl.Value = 0
'  End If
'  Screen.PreviousControl.SetFocus

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - TallyA0_Click[Form_fsub_LP_Seedling])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          ButtonDelete_Click
' Description:  Handles delete button actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Russ DenBleyker, unknown
' Adapted:      Bonnie Campbell, February 11, 2016 - for NCPN tools
' Revisions:
'   RDB, unknown  - initial version
'   BLC, 2/11/2016 - added error handling, documentation, refresh checkbox/no data collected
' ---------------------------------
Private Sub ButtonDelete_Click()
On Error GoTo Err_Handler

  Dim intReply As Integer
  
  intReply = MsgBox("Are you sure you want to delete this record?", vbYesNo, "Delete Record")
    If intReply = vbYes Then
      DoCmd.SetWarnings False
      DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
      DoCmd.DoMenuItem acFormBar, acEditMenu, 6, , acMenuVer70
      DoCmd.SetWarnings True
      Me.Requery
    End If

    '-----------------------------------
    ' update the NoDataCollected info IF no records now exist
    '-----------------------------------
    If Me.RecordsetClone.RecordCount = 0 Then
    
        Dim NoData As Scripting.Dictionary
        
        'remove the no data collected record
        Set NoData = SetNoDataCollected(Me.Parent.Form.Controls("Transect_ID"), "T", "1mBelt-TreeSeedling", 1)
    
        'update checkbox/rectangle
        Me.Parent.Form.Controls("cbxNoSeedlings") = 1
        Me.Parent.Form.Controls("cbxNoSeedlings").Enabled = True
        Me.Parent.Form.Controls("rctNoSeedlings").Visible = True
        
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ButtonDelete_Click[Form_fsub_LP_Seedling])"
    End Select
    Resume Exit_Handler
End Sub
