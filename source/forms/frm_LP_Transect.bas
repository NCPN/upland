﻿Version =21
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =124
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =14220
    DatasheetFontHeight =9
    ItemSuffix =40
    Left =1275
    Top =3360
    Right =15090
    Bottom =12990
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x8e9d18c9b254e340
    End
    RecordSource ="qry_LP_Transect"
    Caption ="frm_Canopy_Transect"
    BeforeInsert ="[Event Procedure]"
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
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            CanGrow = NotDefault
            Height =9900
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =120
                    Top =60
                    Width =630
                    Height =180
                    ColumnWidth =2310
                    FontSize =10
                    Name ="Transect_ID"
                    ControlSource ="Transect_ID"
                    StatusBarText ="Unique record identifier - primary key"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =840
                    Top =60
                    Width =630
                    Height =180
                    ColumnWidth =2310
                    FontSize =10
                    TabIndex =1
                    Name ="Event_ID"
                    ControlSource ="Event_ID"
                    StatusBarText ="M. Link to tbl_Locations (Loc_ID)"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1080
                    Top =60
                    Width =360
                    Height =300
                    ColumnWidth =465
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    ForeColor =255
                    Name ="Transect"
                    ControlSource ="Transect"
                    StatusBarText ="Transect number - 1, 2, or 3"

                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =60
                            Top =60
                            Width =1020
                            Height =300
                            FontSize =10
                            FontWeight =700
                            Name ="Transect_Label"
                            Caption ="Transect"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =3360
                    Top =60
                    Width =960
                    ColumnWidth =1035
                    TabIndex =3
                    Name ="Visit_Date"
                    ControlSource ="Visit_Date"
                    Format ="Short Date"
                    StatusBarText ="Date of visit."
                    InputMask ="99/99/0000;0;_"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2460
                            Top =60
                            Width =900
                            Height =240
                            FontWeight =700
                            Name ="Visit_Date_Label"
                            Caption ="Visit Date"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1620
                    Top =60
                    Width =306
                    Height =306
                    TabIndex =6
                    Name ="ButtonPrevious"
                    Caption ="Command14"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddadadadad1dadadaadadadad11adadaddadadad111dadada ,
                        0xadadad1111adadaddadad11111dadadaadadad1111adadaddadadad111dadada ,
                        0xadadadad11adadaddadadadad1dadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    OnKeyDown ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Previous Record"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1980
                    Top =60
                    Width =306
                    Height =306
                    TabIndex =7
                    Name ="ButtonNext"
                    Caption ="Command15"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddadada1adadadadaadadad11adadadaddadada111adadada ,
                        0xadadad1111adadaddadada11111adadaadadad1111adadaddadada111adadada ,
                        0xadadad11adadadaddadada1adadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    OnKeyDown ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Next Record"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =1650
                    Left =5340
                    Top =60
                    Width =1620
                    TabIndex =4
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Observer"
                    ControlSource ="Observer"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Contacts.Contact_ID, tlu_Contacts.Last_Name, tlu_Contacts.First_Name "
                        "FROM tlu_Contacts WHERE (((tlu_Contacts.Active)=1)); "
                    ColumnWidths ="0;810;840"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4500
                            Top =60
                            Width =840
                            Height =245
                            FontWeight =700
                            Name ="Observer_Label"
                            Caption ="Observer"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =1545
                    Left =7980
                    Top =60
                    Width =1620
                    TabIndex =5
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Recorder"
                    ControlSource ="Recorder"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Contacts.Contact_ID, tlu_Contacts.Last_Name, tlu_Contacts.First_Name "
                        "FROM tlu_Contacts WHERE (((tlu_Contacts.Active)=1)); "
                    ColumnWidths ="0;750;795"
                    OnKeyDown ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =7140
                            Top =60
                            Width =840
                            Height =245
                            FontWeight =700
                            Name ="Recorder_Label"
                            Caption ="Recorder"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =180
                    Top =420
                    Width =13800
                    Height =8400
                    TabIndex =8
                    Name ="fsub_LP_Intercept"
                    SourceObject ="Form.fsub_LP_Intercept"
                    LinkChildFields ="Transect_ID"
                    LinkMasterFields ="Transect_ID"

                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10260
                    Top =60
                    Width =1860
                    Height =300
                    FontWeight =700
                    TabIndex =9
                    Name ="ButtonVerify"
                    Caption ="Verify Soil Surface"
                    OnClick ="[Event Procedure]"
                    OnKeyDown ="[Event Procedure]"

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
Option Explicit

' =================================
' MODULE:       frm_LP_Transect
' Level:        Form module
' Version:      1.01
' Description:  data functions & procedures specific to LP transect data entry
'
' Source/date:  John R. Boetsch, June 2006
' Adapted:      Bonnie Campbell, 2/3/2016
' Revisions:    RDB - unknown  - 1.00 - initial version
'               BLC - 2/3/2016 - 1.01 - added documentation, revised to use transect overlay vs.
'                                       message box
' =================================

' ---------------------------------
' SUB:          Form_BeforeInsert
' Description:  Handles form pre-insert actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 3, 2016 - for NCPN tools
' Revisions:
'   JRB, 6/x/2006  - initial version
'   RDB, unknown   - ?
'   BLC, 2/3/2016  - added documentation
' ---------------------------------
Private Sub Form_BeforeInsert(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Default to Events Start Date if visit date is null
    If IsNull(Me.Parent!Start_Date) Then
      MsgBox "Missing site visit date."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
      GoTo Exit_Handler
    ElseIf IsNull(Me!Visit_Date) Then
      Me!Visit_Date = Me.Parent!Start_Date
    End If
    ' Create the GUID primary key value if necessary
    If IsNull(Me!Transect_ID) Then
        If GetDataType("tbl_LP_Transect", "Transect_ID") = dbText Then
            Me.Transect_ID = fxnGUIDGen
        End If
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[Form_frm_LP_Transect])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          ButtonPrevious_Click
' Description:  Handles form previous click actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 3, 2016 - for NCPN tools
' Revisions:
'   JRB, 6/x/2006  - initial version
'   RDB, unknown   - ?
'   BLC, 2/3/2016  - added documentation, revised to use transect # overlay vs. message box
' ---------------------------------
Private Sub ButtonPrevious_Click()
On Error GoTo Err_Handler

    Dim intTransect As Byte
    Dim db As dao.Database
    Dim Points As dao.Recordset
    Dim strSQL As String
        
'  If IsNull(Me!Recorder) And IsNull(Me!Observer) Then
'      DoCmd.CancelEvent
'      SendKeys "{ESC}"
'  End If
  If Me!Transect = 1 Then
    MsgBox "Already on first transect"
  Else
    intTransect = Me!Transect
    DoCmd.GoToRecord , , acPrevious
'    DoCmd.GoToRecord , , 2
    Me!Transect = intTransect - 1
    ' Set SQL
    Set db = CurrentDb
    strSQL = "SELECT [Point] FROM [tbl_LP_Intercept] WHERE [Transect_ID] = '" & Me![Transect] & "'"
    Set Points = db.OpenRecordset(strSQL)
    If Points.EOF Or IsNull(Points!Point) Then
      Me!fsub_LP_Intercept.Form!ButtonInitialize.ForeColor = 8421376
      Me!fsub_LP_Intercept.Requery
    End If
    
    '---------------------------
    'display overlay - 2/3/2016 - BLC
    '---------------------------
    'MsgBox "You are on transect " & Me!Transect & ".", 0, "Transect Verify"
    DoCmd.OpenForm "frm_Transect_Overlay", OpenArgs:=Me!Transect
    '---------------------------
    
  End If
  
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ButtonPrevious_Click[Form_frm_LP_Transect])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          ButtonNext_Click
' Description:  Handles next button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 3, 2016 - for NCPN tools
' Revisions:
'   JRB, 6/x/2006  - initial version
'   RDB, unknown   - ?
'   BLC, 2/3/2016  - added documentation, revised to use transect # overlay vs. message box
' ---------------------------------
Private Sub ButtonNext_Click()
On Error GoTo Err_Handler

    Dim db As dao.Database
    Dim Points As dao.Recordset
    Dim strSQL As String
    
On Error GoTo Err_Handler

'  If IsNull(Me!Recorder) And IsNull(Me!Observer) Then
'    MsgBox "You must record data in this transect before moving to the next."
'    GoTo Exit_ButtonNext_Click
'  End If
  Dim intTransect As Byte
    If IsNull(Me!Transect) Then
      Me!Transect = 1
    End If
  If Me!Transect = 3 Then
    MsgBox "Three transects maximum!"
    GoTo Exit_Handler
  Else
    intTransect = Me!Transect
    DoCmd.GoToRecord , , acNext
    Me!Transect = intTransect + 1
  
    '---------------------------
    'display overlay - 2/3/2016 - BLC
    '---------------------------
    'MsgBox "You are on transect " & Me!Transect & ".", 0, "Transect Verify"
    DoCmd.OpenForm "frm_Transect_Overlay", OpenArgs:=Me!Transect
    '---------------------------
  End If
  
    ' Set SQL to figure out what color button is needed
    Set db = CurrentDb
    strSQL = "SELECT [Point] FROM [tbl_LP_Intercept] WHERE [Transect_ID] = '" & Me![Transect_ID] & "'"
    Set Points = db.OpenRecordset(strSQL)
    
    If Points.EOF Then
      Me!fsub_LP_Intercept.Form!ButtonInitialize.ForeColor = 8421376
    Else
      Me!fsub_LP_Intercept.Form!ButtonInitialize.ForeColor = 255
      Me!fsub_LP_Intercept.Form.Requery
    End If
    Points.Close

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ButtonNext_Click[Form_frm_LP_Transect])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub ButtonVerify_Click()
On Error GoTo Err_ButtonVerify_Click

  Dim db As dao.Database
  Dim Surface As dao.Recordset
  Dim strSQL As String
  
  strSQL = "SELECT Point from tbl_LP_Intercept WHERE Transect_ID = '" & Me!Transect_ID & "' AND Surface IS NULL"
  Set db = CurrentDb
  Set Surface = db.OpenRecordset(strSQL)
  If Not Surface.EOF Then
    MsgBox "Transect contains at least one empty surface code."
  Else
    MsgBox "OK!"
  End If

Exit_ButtonVerify_Click:
    Exit Sub

Err_ButtonVerify_Click:
    MsgBox Err.Description
    Resume Exit_ButtonVerify_Click
    
End Sub

Private Sub ButtonNext_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub ButtonPrevious_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub ButtonVerify_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Observer_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Recorder_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub

Private Sub Visit_Date_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Ignore Page Down and Page Up keys for they will cycle through records
  Select Case KeyCode
    Case 33, 34
      KeyCode = 0
    End Select
End Sub
