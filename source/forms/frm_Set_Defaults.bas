Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4380
    DatasheetFontHeight =10
    ItemSuffix =9
    Left =9288
    Top =2892
    Right =13668
    Bottom =6060
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa3c57b9aedcee240
    End
    RecordSource ="tsys_App_Defaults"
    Caption =" Set application default values"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin Section
            Height =3180
            BackColor =11056034
            Name ="Detail"
            Begin
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =3600
                    Left =972
                    Top =960
                    Width =1245
                    Height =252
                    FontSize =9
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"10\""
                    Name ="cmbPark"
                    ControlSource ="Park"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Parks.ParkCode, tlu_Parks.ParkName FROM tlu_Parks ORDER BY tlu_Parks."
                        "ParkCode; "
                    ColumnWidths ="720;2880"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =960
                            Width =480
                            Height =255
                            FontSize =9
                            FontWeight =700
                            BackColor =11056034
                            Name ="labPark"
                            Caption ="Park"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =972
                    Top =600
                    Width =3165
                    Height =252
                    FontSize =9
                    BoundColumn =1
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"0\""
                    Name ="cbxUser"
                    ControlSource ="User_name"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Contact_ID, Last_Name & \"_\" & First_Name AS Expr1 FROM tlu_Contacts WHE"
                        "RE (((tlu_Contacts.Active)=1)) ORDER BY tlu_Contacts.Last_Name, tlu_Contacts.Fir"
                        "st_Name;"
                    ColumnWidths ="0;3165"
                    FontName ="Arial"
                    OnChange ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =600
                            Width =468
                            Height =252
                            FontSize =9
                            FontWeight =700
                            BackColor =11056034
                            Name ="labUser"
                            Caption ="User"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =180
                    Top =2280
                    Width =3963
                    Height =417
                    FontSize =9
                    TabIndex =4
                    Name ="Soil_Survey_Area"
                    ControlSource ="Soil_Survey_Area"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =180
                            Top =2040
                            Width =1617
                            Height =252
                            FontSize =9
                            FontWeight =700
                            BackColor =11056034
                            Name ="labActivity"
                            Caption ="Soil Survey Area"
                            FontName ="Arial"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3480
                    Top =120
                    Width =720
                    Height =354
                    FontSize =9
                    FontWeight =700
                    TabIndex =6
                    ForeColor =0
                    Name ="cmdOK"
                    Caption ="OK"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3120
                    Top =1020
                    Width =1035
                    FontSize =9
                    FontWeight =700
                    TabIndex =5
                    ForeColor =0
                    Name ="cmdNewUser"
                    Caption ="New user"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Add a new user"

                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =972
                    Top =1320
                    Width =1245
                    FontSize =9
                    TabIndex =2
                    Name ="cmbDatum"
                    ControlSource ="Datum"
                    RowSourceType ="Value List"
                    RowSource ="\"NAD27\";\"NAD83\";\"WGS84\""
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =1320
                            Width =672
                            Height =252
                            FontSize =9
                            FontWeight =700
                            BackColor =11056034
                            Name ="labDatum"
                            Caption ="Datum"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =1440
                    Left =960
                    Top =1680
                    Width =1260
                    TabIndex =3
                    Name ="Zone"
                    ControlSource ="Zone"
                    RowSourceType ="Value List"
                    RowSource ="12;13"
                    ColumnWidths ="1440"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =180
                            Top =1680
                            Width =660
                            Height =245
                            FontSize =9
                            FontWeight =700
                            Name ="UTM Zone_Label"
                            Caption ="Zone"
                            FontName ="Arial"
                            EventProcPrefix ="UTM_Zone_Label"
                        End
                    End
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
' MODULE:       frm_Set_Defaults
' Level:        Form module
' Version:      1.01
' Description:  data functions & procedures specific to setting data entry/edit defaults
'               (standard module for setting application defaults)
' Data source:  tsys_App_Defaults
' Data access:  edit only, no deletions
' Pages:        none
' Functions:    none
' References:   none
' Source/date:  John R. Boetsch, May 16, 2006
' Adapted:      -
' Revisions:    JRB - 5/16/2006  - 1.00 - initial version
'               BLC - 3/7/2016 - 1.01 - add cbxUser_Change() to set UserID tempVar
'                                       so user changing data can be identified in later forms,
'                                       renamed cmbUser to cbxUser, added documentation for
'                                       all subroutines
' =================================

' ---------------------------------
' SUB:          cbxUser_Change
' Description:  Handles actions when cbxUser changes
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, March 7, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC, 3/7/2016  - initial version
' ---------------------------------
Private Sub cbxUser_Change()
On Error GoTo Err_Handler

    'set TempVars value to the User_ID (cbxUser.Column(0))
    If Not IsNull(TempVars.item("User_ID")) Then
        TempVars.item("User_ID") = cbxUser.Column(0)
    Else
        TempVars.Add "User_ID", cbxUser.Column(0)
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxUser_Change[Form_frm_Set_Defaults])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxUser_NotInList
' Description:  Handles actions when cbxUser is not in the list
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  John Boetsch, May 16, 2006
' Adapted:      -
' Revisions:
'   JRB, 5/16/2006 - initial version
'   BLC, 3/7/2016  - added documentation
' ---------------------------------
Private Sub cbxUser_NotInList(NewData As String, Response As Integer)
On Error GoTo Err_Handler

    Me.ActiveControl.Undo

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxUser_NotInList[Form_frm_Set_Defaults])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cmdNewUser_Click
' Description:  Handles actions when clicking cmdNewUser
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  John Boetsch, May 16, 2006
' Adapted:      -
' Revisions:
'   JRB, 5/16/2006 - initial version
'   BLC, 3/7/2016  - added documentation
' ---------------------------------
Private Sub cmdNewUser_Click()
    On Error GoTo Err_Handler
    
    ' Open the contacts form
    DoCmd.OpenForm "frm_Contacts"

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          cmbPark_AfterUpdate
' Description:  Handles actions after updating cmbPark
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  John Boetsch, May 16, 2006
' Adapted:      -
' Revisions:
'   JRB, 5/16/2006 - initial version
'   BLC, 3/7/2016  - added documentation
' ---------------------------------
Private Sub cmbPark_AfterUpdate()
    On Error GoTo Err_Handler

    Dim strMsg As String
    Dim strZone As String
    Dim strDatum As String
    Dim strNetwork As String

    If Not IsNull(Me!cmbDatum) Or Not IsNull(Me!Zone) Then
    ' On changing the park, prompt for resetting the datum and declination
        strZone = Nz(Me.Zone, "---")
        strDatum = Nz(Me.cmbDatum, "---")
        strMsg = "Changing parks requires verification of other settings." & vbCrLf & vbCrLf
        strMsg = strMsg & "Datum: " & strDatum & "  Zone: " & strZone & vbCrLf & vbCrLf
        strMsg = strMsg & "Would you like to keep these settings?"
        If MsgBox(strMsg, vbYesNo, "Verify park info") = vbNo Then
            Me.cmbDatum = Null
            Me.Zone = Null
        End If
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cmbPark_AfterUpdate[Form_frm_Set_Defaults])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cmdOK_Click
' Description:  Handles cmdOK click actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  John Boetsch, May 16, 2006
' Adapted:      -
' Revisions:
'   JRB, 5/16/2006 - initial version
'   BLC, 3/7/2016  - added documentation
' ---------------------------------
Private Sub cmdOK_Click()
    On Error GoTo Err_Handler

    Dim varOpenArgs As Variant
    
    varOpenArgs = Me.OpenArgs
    
    ' Make sure the information is valid before updating the record
    If varOpenArgs <> 0 Then
        '  Verify that the critical data elements have been completed before saving
        If IsNull(Me.User_name) Then
            MsgBox "Please indicate the user name", vbOKOnly, "Validation error"
            Me.cbxUser.SetFocus
            GoTo Exit_Handler
        ElseIf IsNull(Me.Park) Then
            MsgBox "Please indicate the park", vbOKOnly, "Validation error"
            Me.cmbPark.SetFocus
            GoTo Exit_Handler
        End If
    End If

    DoCmd.Close acForm, Me.Name, acSaveNo
    DoCmd.OpenForm "frm_Switchboard"
    Select Case varOpenArgs
        Case 1
            DoCmd.OpenForm "frm_Data_Gateway", , , , , , varOpenArgs
        Case 2
            DoCmd.OpenForm "frm_Browser", , , , , , varOpenArgs
        Case 3
            DoCmd.OpenForm "frm_Select_Not_Present", , , , , , varOpenArgs
        Case 4
            ' opened by switchboard only ... do nothing
        Case Else
            MsgBox "Error: OpenArgs property out of range", vbCritical
    End Select

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cmdOK_Click[Form_frm_Set_Defaults])"
    End Select
    Resume Exit_Handler
End Sub
