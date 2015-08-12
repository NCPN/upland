Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =127
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6480
    DatasheetFontHeight =9
    ItemSuffix =54
    Left =15060
    Top =3636
    Right =18612
    Bottom =8196
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x9becc7edac0fe340
    End
    RecordSource ="tbl_Locations"
    BeforeUpdate ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
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
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin Section
            CanGrow = NotDefault
            Height =4320
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1200
                    Top =180
                    Width =3825
                    Height =480
                    FontSize =18
                    FontWeight =700
                    Name ="Label0"
                    Caption ="Edit Fuels Transect"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =180
                    Top =120
                    Width =960
                    TabIndex =9
                    Name ="Location_ID"
                    ControlSource ="Location_ID"
                    StatusBarText ="M. Location identifier (Loc_ID)"

                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2025
                    Top =2160
                    Width =810
                    Height =300
                    TabIndex =1
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Bearing_A"
                    ControlSource ="Bearing_A"
                    StatusBarText ="Bearing of the plot slope + 180 in degrees"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =2025
                            Top =1920
                            Width =810
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Bearing_A_Label"
                            Caption ="A"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =255
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2820
                    Top =2160
                    Width =810
                    Height =300
                    TabIndex =3
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Bearing_B"
                    ControlSource ="Bearing_B"
                    StatusBarText ="Bearing of transect 1 in degrees"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =223
                            TextAlign =2
                            Left =2820
                            Top =1920
                            Width =810
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Bearing_B_Label"
                            Caption ="B"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =127
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3645
                    Top =2160
                    Width =810
                    Height =300
                    TabIndex =5
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Bearing_C"
                    ControlSource ="Bearing_C"
                    StatusBarText ="Bearing of transect 3 + 180"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =3645
                            Top =1920
                            Width =810
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Bearing_C_Label"
                            Caption ="C"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4485
                    Top =2160
                    Width =810
                    Height =300
                    TabIndex =7
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Bearing_D"
                    ControlSource ="Bearing_D"
                    StatusBarText ="Bearing of the plot slope"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =4485
                            Top =1920
                            Width =810
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Bearing_D_Label"
                            Caption ="D"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OldBorderStyle =1
                    OverlapFlags =127
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2025
                    Top =2460
                    Width =810
                    Height =300
                    TabIndex =2
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Slope_A"
                    ControlSource ="Slope_A"
                    StatusBarText ="Slope of transect A to nearest half percent"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =720
                            Top =2160
                            Width =1290
                            Height =300
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Slope_A_Label"
                            Caption ="bearing (deg)"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OldBorderStyle =1
                    OverlapFlags =255
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2820
                    Top =2460
                    Width =810
                    Height =300
                    TabIndex =4
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Slope_B"
                    ControlSource ="Slope_B"
                    StatusBarText ="Slope of transect B to nearest half percent"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =119
                            TextAlign =2
                            Left =720
                            Top =2460
                            Width =1290
                            Height =300
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Slope_B_Label"
                            Caption ="slope (%)"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OldBorderStyle =1
                    OverlapFlags =119
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3645
                    Top =2460
                    Width =810
                    Height =300
                    TabIndex =6
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Slope_C"
                    ControlSource ="Slope_C"
                    StatusBarText ="Slope of transect C to nearest half percent"

                End
                Begin TextBox
                    DecimalPlaces =1
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4485
                    Top =2460
                    Width =810
                    Height =300
                    TabIndex =8
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Slope_D"
                    ControlSource ="Slope_D"
                    StatusBarText ="Slope of transect C to nearest half percent"

                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1980
                    Top =1560
                    Width =1860
                    Height =300
                    FontSize =10
                    FontWeight =700
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label45"
                    Caption ="Fuels Transect"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2580
                    Top =3180
                    Width =1035
                    Height =405
                    TabIndex =10
                    Name ="ButtonClose"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =735
                    Left =3300
                    Top =1020
                    Width =2100
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Recorder"
                    RowSourceType ="Table/Query"
                    RowSource ="qry_Contacts"
                    ColumnWidths ="0;735"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =660
                            Top =1020
                            Width =2640
                            Height =245
                            FontWeight =700
                            Name ="Observer_Label"
                            Caption ="Changes made by (Required):"
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

Private Sub ButtonClose_Click()
On Error GoTo Err_ButtonClose_Click


    DoCmd.Close

Exit_ButtonClose_Click:
    Exit Sub

Err_ButtonClose_Click:
    MsgBox Err.Description
    Resume Exit_ButtonClose_Click
    
End Sub

' ---------------------------------
' SUB:          Form_BeforeUpdate
' Description:  Populate centroid UTMs from tbl_Location_History
' Assumptions:  -
' Parameters:   Cancel - species to check (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Russ DenBleyker, date unkown, Northern Colorado Plateau Network
' Adapted:      -
' Revisions:
'   RD  - ?         - initial version
'   BLC - 8/11/2015 - fixed bug improperly populating plot centroid UTMs with
'                     tbl_Location_History deprecated Plot_E_Coord & Plot_N_Coord vs.
'                     E_Coord & N_Coord values, updated error handling & added documentation
' ---------------------------------
Private Sub Form_BeforeUpdate(Cancel As Integer)
    On Error GoTo Err_Handler

    Dim db As Database
    Dim History As DAO.Recordset
    Dim OldLocation As DAO.Recordset
    Dim strSQL As String
    
    If IsNull(Me!Recorder) Then
      MsgBox "You must select a recorder name!"
      Me.Undo
      GoTo Exit_Sub
    Else
      Set db = CurrentDb
      strSQL = "Select * from tbl_Locations WHERE Location_ID = '" & Me!Location_ID & "'"
      Set OldLocation = db.OpenRecordset(strSQL)  '  Get unmodified location record
      If OldLocation.EOF Then
        MsgBox "Location record not found."
        GoTo Exit_Sub
      End If
      Set History = db.OpenRecordset("tbl_Location_History")
        History.AddNew                     ' Create a Location History record
        History!Location_History_ID = fxnGUIDGen
        History!Location_ID = Me!Location_ID
        History!Modify_Date = Now()        ' Date of update
        History!Recorder = Me!Recorder     ' Person committing update
        History!Unit_Code = OldLocation!Unit_Code
        History!Plot_ID = OldLocation!Plot_ID
        
        'populate plot centroid UTMs E & N Coord
        History!E_Coord = OldLocation!E_Coord
        History!N_Coord = OldLocation!N_Coord
        
        History!Plot_Slope = OldLocation!Plot_Slope
        History!Plot_Aspect = OldLocation!Plot_Aspect
        History!Azimuth = OldLocation!Azimuth
        History!T1O_UTME = OldLocation!T1O_UTME
        History!T1O_UTMN = OldLocation!T1O_UTMN
        History!T1O_Rebar = OldLocation!T1O_Rebar
        History!T1E_UTME = OldLocation!T1E_UTME
        History!T1E_UTMN = OldLocation!T1E_UTMN
        History!T1E_Rebar = OldLocation!T1E_Rebar
        History!T1_Elevation = OldLocation!T1_Elevation
        History!T2O_UTME = OldLocation!T2O_UTME
        History!T2O_UTMN = OldLocation!T2O_UTMN
        History!T2O_Rebar = OldLocation!T2O_Rebar
        History!T2E_UTME = OldLocation!T2E_UTME
        History!T2E_UTMN = OldLocation!T2E_UTMN
        History!T2E_Rebar = OldLocation!T2E_Rebar
        History!T2_Elevation = OldLocation!T2_Elevation
        History!T3O_UTME = OldLocation!T3O_UTME
        History!T3O_UTMN = OldLocation!T3O_UTMN
        History!T3O_Rebar = OldLocation!T3O_Rebar
        History!T3E_UTME = OldLocation!T3E_UTME
        History!T3E_UTMN = OldLocation!T3E_UTMN
        History!T3E_Rebar = OldLocation!T3E_Rebar
        History!T3_Elevation = OldLocation!T3_Elevation
        History!Plot_Directions = OldLocation!Plot_Directions
        
        ' Fuels bearings and slopes
        History!Bearing_A = OldLocation!Bearing_A
        History!Bearing_B = OldLocation!Bearing_B
        History!Bearing_C = OldLocation!Bearing_C
        History!Bearing_D = OldLocation!Bearing_D
        
        History!Slope_A = OldLocation!Slope_A
        History!Slope_B = OldLocation!Slope_B
        History!Slope_C = OldLocation!Slope_C
        History!Slope_D = OldLocation!Slope_D
        
        ' Plot side slopes
        History!SlopeA = OldLocation!SlopeA
        History!SlopeAUD = OldLocation!SlopeAUD
        History!SlopeB = OldLocation!SlopeB
        History!SlopeBUD = OldLocation!SlopeBUD
        History!SlopeC = OldLocation!SlopeC
        History!SlopeCUD = OldLocation!SlopeCUD
        History!SlopeD = OldLocation!SlopeD
        History!SlopeDUD = OldLocation!SlopeDUD
        History.Update
        History.Close
        Set History = Nothing
        OldLocation.Close
        Set OldLocation = Nothing
    End If
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - BeforeUpdate[Form_frm_Edit_Fuel_Transect])"
    End Select
    Resume Exit_Sub

End Sub
