Version =21
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =124
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    DatasheetFontHeight =9
    ItemSuffix =34
    Left =5490
    Top =1830
    Right =12690
    Bottom =7335
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x1385341e7574e340
    End
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
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
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Line
            BorderLineStyle =0
            BorderThemeColorIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin Section
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =5520
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ComboBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    ListWidth =525
                    Left =3855
                    Top =960
                    Width =900
                    ColumnInfo ="\"\";\"@\";\"10\";\"510\""
                    Name ="cbxPark"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT distinct tbl_Revisit_List.PARK FROM tbl_Revisit_List ORDER BY tbl_Revisit"
                        "_List.PARK;"
                    ColumnWidths ="525"
                    AfterUpdate ="[Event Procedure]"
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =3855
                    LayoutCachedTop =960
                    LayoutCachedWidth =4755
                    LayoutCachedHeight =1200
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =1635
                            Top =960
                            Width =2085
                            Height =245
                            FontWeight =700
                            Name ="Select a park if desired_Label"
                            Caption ="Select a park if desired"
                            EventProcPrefix ="Select_a_park_if_desired_Label"
                            LayoutCachedLeft =1635
                            LayoutCachedTop =960
                            LayoutCachedWidth =3720
                            LayoutCachedHeight =1205
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =67
                    Left =4200
                    Top =3120
                    Width =1334
                    Height =300
                    TabIndex =7
                    Name ="Button_Close"
                    Caption ="&Close Form"
                    OnClick ="[Event Procedure]"
                    UnicodeAccessKey =67

                    LayoutCachedLeft =4200
                    LayoutCachedTop =3120
                    LayoutCachedWidth =5534
                    LayoutCachedHeight =3420
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin ComboBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ListWidth =720
                    Left =3840
                    Top =1200
                    Width =900
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"3\";\"2\""
                    Name ="cbxYear"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qry_Event_Date.Visit_Year FROM qry_Event_Date; "
                    ColumnWidths ="720"

                    LayoutCachedLeft =3840
                    LayoutCachedTop =1200
                    LayoutCachedWidth =4740
                    LayoutCachedHeight =1440
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =95
                            Left =1620
                            Top =1200
                            Width =2100
                            Height =245
                            FontWeight =700
                            Name ="Select a date if desired_Label"
                            Caption ="Select a year if desired"
                            EventProcPrefix ="Select_a_date_if_desired_Label"
                            LayoutCachedLeft =1620
                            LayoutCachedTop =1200
                            LayoutCachedWidth =3720
                            LayoutCachedHeight =1445
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2400
                    Top =300
                    Width =2280
                    Height =390
                    FontSize =14
                    FontWeight =700
                    Name ="Label6"
                    Caption ="Revisit Reports"
                    LayoutCachedLeft =2400
                    LayoutCachedTop =300
                    LayoutCachedWidth =4680
                    LayoutCachedHeight =690
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =71
                    Left =4200
                    Top =2340
                    Width =1320
                    Height =480
                    TabIndex =6
                    Name ="Button_rpt_by_Park"
                    Caption ="&Generate Reports"
                    OnClick ="=Get_Reports_Setup()"
                    UnicodeAccessKey =71

                    LayoutCachedLeft =4200
                    LayoutCachedTop =2340
                    LayoutCachedWidth =5520
                    LayoutCachedHeight =2820
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3840
                    Top =1440
                    Width =900
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"3\";\"2\""
                    Name ="cbxPlot"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Revisit_List.Plot FROM tbl_Revisit_List WHERE (((tbl_Revisit_List.Par"
                        "k) = 'CEBR' )) ORDER BY tbl_Revisit_List.Plot;"
                    ColumnWidths ="420"
                    OnGotFocus ="[Event Procedure]"

                    LayoutCachedLeft =3840
                    LayoutCachedTop =1440
                    LayoutCachedWidth =4740
                    LayoutCachedHeight =1680
                    Begin
                        Begin Label
                            OverlapFlags =87
                            Left =1620
                            Top =1440
                            Width =2100
                            Height =245
                            FontWeight =700
                            Name ="Plot_Select_Label"
                            Caption ="Select a plot if desired"
                            LayoutCachedLeft =1620
                            LayoutCachedTop =1440
                            LayoutCachedWidth =3720
                            LayoutCachedHeight =1685
                        End
                    End
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =4860
                    Top =1440
                    Width =1740
                    Height =420
                    ForeColor =16711680
                    Name ="lblPlotHint"
                    Caption ="Park selection required to select a plot."
                    LayoutCachedLeft =4860
                    LayoutCachedTop =1440
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =1860
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =900
                    Top =3720
                    Width =2460
                    Height =480
                    FontSize =13
                    FontWeight =900
                    TabIndex =8
                    ForeColor =1643706
                    Name ="btnGetItAll"
                    Caption ="I WANT IT ALL!!"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="And I Want It NOW!"

                    LayoutCachedLeft =900
                    LayoutCachedTop =3720
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =4200
                    UseTheme =1
                    BackColor =16764057
                    HoverColor =967423
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CheckBox
                    OverlapFlags =93
                    AccessKey =79
                    Left =1680
                    Top =2760
                    TabIndex =4
                    BorderColor =10921638
                    Name ="chkTrees"
                    DefaultValue ="Yes"
                    UnicodeAccessKey =79
                    GridlineColor =10921638

                    LayoutCachedLeft =1680
                    LayoutCachedTop =2760
                    LayoutCachedWidth =1940
                    LayoutCachedHeight =3000
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =1910
                            Top =2730
                            Width =1395
                            Height =240
                            Name ="lblTrees"
                            Caption ="&Overstory Census"
                            LayoutCachedLeft =1910
                            LayoutCachedTop =2730
                            LayoutCachedWidth =3305
                            LayoutCachedHeight =2970
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    AccessKey =83
                    Left =1680
                    Top =3150
                    TabIndex =5
                    BorderColor =10921638
                    Name ="chkSpecies"
                    DefaultValue ="Yes"
                    UnicodeAccessKey =83
                    GridlineColor =10921638

                    LayoutCachedLeft =1680
                    LayoutCachedTop =3150
                    LayoutCachedWidth =1940
                    LayoutCachedHeight =3390
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =1910
                            Top =3120
                            Width =1395
                            Height =240
                            Name ="lblSpecies"
                            Caption ="&Species List"
                            LayoutCachedLeft =1910
                            LayoutCachedTop =3120
                            LayoutCachedWidth =3305
                            LayoutCachedHeight =3360
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    AccessKey =80
                    Left =1680
                    Top =2370
                    TabIndex =3
                    BorderColor =10921638
                    Name ="chkOverview"
                    DefaultValue ="Yes"
                    UnicodeAccessKey =80
                    GridlineColor =10921638

                    LayoutCachedLeft =1680
                    LayoutCachedTop =2370
                    LayoutCachedWidth =1940
                    LayoutCachedHeight =2610
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =1910
                            Top =2340
                            Width =1395
                            Height =240
                            Name ="lblOverview"
                            Caption ="&Plot Revisit Report"
                            LayoutCachedLeft =1910
                            LayoutCachedTop =2340
                            LayoutCachedWidth =3305
                            LayoutCachedHeight =2580
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =1560
                    Top =2280
                    Width =1860
                    Height =1200
                    BorderColor =10921638
                    Name ="Box24"
                    GridlineColor =10921638
                    LayoutCachedLeft =1560
                    LayoutCachedTop =2280
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =3480
                End
                Begin Label
                    OverlapFlags =247
                    Left =1620
                    Top =2040
                    Width =780
                    Height =240
                    Name ="Label25"
                    Caption ="INCLUDE:"
                    LayoutCachedLeft =1620
                    LayoutCachedTop =2040
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =2280
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =3660
                    Top =3720
                    Width =2760
                    Height =480
                    FontSize =13
                    FontWeight =900
                    TabIndex =9
                    ForeColor =1643706
                    Name ="btnFeelingLucky"
                    Caption ="I'M FEELING LUCKY"
                    ControlTipText ="And I want to print EVERYTHING"

                    LayoutCachedLeft =3660
                    LayoutCachedTop =3720
                    LayoutCachedWidth =6420
                    LayoutCachedHeight =4200
                    UseTheme =1
                    BackColor =16764057
                    HoverColor =967423
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    Left =420
                    Top =5040
                    Width =1950
                    Height =300
                    TabIndex =10
                    Name ="btnOT"
                    Caption ="Overstory Plot Revisit"
                    OnClick ="=Get_Old_Forms()"

                    LayoutCachedLeft =420
                    LayoutCachedTop =5040
                    LayoutCachedWidth =2370
                    LayoutCachedHeight =5340
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    Left =2610
                    Top =5040
                    Width =1950
                    Height =300
                    TabIndex =11
                    Name ="btnPR"
                    Caption ="Plot Revisit Data Sheet"
                    OnClick ="=Get_Old_Forms()"

                    LayoutCachedLeft =2610
                    LayoutCachedTop =5040
                    LayoutCachedWidth =4560
                    LayoutCachedHeight =5340
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    Left =4800
                    Top =5040
                    Width =1950
                    Height =300
                    TabIndex =12
                    Name ="btnSP"
                    Caption ="Species Presence by Plot"
                    OnClick ="=Get_Old_Forms()"

                    LayoutCachedLeft =4800
                    LayoutCachedTop =5040
                    LayoutCachedWidth =6750
                    LayoutCachedHeight =5340
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Rectangle
                    OverlapFlags =215
                    Left =300
                    Top =4920
                    Width =6600
                    Height =540
                    BorderColor =10921638
                    Name ="Box30"
                    GridlineColor =10921638
                    LayoutCachedLeft =300
                    LayoutCachedTop =4920
                    LayoutCachedWidth =6900
                    LayoutCachedHeight =5460
                End
                Begin Label
                    OverlapFlags =85
                    Left =2940
                    Top =4560
                    Width =1320
                    Height =240
                    Name ="Label31"
                    Caption ="Open Prior Forms"
                    LayoutCachedLeft =2940
                    LayoutCachedTop =4560
                    LayoutCachedWidth =4260
                    LayoutCachedHeight =4800
                End
                Begin Line
                    OverlapFlags =85
                    Left =420
                    Top =4680
                    Width =2400
                    Name ="Line32"
                    GridlineColor =10921638
                    LayoutCachedLeft =420
                    LayoutCachedTop =4680
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =4680
                End
                Begin Line
                    OverlapFlags =85
                    Left =4380
                    Top =4680
                    Width =2400
                    Name ="Line33"
                    GridlineColor =10921638
                    LayoutCachedLeft =4380
                    LayoutCachedTop =4680
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =4680
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
Dim i As Integer
Dim dBtitle As String
Dim resp As Integer

' =========================================
' MODULE:       Form_frm_revRender_Revisit_Reports  (formerly frm_Species_Report_Select)
' Level:        Form module
' Version:      2.00
' Description:  data functions and procedures for producting revisit reports (including OT & Spp)
'               (formerly: data functions & procedures specific to species report by park)
'
' Source/date:  Bonnie Campbell, 2/2/2016
' Revisions:    RDB - unknown  - 1.00 - initial version
'               BLC - 2/2/2016 - 1.01 - added documentation, no data collected integration
'               BLC - 3/7/2016 - 1.02 - fixed strSQL replacement issue in Button_rpt_by_Park_Click
'               BLC - 3/8/2016 - 1.03 - fixed old data refresh issue (Button_rpt_by_Park_Click),
'                                       added documentation & error handling to other subroutines
'               BLC - 3/09/2016 - 1.04 - added primary key & index to improve report performance (Button_rpt_by_Park_Click)
'               BLC - 3/16/2016 - 1.05 - added enabling plot dropdown when park is chosen
'                                                           (Park_Code_AfterUpdate, Park_Code_Change)
'               AZ  - 3/26/2022 - 2.00 - this is now the main form for generating reports; many, many changes
' =========================================


' LIST OF PROCEDURES:
'---------------------------------------------------------------
'---------------------------------------------------------------
' SUB:    Form_Load        (sets dBtitle to title of current database, using the form's opening arguments)
' SUB:    cbxPlot_GotFocus (prevents selection of plot if park combo box is null)
' SUB:    cbxPark_Change   (enables and clears plot combo box)
' SUB:    cbxPlot_AfterUpdate  (populates combo plot rowsource w/appropriate plots)
' SUB:    Button_Close_Click   (closes form)
'---------------------------------------------------------------
' FUNC:*  Get_Reports_Setup (checks for templates; creates species table; checks for need to loop parks/plots
' SUB:    Get_Reportz       (directory mgmt; calls procedures to produce individual reports
' SUB:*   sub_Loop_Parks    (currently not doing anything)
' SUB:    sub_Loop_Plots    (loops through plots for one park; calls Get_Reportz; tracks progress)
' SUB:    Generate_Species_Table (generates rollup table which is the datasource for the species table)
' SUB:    getPlotInfo       (generates one plot overview report and saves it)
' SUB:    getTrees          (generates overstory census report for one plot and saves it)
' SUB:    getSpecies        (generates species list for one plot and saves it)
' FUNC:   MonTrees          (determines if plot has any monument trees)
' FUNC:   getYear           (gets year of most recent plot visit)
' FUNC:   MkMyDir           (public function in another module; creates needed directories)
' SUB:    btnGetItAll_Click (under construction)
'
'*these two procedures use the title of the database; this is set by the form's opening arguments:
' (Forms!frm_Switchboard.btnQA OnClick event)
'---------------------------------------------------------------
'---------------------------------------------------------------

' Subroutine:   Form_Load
' Description:  Sets value of dBtitle to opening arguments
' Assumptions:  Those opening arguments are actually the title of the database.
' Parameters:
' Returns:
' Throws:       none
' References:   none
' Source/date:  AZ, March 2022
' Adapted:      -
' Revisions:
'   AZ  - 3/28/2022  - initial version

Private Sub Form_Load()
    dBtitle = Me.OpenArgs
    Debug.Print dBtitle
    'MsgBox "Title must match frm_switchboard.OpenArgs, which is currently:" & vbNewLine & vbNewLine _
                & dBtitle & vbNewLine & vbNewLine _
                & "If it doesn't, close the form and fix it."
           
End Sub

' SUB:          cbxPlot_GotFocus (formerly Plot_GotFocus)
' Description:  Contains actions occuring on plot focus
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Russ DenBleyker, unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   RD - ?          - initial version
'   BLC - 3/8/2016  - added documentation & error handling
'   BLC - 3/16/2016 - deprecated but not removed - the park should now always be selected since plot is disabled until
'                     park is chosen, as a result the park code "You must select a park first." should never be called
'   AZ  - 3/26/2022 - adjusted code to new name for combo boxes; no longer deprecated
' ---------------------------------

Private Sub cbxPlot_GotFocus()
On Error GoTo Err_Handler

  If IsNull(Me!cbxPark) Then
    MsgBox "You must select a park first."
    Me!cbxPark.SetFocus
  End If
  
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Plot_GotFocus[frm_Render_Revisit_Reports])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxPark_Change (formerly Park_Code_Change)
' Description:  Contains actions occuring after changing park code
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       thePark = Me!cbxPark: invalid use of null; when using keyboard to change value
' References:   none
' Source/date:  Bonnie Campbell, March 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/16/2016 - initial version
'   AZ  - 3/23/2022 - changed query for plot drop down to include only plots in tbl_Revisit_List
'   AZ  - 3/26/2022 - adjusted code to new names for combo boxes
'   AZ  - 3/28/2022 - comment out code to create new rowsource; throws error when using keyboard
' ---------------------------------

Private Sub cbxPark_Change()
On Error GoTo Err_Handler

    Dim thePark As String
    Dim strSQL As String

'   thePark = Me!cbxPark
    strSQL = "SELECT tbl_Revisit_List.Plot FROM tbl_Revisit_List " & _
        "WHERE (((tbl_Revisit_List.Park) = '" & thePark & "' )) ORDER BY tbl_Revisit_List.Plot;"
        
   'plot dropdown should be disabled (default)
   Me!cbxPlot.Enabled = False
   Me.Refresh
  
  If Not IsNull(Me!cbxPark) Then
'    Me!cbxPlot.RowSource = "SELECT Plot_ID FROM tbl_Locations WHERE Unit_Code = '" & Me!Park_Code & "' ORDER BY Plot_ID"
'    Me!cbxPlot.RowSource = strSQL
'    Me!cbxPlot.Requery
    'enable plot dropdown
    Me!cbxPlot.Enabled = True
    Me!cbxPlot = Null
  End If
  
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Park_Code_Change[frm_Render_Revisit_Reports])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxPlot_AfterUpdate  (formerly Park_Code_AfterUpdate)
' Description:  Contains actions occuring after updating park code
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Russ DenBleyker, unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   RD - ?          - initial version
'   BLC - 3/08/2016  - added documentation & error handling
'   BLC - 3/16/2016 - enabled plot dropdown when park code is selected (otherwise it's disabled)
'   AZ  - 3/23/2022 - changed query for plot drop down to include only plots in tbl_Revisit_List
'   AZ  - 3/26/2022 - adjusted code to new names for combo boxes
'                   - added code to clear contents of plot combo box (not the rowsource)
' ---------------------------------
Private Sub cbxPark_AfterUpdate()
On Error GoTo Err_Handler

Dim strSQL As String

strSQL = "SELECT tbl_Revisit_List.Plot FROM tbl_Revisit_List " & _
        "WHERE (((tbl_Revisit_List.Park) = '" & Me!cbxPark & "' )) ORDER BY tbl_Revisit_List.Plot;"
 
  If Not IsNull(Me!cbxPark) Then
    'Me!Plot.RowSource = "SELECT Plot_ID FROM tbl_Locations WHERE Unit_Code = '" & Me!Park_Code & "' ORDER BY Plot_ID"
    Me!cbxPlot.RowSource = strSQL
    Me!cbxPlot.Requery
    'enable plot dropdown
    Me!cbxPlot.Enabled = True
  End If
  
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Park_Code_AfterUpdate[frm_Render_Revisit_Reports])"
    End Select
    Resume Exit_Handler
    
End Sub

' ---------------------------------
' SUB:          Button_Close_Click
' Description:  Contains actions occuring when close button is clicked
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Russ DenBleyker, unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   RD - ?          - initial version
'   BLC - 3/8/2016  - added documentation & error handling
' ---------------------------------
Private Sub Button_Close_Click()
On Error GoTo Err_Handler

    DoCmd.Close

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Button_Close_Click[frm_Render_Revisit_Reports])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Function:     Get_Reports_Setup
' Description:  preliminary checks; creates new species table (which underlies report)
' Assumptions:  Title of the database = opening arguments supplied by Forms!frm_Switchboard.btnQA_Click()
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  AZ, March 2022
' Adapted:      Allen Browne -https://bytes.com/topic/access/answers/208905-how-determine-if-directory-exists
' Revisions:
'   AZ  - 3/25/2022  - initial version
' ---------------------------------
Private Function Get_Reports_Setup()
On Error GoTo Err_Handler

    Debug.Print dBtitle
    'MsgBox dBtitle
             
    Dim docpath As String
    Dim FolderExists As Boolean
    Dim TemplatesExist As Boolean
    Dim response As Integer
    Debug.Print vbNewLine & "begin processing..."
    
    If Not (Me!chkOverview) And Not Me!chkSpecies And Not Me!chkTrees Then
        MsgBox "You must select a report type.", vbOKOnly
        GoTo Exit_Handler
    End If
    
    'check for presence of templates
    If Me!chkOverview Then
        TemplatesExist = True
         
        docpath = Application.CurrentProject.Path & "\Plot_Establishment.dot"
        FolderExists = (Len(Dir$(docpath, vbDirectory)) > 0&)
        If Not FolderExists Then
            MsgBox "You are missing 'Plot_Establishment.dot'", vbCritical
            TemplatesExist = False
        End If
        
        docpath = Application.CurrentProject.Path & "\Plot_Establishment2.dot"
        FolderExists = (Len(Dir$(docpath, vbDirectory)) > 0&)
        If Not FolderExists Then
            MsgBox "You are missing 'Plot_Establishment2.dot'", vbCritical
            TemplatesExist = False
        End If
        
        If Not (TemplatesExist) Then
            MsgBox "You cannont generate plot revisit overview reports without the templates." _
                & vbNewLine & "Please place these documents in the same folder as the front end" _
                & " or deselect PlotRevisitReport", vbOK
            GoTo Exit_Handler
        End If
    End If
    
    'verify user wants to generate species table
    If Me!chkSpecies And Not IsNull(Me!cbxPark) Then
        response = MsgBox("Generating the species table will take about 2 minutes." _
                   & vbNewLine & "Do you wish to continue?", vbOKCancel)
        If response = vbOK Then Call Generate_Species_Table
    End If
    
    Debug.Print Me!cbxPark & "-" & Me!cbxPlot
    
    Dim r As Integer
    Randomize
    r = Int(100 * Rnd)
    If r = 33 Then Call zDoesItWork(vbYes)
    If r = 66 Then Call zDoesItWork(vbNo)
    
    'determine if a park or park & plot are selected and route accordingly
    If Not IsNull(Me!cbxPark) And Not IsNull(Me!cbxPlot) Then
        Debug.Print "both park and plot"
        DoCmd.OpenForm "frm_Report_Rendering_Updates", OpenArgs:=101 & "|" & 0
        Call Get_Reportz(Me!cbxPark, Me!cbxPlot)
        AppActivate dBtitle
        DoCmd.Close acForm, "frm_Report_Rendering_Updates"
        MsgBox "C'EST TOUT."
    End If
    
    If Not IsNull(Me!cbxPark) And IsNull(Me!cbxPlot) Then Call sub_Loop_Plots
    If IsNull(Me!cbxPark) And IsNull(Me!cbxPlot) Then MsgBox "A Park is currently required."
 
Exit_Handler:
    Screen.MousePointer = 1
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Get_Reports_Setup[frm_Render_Revisit_Reports])"
    End Select
    Resume Exit_Handler
    
End Function
' ---------------------------------
' SUB:          Get_Reportz
' Description:  Directory management. Calls appropriate procedures to generate reports.
' Assumptions:
' Parameters:   leparc (park), laplacette (plot)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  AZ, March 2022
' Adapted:      Allen Browne -https://bytes.com/topic/access/answers/208905-how-determine-if-directory-exists
' Revisions:
'   AZ  - 3/25/2022  - initial version
' ---------------------------------
Private Sub Get_Reportz(leparc As String, laplacette As Integer)
On Error GoTo Err_Handler
    
    'check for directories and create if they don't exist
    'then call procedure to generate reports
    'called by Get_Reports_Setup if park & plot selected(above); otherwise by looping subs(below)
     
    Dim spath As String
    Dim FolderExists As Boolean
    
    'Plot Overviews----------------------------------------------------------------------
    If Me!chkOverview Then
        spath = Application.CurrentProject.Path & "\RevisitReports\" & leparc & "\Revisit_Reports"
        FolderExists = (Len(Dir$(spath, vbDirectory)) > 0&)
        If Not FolderExists Then
            Call MkMyDir(spath)
        End If
        Call getPlotInfo(leparc, laplacette, spath)  'call procedure to generate overview reports
    End If
    
    'Overstory Census--------------------------------------------------------------------
    If Me!chkTrees Then
        Debug.Print "getting tree report"
                
        spath = Application.CurrentProject.Path & "\RevisitReports\" & leparc & "\OTcensus"
        FolderExists = (Len(Dir$(spath, vbDirectory)) > 0&)
        If Not FolderExists Then
            Call MkMyDir(spath)
        End If
        Call getTrees(leparc, laplacette, spath)     'call procedure to generate overstory reports
    End If
     
    'Species Presence--------------------------------------------------------------------
    If Me!chkSpecies Then
        spath = Application.CurrentProject.Path & "\RevisitReports\" & leparc & "\Species"
        FolderExists = (Len(Dir$(spath, vbDirectory)) > 0&)
        If Not FolderExists Then
            Call MkMyDir(spath)
        End If
        Call getSpecies(leparc, laplacette, spath)   'call procedure to generate species reports
    End If
    
Exit_Handler:
    Screen.MousePointer = 1
    Exit Sub

Err_Handler:
      Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetReportz[frm_Render_Revisit_Report])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          sub_Loop_Parks
' Description:  Loop through all parks (units).
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  AZ, March 2022
' Adapted:      -
' Revisions:
'   AZ  - 3/25/2022  - under construction
' ---------------------------------

Private Sub sub_Loop_Parks()
On Error GoTo Err_Handler
    'under construction
    
Exit_Handler:
    Exit Sub

Err_Handler:
      Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - sub_Loop_Parks[frm_Render_Revisit_Report])"
    End Select
    Resume Exit_Handler
    
End Sub

' SUB:          sub_Loop_Plots
' Description:  Loops through plots; calls Get_Reportz to generate reports for each plot
' Assumptions:  Title of the database = opening arguments supplied by Forms!frm_Switchboard.btnQA_Click()-
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  AZ, March 2022
' Adapted:      -
' Revisions:
'   AZ  - 3/25/2022  - initial version
' ---------------------------------

Private Sub sub_Loop_Plots()
On Error GoTo Err_Handler

    Debug.Print "i am looping through the plots"
        
    Dim thePark As String
    Dim thePlot As Integer
    Dim r As Integer
    Dim strSQL As String
    Dim rs As dao.Recordset
    Dim response As Integer
    Dim CountOfPlots, ProgressNumber As Integer
    Dim StartTimeTotal As Double
    Dim SecondsElapsedTotal As Double
    Dim CurrentTime As Double
    Dim ResponseToAbortOpportunity As Integer
        
    StartTimeTotal = Timer
    i = 1
    thePark = Forms!frm_rev_Render_Revisit_Reports!cbxPark
    
    strSQL = "SELECT tbl_Revisit_List.Plot FROM tbl_Revisit_List " & _
        "WHERE (((tbl_Revisit_List.Park) = '" & thePark & "' ))ORDER BY tbl_Revisit_List.Plot;"

    Set rs = CurrentDb.OpenRecordset(strSQL)
    
    If Not rs.BOF And Not rs.EOF Then
        'get total number of plots & combine with current plot for form
        rs.MoveLast
        CountOfPlots = rs.RecordCount
                
        rs.MoveFirst
        While (Not rs.EOF)
            'set up form to notify user of progress
            ProgressNumber = (CountOfPlots * 100) + i
            CurrentTime = Round((Timer - StartTimeTotal), 2)
            AppActivate dBtitle
            DoCmd.OpenForm "frm_Report_Rendering_Updates", OpenArgs:=ProgressNumber & "|" & CurrentTime
            
            'call procedure to generate reports
            thePlot = rs.Fields("Plot")
            Call Get_Reportz(thePark, thePlot)
           
            Debug.Print "Call reports . . ."
            
            DoCmd.Close acForm, "frm_Report_Rendering_Updates"
            
            If i = 1 Then   'allow abort after generating reports for one plot
                AppActivate dBtitle
                ResponseToAbortOpportunity = MsgBox("It took " & Round((Timer - StartTimeTotal), 2) _
                    & " seconds to complete 1 of " & CountOfPlots & " reports." & vbNewLine _
                    & "This is your last chance to abort!" & vbNewLine _
                    & "Do you wish to continue?", vbOKCancel)
                If ResponseToAbortOpportunity = vbCancel Then
                    rs.Close
                    GoTo Exit_Handler
                End If
            End If
            i = i + 1
            rs.MoveNext
        Wend
    End If
    rs.Close
    Set rs = Nothing
    
    SecondsElapsedTotal = Round(Timer - StartTimeTotal, 3)
    
    AppActivate dBtitle
    MsgBox "C'EST TOUT." & vbNewLine & vbNewLine & _
        CountOfPlots & " reports rendered in " & SecondsElapsedTotal & " seconds (" _
        & Round(SecondsElapsedTotal / 60, 2) & " minutes).", vbOKOnly, Title:="AllDone"
    
Exit_Handler:
    Exit Sub

Err_Handler:
      Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - sub_Loop_Plots[frm_Render_Revisit_Reports])"
    End Select
    Resume Exit_Handler
    
End Sub
    
' ---------------------------------
' SUB:          Generate_Species_Table  (formerly Button_rpt_by_Park_Click)
' Description:  Button click actions (NO!)
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Juan Soto, October 31, 2011
'   https://accessexperts.com/blog/2011/10/31/more-alter-table-sql-statement-help/
'   Aiken, December 3, 2014
'   http://stackoverflow.com/questions/19369132/declare-and-initialize-string-array-in-vba
' Source/date:  Russ DenBleyker, unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   RD - ?          - initial version
'   BLC - 8/21/2015 - added update for underlying qry's table (temp_Sp_Rpt_by_Park_Complete)
'   BLC - 9/1/2015  - updated comment reference since underlying query qry_Sp_Rpt_by_Park_Rollup was
'                     replaced by the table temp_Sp_Rpt_by_Park_Rollup, added temporary revision
'                     of qry_Sp_Rpt_by_Park_Complete_Create_Table for filtered data (added appropriate
'                     WHERE clause) & reverted to original qdf SQL after running (for next time)
'   BLC - 3/7/2016  - fixed error where strSQL was not being replaced w/ strSQLNew causing report
'                     to ignore change in WHERE clause/prior data & not regenerate the proper report data set
'   BLC - 3/8/2016  - fixed issue where old data wasn't being refreshed in temp_Sp_Rpt_by_Park_Rollup
'                     by running qry_Sp_Rpt_by_Park_Rollup_Create_Table to refresh data from current back-end
'   BLC - 3/9/2016  - added multi-column indices to speed _Rollup query & report generation
'                     indexes performed slightly better than adding multicolum/autogenerated ID primary keys to
'                     temp_Sp_Rpt_by_Park_Complete & temp_Sp_Rpt_by_Park_Rollup tables
'                     report generation time dropped from ~5min -> ~48sec for 18212 results across all parks/years,
'                     added additional query/table statusbar messaging
'                     n.b. status bar messages are superceded by app Run Query, etc. (fix later)
'   BLC - 3/16/2016 - fix issue resulting in parameter error #3474 data type mismatch in criteria expression
'                     due to where clause in existing qry_Sp_Rpt_by_Park_Complete_Create_Table which was:
'                     WHERE Len(SpeciesYears) > Len(Replace(SpeciesYears, CStr(2014), ''))
'                     remove it and leave existing ORDER BY clause
'                     added report filter display via OpenArgs
'   AZ  - 3/26/2022 - Replaced qry_Sp_Rpt_All with qry_Sp_Rpt_All_Revisits, which limits query to current year plots.
'                   - Adjusted code to new names for combo boxes
'                   - Decluttered: Deleted all status bar messages which aren't working anyway
'                                  Deleted all old, commented-out code; you can still find it in the original form
'                   - Added code to close the report; this produces one report which includes all plots in a park
' ---------------------------------
Private Sub Generate_Species_Table()
On Error GoTo Err_Handler

    Screen.MousePointer = 11 'Hour Glass
    
    Dim strFilter As String, strWhere As String, strParkWhere As String, strPlotWhere As String, strYrWhere As String, strSpeciesYear As String
    Dim stDocName As String

    'defaults
    strFilter = ""
    strWhere = ""
    strParkWhere = ""
    strPlotWhere = ""
    strYrWhere = ""

    stDocName = "rpt_Species_by_Park"
    
    ' Set where condition if needed
    If (IsNull(Me!cbxPark) + IsNull(Me!cbxYear) + IsNull(Me!cbxPlot)) > -3 Then
      
      'park
      If Not IsNull(Me!cbxPark) Then
        strParkWhere = "Unit_Code = '" & Me!cbxPark & "'"
        strFilter = Me!cbxPark
      End If
      
      'plot --> NOTE: assumes UI will not allow plot selection w/o park
      If Not IsNull(Me!cbxPlot) Then
        strPlotWhere = "Plot_ID = " & Me!cbxPlot
        strFilter = strFilter & "- plot #" & Me!cbxPlot
      End If
      
      'year
      If Not IsNull(Me!cbxYear) Then
        strSpeciesYear = "(qry_Sp_Rpt_All_Revisit.Utah_Species+' - '+CStr(qry_Sp_Rpt_All_Revisit.Year))"
        strYrWhere = "Len(" & strSpeciesYear & ") > Len(Replace(" & strSpeciesYear & ", CStr(" & Me!cbxYear & "), ''))"
        
        'set filter display
        Select Case Len(strFilter)
            Case 0 'year only
                strFilter = CStr(Me!cbxYear)
            Case 4 'park only
                strFilter = strFilter & "-" & CStr(Me!cbxYear)
            Case Is > 4 'park & plot
                strFilter = Replace(strFilter, "-", "-" & CStr(Me!cbxYear) & " ")
        End Select
      
      Else
        'clear extra "-" for park & plot filter
        strFilter = Replace(strFilter, "-", "")
      End If
      
      'prepare where using string array & PrepareWhereClause
      Dim ary() As String
      ary = Split(strParkWhere & ";" & strPlotWhere & ";" & strYrWhere, ";")
      strWhere = PrepareWhereClause(ary)
      
    End If      'for setting where conditions
    
    'retrieve querydef
    Dim qdf As QueryDef
    Dim strSQL As String
    
    Set qdf = CurrentDb.QueryDefs("qry_Sp_Rpt_by_Park_Complete_Create_Table")
    strSQL = qdf.SQL

    'update the SQL if parameters exist
    If Len(strWhere) > 0 Then
        Dim iOrderBy As Integer
        Dim strSQLNew As String
        
        'replace ORDER with WHERE clause + ORDER
        strSQLNew = Replace(strSQL, "ORDER", " WHERE " & strWhere & " ORDER")
        qdf.SQL = strSQLNew 'was strSQL
    End If
    
    'update underlying table (temp_Sp_Rpt_by_Park_Complete is used in report's underlying table temp_Sp_Rpt_by_Park_Rollup)
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "qry_Sp_Rpt_by_Park_Complete_Create_Table", acViewNormal
    
    'add an index to improve report performance
    Dim strIdxSQL As String
    
    strIdxSQL = "CREATE INDEX idxParkPlotSpeciesYear ON temp_Sp_Rpt_by_Park_Complete (ParkPlotSpecies, Year)"
    CurrentDb.Execute strIdxSQL
    
    DoCmd.SetWarnings True
    
    'reset qdf SQL
    qdf.SQL = strSQL
    
    'update underlying table (temp_Sp_Rpt_by_Park_Rollup)
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "qry_Sp_Rpt_by_Park_Rollup_Create_Table", acViewNormal
    
   'add an index to improve report performance
    strIdxSQL = "CREATE INDEX idxParkPlotSpeciesYears ON temp_Sp_Rpt_by_Park_Rollup (ParkPlotSpecies, SpeciesYears)"
    CurrentDb.Execute strIdxSQL
    
    DoCmd.SetWarnings True
    
    'translate SQL Where for rollup --> SpeciesYear = SpeciesYears, ,qry_Sp_Rpt_All.Year = SpeciesYears, qry_Sp_Rpt_All.Utah_species = "Utah.species"
    Dim aryText() As String
    aryText = Split("SpeciesYear|SpeciesYears||qry_Sp_Rpt_All_Revisit.Year|SpeciesYears||qry_Sp_Rpt_All_Revisit.Utah_species|Utah_species", "||")
    strWhere = ReplaceMulti(strWhere, aryText)
    
    'open report --> strWhere = WHERE clause filter, strFilter = display for filter if present
    DoCmd.OpenReport stDocName, acViewPreview, , strWhere, acWindowNormal, strFilter
    DoCmd.Close acReport, stDocName
    
    Screen.MousePointer = 1 'boring old arrow
     
Exit_Handler:
    Set qdf = Nothing
    Screen.MousePointer = 1
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Generate_Species_Table[frm_Render_Revisit_Reports])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          getPlotInfo   (formerly Button_Print_Click)
' Description:  Generate Plot Establishment sheets for selected park
' Assumptions:
' Parameters:   prkCode As String, pltNum As Integer, spath As String
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Russ DenBleyker, April, 2008
' Adapted:      -
' Revisions:
'   Russ DenBleyker - Northern Colorado Plateau Network April, 2008.
'   Added Forest Woodlands June, 2009.
'   Added revisit comments, March, 2010.
'   -----------------------------------------
'   HT -  3/24/2015 - MArray adjustments to include missing monument tree entries for T2
'   -----------------------------------------
'   BLC - 8/10/2015 - specified recordset objects as ADODB.Recordset to avoid
'                     compiler errors "method or object not found" on rst/Mrst/RevisitComments.Open
' -------------------------------------------
'   AZ -  3/23/2022 - added code to suppress monument tree array if there are no monument trees
'                   - deleted code checking for null combo boxes (now done in the calling procedures)
'                   - added code to save file in appropriate directory & close
'                   - also deleted one line of commented-out dlookup code
'--------------------------------------------

Private Sub getPlotInfo(prkCode As String, pltNum As Integer, spath As String)
   On Error GoTo Err_Handler
    
    Dim objWord As Word.Application
    Dim fld As field
    Dim rst As ADODB.Recordset
    Dim Mrst As ADODB.Recordset
    Dim cat As ADOX.Catalog
    Dim tbl_Work_Surface_Type As ADOX.table
    Dim RevisitComments As ADODB.Recordset
    Dim strSQL As String
    ' Dim MArray(12) As String     Monument tree entries for T2 were inadvertently omitted from initial version of database.
    Dim MArray(18) As String     ' MArray was modified to accommodate these entries. [HT, 3/24/2015]
    Dim intIndex As Integer
    Dim strFieldName As String
    Dim intResponse As Integer
    Dim strPrompt As String
    Dim MonumentsExist As Boolean
    
   'call to function to determine if current plot has any monument trees
    MonumentsExist = MonTrees(prkCode, pltNum)
       
If MonumentsExist Then    'If-Then #1
  ' Initialize monument tree array if there are any monument trees
    MArray(0) = "T1-Oa"
    MArray(1) = "T1-Ob"
    MArray(2) = "T1-Oc"
    MArray(3) = "T1-Ea"
    MArray(4) = "T1-Eb"
    MArray(5) = "T1-Ec"
  ' Modified to accommodate T2 monument tree entries. [HT, 3/24/2015]
  ' MArray(6) = "T3-Oa"
  ' MArray(7) = "T3-Ob"
  ' MArray(8) = "T3-Oc"
  ' MArray(9) = "T3-Ea"
  ' MArray(10) = "T3-Eb"
  ' MArray(11) = "T3-Ec"
    MArray(6) = "T2-Oa"
    MArray(7) = "T2-Ob"
    MArray(8) = "T2-Oc"
    MArray(9) = "T2-Ea"
    MArray(10) = "T2-Eb"
    MArray(11) = "T2-Ec"
    MArray(12) = "T3-Oa"
    MArray(13) = "T3-Ob"
    MArray(14) = "T3-Oc"
    MArray(15) = "T3-Ea"
    MArray(16) = "T3-Eb"
    MArray(17) = "T3-Ec"
End If          'end if  #1

   'Launch Word and load the report template
    Set objWord = New Word.Application
   'Set objWord = GetObject(, "Word.Application")   wanting to not reopen application for every report
    
    If MonumentsExist Then
        objWord.Documents.Add _
        Application.CurrentProject.Path & "\Plot_Establishment.dot"
    Else
        objWord.Documents.Add _
        Application.CurrentProject.Path & "\Plot_Establishment2.dot"
    End If
    
    objWord.Visible = True
    
  'Build main SQL string for details
    strSQL = "SELECT * FROM tbl_locations WHERE [Unit_Code] = '" & prkCode & "' AND [Plot_ID] = " & pltNum
    
  'Get the database record.
    Set rst = New ADODB.Recordset
    rst.Open strSQL, CurrentProject.Connection
    If rst.EOF Then
      MsgBox "Record not found."
      GoTo Exit_Handler
    End If

    rst.MoveFirst

   'Update Word template
With objWord.ActiveDocument.Bookmarks
      If Not IsNull(rst.Fields("Unit_Code")) Then
        .Item("Unit_Code").Range.text = rst.Fields("Unit_Code")
      End If
      If Not IsNull(rst.Fields("Plot_ID")) Then
        .Item("Plot_ID").Range.text = rst.Fields("Plot_ID")
      End If
      If Not IsNull(rst.Fields("E_Coord")) Then
        .Item("E_Coord").Range.text = rst.Fields("E_Coord")
      End If
      If Not IsNull(rst.Fields("N_Coord")) Then
        .Item("N_Coord").Range.text = rst.Fields("N_Coord")
      End If
      If Not IsNull(rst.Fields("UTM_Zone")) Then
        .Item("UTM_Zone").Range.text = rst.Fields("UTM_Zone")
      End If
      If Not IsNull(rst.Fields("Plot_Slope")) Then
        .Item("Plot_Slope").Range.text = " " & rst.Fields("Plot_Slope")
      End If
      If Not IsNull(rst.Fields("Plot_Aspect")) Then
        .Item("Plot_Aspect").Range.text = " " & rst.Fields("Plot_Aspect")
      End If
      If Not IsNull(rst.Fields("Datum")) Then
        .Item("Datum").Range.text = rst.Fields("Datum")
      End If
      If Not IsNull(rst.Fields("Azimuth")) Then
        .Item("Azimuth").Range.text = " " & rst.Fields("Azimuth")
      End If
      If Not IsNull(rst.Fields("T1O_UTME")) Then
        .Item("T1O_UTME").Range.text = " " & rst.Fields("T1O_UTME")
      End If
      If Not IsNull(rst.Fields("T1O_UTMN")) Then
        .Item("T1O_UTMN").Range.text = " " & rst.Fields("T1O_UTMN")
      End If
      If Not IsNull(rst.Fields("T2O_UTME")) Then
        .Item("T2O_UTME").Range.text = " " & rst.Fields("T2O_UTME")
      End If
      If Not IsNull(rst.Fields("T2O_UTMN")) Then
        .Item("T2O_UTMN").Range.text = " " & rst.Fields("T2O_UTMN")
      End If
      If Not IsNull(rst.Fields("T3O_UTME")) Then
        .Item("T3O_UTME").Range.text = " " & rst.Fields("T3O_UTME")
      End If
      If Not IsNull(rst.Fields("T3O_UTMN")) Then
        .Item("T3O_UTMN").Range.text = " " & rst.Fields("T3O_UTMN")
      End If
      If Not IsNull(rst.Fields("T1E_UTME")) Then
        .Item("T1E_UTME").Range.text = " " & rst.Fields("T1E_UTME")
      End If
      If Not IsNull(rst.Fields("T1E_UTMN")) Then
        .Item("T1E_UTMN").Range.text = " " & rst.Fields("T1E_UTMN")
      End If
      If Not IsNull(rst.Fields("T2E_UTME")) Then
        .Item("T2E_UTME").Range.text = " " & rst.Fields("T2E_UTME")
      End If
      If Not IsNull(rst.Fields("T2E_UTMN")) Then
        .Item("T2E_UTMN").Range.text = " " & rst.Fields("T2E_UTMN")
      End If
      If Not IsNull(rst.Fields("T3E_UTME")) Then
        .Item("T3E_UTME").Range.text = " " & rst.Fields("T3E_UTME")
      End If
      If Not IsNull(rst.Fields("T3E_UTMN")) Then
        .Item("T3E_UTMN").Range.text = " " & rst.Fields("T3E_UTMN")
      End If
      If Not IsNull(rst.Fields("T1O_Rebar")) Then
        .Item("T1O_Rebar").Range.text = " " & rst.Fields("T1O_Rebar")
      End If
      If Not IsNull(rst.Fields("T1E_Rebar")) Then
        .Item("T1E_Rebar").Range.text = " " & rst.Fields("T1E_Rebar")
      End If
      If Not IsNull(rst.Fields("T2O_Rebar")) Then
        .Item("T2O_Rebar").Range.text = " " & rst.Fields("T2O_Rebar")
      End If
      If Not IsNull(rst.Fields("T2E_Rebar")) Then
        .Item("T2E_Rebar").Range.text = " " & rst.Fields("T2E_Rebar")
      End If
      If Not IsNull(rst.Fields("T3O_Rebar")) Then
        .Item("T3O_Rebar").Range.text = " " & rst.Fields("T3O_Rebar")
      End If
      If Not IsNull(rst.Fields("T3E_Rebar")) Then
        .Item("T3E_Rebar").Range.text = " " & rst.Fields("T3E_Rebar")
      End If
      If Not IsNull(rst.Fields("T1_Elevation")) Then
        .Item("T1_Elevation").Range.text = " " & rst.Fields("T1_Elevation")
      End If
      If Not IsNull(rst.Fields("T2_Elevation")) Then
        .Item("T2_Elevation").Range.text = " " & rst.Fields("T2_Elevation")
      End If
      If Not IsNull(rst.Fields("T3_Elevation")) Then
        .Item("T3_Elevation").Range.text = " " & rst.Fields("T3_Elevation")
      End If
      If Not IsNull(rst.Fields("Plot_Directions")) Then
        .Item("Plot_Directions").Range.text = " " & rst.Fields("Plot_Directions")
      End If
      
   If rst.Fields("Vegetation_Type") <> "grassland/shrubland" Then        'if-then #2
         If Not IsNull(rst.Fields("SlopeA")) Then
           .Item("SlopeA").Range.text = " " & rst.Fields("SlopeA")
         End If
         If Not IsNull(rst.Fields("SlopeAUD")) Then
           .Item("SlopeAUD").Range.text = " " & rst.Fields("SlopeAUD")
         End If
         If Not IsNull(rst.Fields("SlopeB")) Then
           .Item("SlopeB").Range.text = " " & rst.Fields("SlopeB")
         End If
         If Not IsNull(rst.Fields("SlopeBUD")) Then
           .Item("SlopeBUD").Range.text = " " & rst.Fields("SlopeBUD")
         End If
         If Not IsNull(rst.Fields("SlopeC")) Then
           .Item("SlopeC").Range.text = " " & rst.Fields("SlopeC")
         End If
         If Not IsNull(rst.Fields("SlopeCUD")) Then
           .Item("SlopeCUD").Range.text = " " & rst.Fields("SlopeCUD")
         End If
         If Not IsNull(rst.Fields("SlopeD")) Then
           .Item("SlopeD").Range.text = " " & rst.Fields("SlopeD")
         End If
         If Not IsNull(rst.Fields("SlopeDUD")) Then
           .Item("SlopeDUD").Range.text = " " & rst.Fields("SlopeDUD")
         End If
'start monument trees
      If MonumentsExist Then        'if-then #3
         strSQL = "Select * from tbl_Monument WHERE Location_ID = '" & rst.Fields("Location_ID") & "'"
         ' Get the monument tree records.
         Set Mrst = New ADODB.Recordset
         Mrst.Open strSQL, CurrentProject.Connection
         If Not Mrst.EOF Then      'if-then #4
            Mrst.MoveFirst
            Do Until Mrst.EOF
              intIndex = 0   ' Initialize index
              ' Modified to accommodate monument tree entries for T2. [HT, 3/24/2015]
              ' Do Until intIndex > 11
              Do Until intIndex > 17
                If MArray(intIndex) = Mrst.Fields("Monument_Code") Then
                  Exit Do
                End If
                intIndex = intIndex + 1
              Loop
              ' Modified to accommodate monument tree entries for T2. [HT, 3/24/2015]
              ' If intIndex > 11 Then
              If intIndex > 17 Then
                Exit Do  ' Code not found - skip it
              End If
              intIndex = intIndex + 1  ' set index to correct subscript for bookmarks
              If Not IsNull(Mrst.Fields("Tag_No")) Then
                strFieldName = "Tag_No" & intIndex  ' set bookmark name
                .Item(strFieldName).Range.text = " " & Mrst.Fields("Tag_No")
              End If
              If Not IsNull(Mrst.Fields("Species")) Then
                strFieldName = "Species" & intIndex  ' set bookmark name
                .Item(strFieldName).Range.text = " " & DLookup("[LU_Code]", "tlu_NCPN_Plants", "[Master_Plant_Code] = '" & Mrst.Fields("Species") & "'")
              End If
              If Not IsNull(Mrst.Fields("DBH")) Then
                strFieldName = "DBH" & intIndex  ' set bookmark name
                .Item(strFieldName).Range.text = " " & Mrst.Fields("DBH")
              End If
              If Not IsNull(Mrst.Fields("Bearing")) Then
                strFieldName = "Bearing" & intIndex  ' set bookmark name
                .Item(strFieldName).Range.text = " " & Mrst.Fields("Bearing")
              End If
              If Not IsNull(Mrst.Fields("Rebar_Distance")) Then
                strFieldName = "Rebar_Distance" & intIndex  ' set bookmark name
                .Item(strFieldName).Range.text = " " & Mrst.Fields("Rebar_Distance")
              End If
              Mrst.MoveNext
            Loop
         End If    ' for mon tree recordset loop   end if #4
      End If  'for monument trees   end if #3
   End If  ' End if for forest/woodland compare      end if #2
   
     'Check for revisit comments and print if necessary
     'If Not IsNull(Me!Park_Code) And Not IsNull(Me!Plot_ID) Then
      strSQL = "SELECT * FROM qry_Visit_Comments WHERE [Unit_Code] = '" & prkCode & "' AND [Plot_ID] = " & pltNum & " ORDER BY [Start_Date] DESC"
        ' Get the database record.
      Set RevisitComments = New ADODB.Recordset
      RevisitComments.Open strSQL, CurrentProject.Connection
      If Not RevisitComments.EOF Then
         RevisitComments.MoveFirst
         If Not IsNull(RevisitComments!Comments) Then
       ' MsgBox below commented out because it often displayed behind Word document and so was easily missed, and was not really necessary (ht 8/19/2010).
       '     strPrompt = "There are revisit comments dated " & RevisitComments.Fields("Start_Date") & " for this plot." & Chr(13) & Chr(10) & "Do you want to print them?"
       '     intReply = MsgBox(strPrompt, vbYesNo, "Want revisit comments?")
       '     If intReply = vbYes Then
            .Item("Revisit_Comments").Range.text = " " & RevisitComments.Fields("Comments")
       '     End If
         End If ' End if for null comments test
      End If    ' End if for comments eof test
     'End If    ' End if for null unit or plot test
End With        '(objWord.ActiveDocument.Bookmarks)

    rst.Close    ' Close for main query
    
   'focus to Word; save & close document; exit Word
    objWord.Activate
    objWord.ActiveDocument.SaveAs2 filename:=spath & "\PlotInfo_" & prkCode & pltNum & ".docx"
    objWord.ActiveDocument.Close
    objWord.Quit
    
Exit_Handler:
    Exit Sub

Err_Handler:
    If Err.Description = "Could not open macro storage." Then
        MsgBox "ERROR! Open both word templates:" & vbNewLine & _
        "Plot_Establishment.dot and Plot_Establishment2.dot" & vbNewLine _
            & " and click to allow edits."
    Else
        MsgBox Err.Description
    End If
    Resume Exit_Handler
    
End Sub

' SUB:          getTrees
' Description:  Generates OTcensus report for one plot and saves it in the appropriate folder
' Assumptions:  the desired year is the most recent
' Parameters:   unit code, plot, directory path
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  AZ, March 2022
' Adapted:      -
' Revisions:
'   AZ  - 3/25/2022  - initial version
' ---------------------------------
Private Function getTrees(prkCode As String, pltNum As Integer, spath As String)
On Error GoTo Err_Handler

    Dim stDocName As String
    Dim stWhereCondition As String
    Dim maxyear As Variant
    Dim filename, nodataname As String
    Dim pltNum2 As Integer
    Dim PP As String
    
    pltNum2 = CInt(pltNum)
    PP = prkCode & pltNum
   
    If IsNull(DLookup("[ParkPlot]", "qaz_sel_OT_Census_Report_e", "[ParkPlot] = '" & PP & "'")) Then
         'MsgBox "is null; not on list; less room ok; calling new report"
         stDocName = "rpt_OT_Census_mini"
    Else
         'MsgBox "not null; it's on list; need more room; calling original report"
         stDocName = "rpt_OT_Census"
    End If
    
    'set criteria
   
    maxyear = GetYear(prkCode, pltNum)
    
    stWhereCondition = "[Unit_Code] = '" & prkCode & "' AND " _
                  & "[Plot_Id] = " & pltNum2 & " AND " _
                    & "[Visit_Year] = " & maxyear
   
    filename = spath & "\OTcensus_" & prkCode & pltNum & ".pdf"
    nodataname = spath & "\x_nodata_OTcensus_" & prkCode & CInt(pltNum) & ".pdf"
    
    'open report, save, and close
    DoCmd.OpenReport stDocName, acViewPreview, , stWhereCondition
    If Reports(stDocName).HasData Then
        DoCmd.OutputTo objecttype:=acOutputReport, outputformat:=acFormatPDF, outputfile:=filename
    Else
        DoCmd.OutputTo objecttype:=acOutputReport, outputformat:=acFormatPDF, outputfile:=nodataname
    End If
    DoCmd.Close acReport, stDocName
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case 2501
        MsgBox "Are you maybe trying to create a file that is sitting open in Acrobat/Adobe?"
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - getSpecies[frm_Render_Revisit_Reports])"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - getTrees[frm_Render_Revisit_Reports])"
    End Select
    Resume Exit_Handler

End Function

' SUB:          getSpecies
' Description:  Generates Species Present report for one plot and saves it in the appropriate folder
' Assumptions:  all years are desired
' Parameters:   unit code (park), plot, directory path
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  AZ, March 2022
' Adapted:      - Daniel Pineault - https://www.devhut.net/ms-access-print-individual-pdfs-of-a-report/
' Revisions:
'   AZ  - 3/25/2022  - initial version
' ---------------------------------
Private Sub getSpecies(prkCode As String, pltNum As Integer, spath As String)
On Error GoTo Err_Handler

    Dim stDocName As String
    Dim sprpt As Access.Report
    Dim spSQL As String
    Dim sprs As dao.Recordset
    Dim thePlot As Integer
    Dim filename, nodataname As String
       
    Debug.Print spath
    
    stDocName = "rpt_Species_by_Park"
    filename = spath & "\Species_" & prkCode & pltNum & ".pdf"
    nodataname = spath & "\x_nodata_Species_" & prkCode & pltNum & ".pdf"
    
    'open report
    DoCmd.OpenReport stDocName, acViewPreview
    
    'save it as a report object
    Set sprpt = Reports("rpt_Species_by_Park").Report
    
    '-------------------------------------------------------------
    'commented out code can be used to loop through a list of plots:
     'open a recordset with a criteria (the plots for the park in the revisit list)
    'spSQL = "SELECT tbl_Revisit_List.Plot FROM tbl_Revisit_List " & _
        "WHERE (((tbl_Revisit_List.Park) = '" & prkCode & "' ) AND " & _
            "((tbl_Revisit_List.Plot) = '" & pltNum & "')) ORDER BY tbl_Revisit_List.Plot;"
     'Set sprs = CurrentDb.OpenRecordset(spSQL)
     '-----------------------------------------------------------
     
     'create and apply a filter
     sprpt.Filter = "Plot_ID = " & pltNum
     sprpt.FilterOn = True
     DoEvents
     
     'check for data, save the result, and close
     If Reports(stDocName).HasData Then
         DoCmd.OutputTo acOutputReport, stDocName, acFormatPDF, filename
     Else
         DoCmd.OutputTo acOutputReport, stDocName, acFormatPDF, nodataname
         If i = 1 Then _
            MsgBox "No species data for plot " & pltNum & "." & vbNewLine _
            & "This could be because you didn't update the species table.", vbOKOnly
     End If
    
     DoCmd.Close acReport, stDocName
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 2501
        MsgBox "Are you maybe trying to create a file that is sitting open in Acrobat/Adobe?"
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - getSpecies[frm_Render_Revisit_Reports])"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - getSpecies[frm_Render_Revisit_Reports])"
    End Select
    Resume Exit_Handler

End Sub

' Function:     MonTrees
' Description:  Checks if there are any monument trees for specified plot.
' Assumptions:
' Parameters:   unit code, plot
' Returns:      Boolean: yes if there monument trees, no if not
' Throws:       none
' References:   none
' Source/date:  AZ, March 2022
' Adapted:      -
' Revisions:
'   AZ  - 3/25/2022  - initial version
' ---------------------------------
Private Function MonTrees(theprk As String, theplt As Integer) As Boolean
    On Error GoTo Err_Handler
    
    Dim rst As ADODB.Recordset
    Dim strSQL3 As String
    Dim PrkPlt As String
    
    PrkPlt = theprk & theplt
    MonTrees = False
    
    strSQL3 = "SELECT DISTINCT [Unit_Code] & [Plot_ID] AS ParkPlot " _
             & "FROM tbl_Locations INNER JOIN tbl_Monument ON " _
             & "tbl_Locations.Location_ID = tbl_Monument.Location_ID;"
 
    Set rst = New ADODB.Recordset
    rst.Open strSQL3, CurrentProject.Connection
    If Not (rst.EOF And rst.BOF) Then
        rst.MoveFirst
        Do Until rst.EOF = True
            If rst!ParkPlot = PrkPlt Then
                MonTrees = True
                GoTo Exit_Handler
            End If
            rst.MoveNext
        Loop
        
     Else
        MsgBox "There are no monument trees at any location?!?"
     End If
     
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - MonTrees[frm_Render_Revisit_Reports])"
    End Select
    Resume Exit_Handler
   
End Function

' Function:     GetYear
' Description:  Gets the most recent visit year for a given plot.
' Assumptions:
' Parameters:   unit code, plot
' Returns:      year as variant
' Throws:       none
' References:   none
' Source/date:  AZ, March 2022
' Adapted:      -
' Revisions:
'   AZ  - 3/25/2022  - initial version
' ---------------------------------
Private Function GetYear(theprk As String, theplt As Integer) As Variant
On Error GoTo Err_Handler

    Dim rst As ADODB.Recordset
    Dim strSQL2 As String
  
  'selects all years the given plot was visited; sorts in descending order
    strSQL2 = "SELECT DatePart('yyyy',tbl_events.Start_Date) AS Visit_Year " _
            & "FROM tbl_Locations " _
            & "LEFT JOIN tbl_Events ON tbl_Locations.Location_ID = tbl_Events.Location_ID " _
            & "WHERE (DatePart('yyyy',tbl_events.Start_Date) Is Not Null) " _
            & "AND [Unit_Code] = '" & theprk & "' " _
            & "AND [Plot_ID] = " & theplt & " " _
            & "ORDER BY Year([Start_Date]) DESC;"
    
   ' strSQL2 = "SELECT Year([Start_Date]) AS Visit_Year " _
    '        & "FROM tbl_Locations " _
     '       & "LEFT JOIN tbl_Events ON tbl_Locations.Location_ID = tbl_Events.Location_ID " _
      '      & "WHERE (((Year([Start_Date])) Is Not Null)) " _
       '     & "AND [Unit_Code] = '" & theprk & "' " _
        '    & "AND [Plot_ID] = " & theplt & " " _
         '   & "ORDER BY DatePart('yyyy',tbl_events.Start_Date) DESC;"
    
  'the first year will be the most recent
    Set rst = New ADODB.Recordset
    rst.Open strSQL2, CurrentProject.Connection
    rst.MoveFirst
    GetYear = rst.Fields.Item(0).Value
     
    rst.Close
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetYear[frm_Render_Revisit_Reports])"
    End Select
    Resume Exit_Handler

End Function

Private Sub btnGetItAll_Click()
    MsgBox "Sorry. That functionality is currently unavailable."
End Sub

Private Sub zDoesItWork(response As Integer) 'test sub; delete (az)
    If response = vbYes Then
        MsgBox "A chemist and his friend walked into a bar." & vbNewLine & "The chemist ordered H-2-O"
        MsgBox "His friend ordered H-2-O-2"
    Else: MsgBox "If you're not part of the solution, you're part of the precipitate."
    End If
End Sub
