Version =21
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5940
    DatasheetFontHeight =9
    ItemSuffix =29
    Left =9615
    Top =-10935
    Right =15555
    Bottom =-5010
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x9d2210c6b41ee340
    End
    Caption ="Select for Plot Revisit Data Sheet"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
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
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin Section
            Height =5940
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =3600
                    Left =2280
                    Top =1320
                    Width =840
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"8\""
                    Name ="Park_Code"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT tbl_Locations.Unit_Code, tlu_Parks.ParkName FROM tlu_Parks INNER"
                        " JOIN tbl_Locations ON tlu_Parks.ParkCode = tbl_Locations.Unit_Code ORDER BY tbl"
                        "_Locations.Unit_Code;"
                    ColumnWidths ="720;2880"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =960
                            Top =1320
                            Width =1260
                            Height =245
                            FontWeight =700
                            Name ="Select a Park_Label"
                            Caption ="Select a Park"
                            EventProcPrefix ="Select_a_Park_Label"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1237
                    Top =540
                    Width =2520
                    Height =330
                    FontSize =12
                    FontWeight =700
                    Name ="Label2"
                    Caption ="Plot Revisit Reports"
                    LayoutCachedLeft =1237
                    LayoutCachedTop =540
                    LayoutCachedWidth =3757
                    LayoutCachedHeight =870
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =67
                    Left =3960
                    Top =3180
                    Width =1320
                    Height =300
                    TabIndex =7
                    Name ="ButtonClose"
                    Caption ="&Close Form"
                    OnClick ="[Event Procedure]"
                    UnicodeAccessKey =67

                    LayoutCachedLeft =3960
                    LayoutCachedTop =3180
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =3480
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListRows =24
                    ListWidth =540
                    Left =2280
                    Top =1800
                    Width =840
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"3\";\"2\""
                    Name ="Plot_ID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Plot_ID FROM tbl_locations WHERE [Unit_Code] = 'ZION' ORDER BY Plot_ID"
                    ColumnWidths ="540"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1560
                            Top =1800
                            Width =645
                            Height =245
                            FontWeight =700
                            Name ="Plot ID_Label"
                            Caption ="Plot ID"
                            EventProcPrefix ="Plot_ID_Label"
                        End
                    End
                End
                Begin OptionGroup
                    SpecialEffect =0
                    BorderWidth =2
                    OverlapFlags =85
                    Left =3420
                    Top =1140
                    Width =2297
                    Height =1273
                    TabIndex =2
                    BorderColor =3487637
                    Name ="optPrintOptions"
                    DefaultValue ="3"
                    GridlineColor =10921638

                    LayoutCachedLeft =3420
                    LayoutCachedTop =1140
                    LayoutCachedWidth =5717
                    LayoutCachedHeight =2413
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    Begin
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            Left =3606
                            Top =1378
                            OptionValue =1
                            BorderColor =10921638
                            Name ="Option11"
                            GridlineColor =10921638

                            LayoutCachedLeft =3606
                            LayoutCachedTop =1378
                            LayoutCachedWidth =3866
                            LayoutCachedHeight =1618
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =3836
                                    Top =1350
                                    Width =1020
                                    Height =240
                                    Name ="Label12"
                                    Caption ="Print Preview"
                                    LayoutCachedLeft =3836
                                    LayoutCachedTop =1350
                                    LayoutCachedWidth =4856
                                    LayoutCachedHeight =1590
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            AccessKey =68
                            Left =3606
                            Top =1708
                            TabIndex =1
                            OptionValue =2
                            BorderColor =10921638
                            Name ="Option13"
                            UnicodeAccessKey =68
                            GridlineColor =10921638

                            LayoutCachedLeft =3606
                            LayoutCachedTop =1708
                            LayoutCachedWidth =3866
                            LayoutCachedHeight =1948
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =3836
                                    Top =1680
                                    Width =1695
                                    Height =240
                                    Name ="Label14"
                                    Caption ="Print to &Default Printer"
                                    LayoutCachedLeft =3836
                                    LayoutCachedTop =1680
                                    LayoutCachedWidth =5531
                                    LayoutCachedHeight =1920
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            AccessKey =70
                            Left =3606
                            Top =2038
                            TabIndex =2
                            OptionValue =3
                            BorderColor =10921638
                            Name ="Option15"
                            UnicodeAccessKey =70
                            GridlineColor =10921638

                            LayoutCachedLeft =3606
                            LayoutCachedTop =2038
                            LayoutCachedWidth =3866
                            LayoutCachedHeight =2278
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =3836
                                    Top =2010
                                    Width =885
                                    Height =240
                                    Name ="Label16"
                                    Caption ="Print to &File"
                                    LayoutCachedLeft =3836
                                    LayoutCachedTop =2010
                                    LayoutCachedWidth =4721
                                    LayoutCachedHeight =2250
                                End
                            End
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =80
                    Left =3960
                    Top =2700
                    Width =1320
                    Height =330
                    TabIndex =6
                    Name ="btnSpeciesForm"
                    Caption ="S&pecies Reports"
                    UnicodeAccessKey =112
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="frm_Species_Report_Select"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"btnSpeciesForm\" xmlns=\"http://schemas.microsoft.com/office"
                                "/accessservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Arg"
                                "ument Name=\"FormName\">frm_Speci"
                        End
                        Begin
                            Comment ="_AXL:es_Report_Select</Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =3960
                    LayoutCachedTop =2700
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =3030
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =83
                    Left =1560
                    Top =3120
                    Width =1560
                    Height =480
                    TabIndex =5
                    Name ="btnOverstoryForm"
                    Caption ="Generate One Over&Story Report"
                    ControlTipText ="currently opens old form for generating tree reports one-by-one"
                    UnicodeAccessKey =83
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="frm_Select_Overstory_Revisit"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"btnOverstoryForm\" xmlns=\"http://schemas.microsoft.com/offi"
                                "ce/accessservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><A"
                                "rgument Name=\"FormName\">frm_Sel"
                        End
                        Begin
                            Comment ="_AXL:ect_Overstory_Revisit</Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =1560
                    LayoutCachedTop =3120
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =3600
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =2
                    OverlapFlags =93
                    Left =1560
                    Top =2520
                    Width =1560
                    Height =480
                    BorderColor =3487637
                    Name ="Box21"
                    GridlineColor =10921638
                    LayoutCachedLeft =1560
                    LayoutCachedTop =2520
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =3000
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =2
                    OverlapFlags =93
                    Left =300
                    Top =2640
                    Width =900
                    Height =780
                    BorderColor =3487637
                    Name ="Box22"
                    GridlineColor =10921638
                    LayoutCachedLeft =300
                    LayoutCachedTop =2640
                    LayoutCachedWidth =1200
                    LayoutCachedHeight =3420
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin CommandButton
                    OverlapFlags =215
                    AccessKey =77
                    Left =300
                    Top =2640
                    Width =900
                    Height =780
                    TabIndex =3
                    Name ="btnGetReportz"
                    Caption ="Get &Multiple Reports"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="generates plot datasheets and tree reports (GetReportz)\015\012 for all current-"
                        "year plots in specified park"
                    UnicodeAccessKey =77

                    LayoutCachedLeft =300
                    LayoutCachedTop =2640
                    LayoutCachedWidth =1200
                    LayoutCachedHeight =3420
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    AccessKey =72
                    Left =1560
                    Top =2520
                    Width =1560
                    Height =480
                    TabIndex =4
                    Name ="btnOverviewReport"
                    Caption ="Generate One Plot Datas&heet"
                    OnClick ="=GetReport()"
                    ControlTipText ="one plot datasheet only"
                    UnicodeAccessKey =104

                    LayoutCachedLeft =1560
                    LayoutCachedTop =2520
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =3000
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =840
                    Top =4680
                    Width =4200
                    Height =480
                    FontSize =9
                    TabIndex =8
                    ForeColor =255
                    Name ="btnTooManyWordDocs"
                    Caption ="Help! I have a gazillion word documents open!!! Is there an easy way to close th"
                        "em??"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =840
                    LayoutCachedTop =4680
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =5160
                    HoverForeColor =255
                    PressedForeColor =255
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    Visible = NotDefault
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =49
                    Left =840
                    Top =5220
                    Width =4200
                    Height =600
                    Name ="lblKILL"
                    Caption ="Glad you asked.\015\012Try: taskkill /f /im winword.exe\015\012in the command pr"
                        "ompt."
                    FontName ="Lucida Console"
                    LayoutCachedLeft =840
                    LayoutCachedTop =5220
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =5820
                End
                Begin Label
                    OverlapFlags =85
                    Left =540
                    Top =3840
                    Width =4950
                    Height =630
                    Name ="Label28"
                    Caption ="Before clicking on anything, check that you have the following in the same folde"
                        "r as your front end:  Plot_Establishment.dot, Plot_Establishment2.dot, an empty "
                        "folder called RevisitReports."
                    LayoutCachedLeft =540
                    LayoutCachedTop =3840
                    LayoutCachedWidth =5490
                    LayoutCachedHeight =4470
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
Dim rpt(1 To 48) As Report

Private Sub btnGetReportz_Click()
 On Error GoTo Err_GetReportz
    
  '  Dim st As String
    Dim thePark As String
    Dim thePlot As Integer
    Dim r As Integer
    Dim strSQL As String
    Dim rs As dao.Recordset
    Dim response As Integer
    Dim i As Integer
    Dim CountOfPlots, ProgressNumber As Integer
    Dim StartTimeTotal As Double
    Dim SecondsElapsedTotal As Double
    Dim CurrentTime As Double
    Dim ResponseToAbortOpportunity
    
    
    If IsNull(Me!Park_Code) Then
      MsgBox "You must select a park!"
      GoTo Exit_GetReportz
    End If
    
    If Me!optPrintOptions = 1 Then
        response = MsgBox("The print preview option does not work for trees." & vbNewLine & _
                "Only plot datasheet reports will be generated.", vbOKCancel)
        If response = vbCancel Then GoTo Exit_GetReportz
    End If
       
    StartTimeTotal = Timer
    DoCmd.Hourglass True
    i = 1
    
    thePark = Forms!frm_Render_Revisit_Reports!Park_Code
    
    'Debug.Print thePark
    strSQL = "SELECT tbl_Revisit_List.Plot FROM tbl_Revisit_List " & _
        "WHERE (((tbl_Revisit_List.Park) = '" & thePark & "' ))ORDER BY tbl_Revisit_List.Plot;"
    'Debug.Print strSQL
    Set rs = CurrentDb.OpenRecordset(strSQL)
    
    If Not rs.BOF And Not rs.EOF Then
        'get total number of plots & combine with current plot for form
        rs.MoveLast
        CountOfPlots = rs.RecordCount
                
        rs.MoveFirst
        While (Not rs.EOF)
            ProgressNumber = (CountOfPlots * 100) + i
            CurrentTime = Round((Timer - StartTimeTotal), 2)
            AppActivate "edaz NCPN Upland Monitoring Database"
            DoCmd.OpenForm "frm_Report_Rendering_Updates", OpenArgs:=ProgressNumber & "|" & CurrentTime
            Debug.Print ProgressNumber & "\' & CurrentTime"
            thePlot = rs.Fields("Plot")
            Call ButtonPrint(thePark, thePlot)  'call function/sub to generate overview reports
            Call getTrees(thePark, thePlot)     'call function/sub to generate overstory reports
            
            DoCmd.Close acForm, "frm_Report_Rendering_Updates"
            If i = 1 Then
                AppActivate "edaz NCPN Upland Monitoring Database"
                ResponseToAbortOpportunity = MsgBox("This is your last chance to abort!" _
                        & vbNewLine & "Do you wish to continue?", vbOKCancel)
                If ResponseToAbortOpportunity = vbCancel Then
                    rs.Close
                    GoTo Exit_GetReportz
                End If
            End If
            i = i + 1
            rs.MoveNext
        Wend
    End If
    rs.Close
    Set rs = Nothing
    
    DoCmd.Hourglass False
    SecondsElapsedTotal = Round(Timer - StartTimeTotal, 2)
    
    r = Int(100 * Rnd)
    If r = 33 Then Call zDoesItWork(vbNo)    'call test sub
    If r = 66 Then Call zDoesItWork(vbYes)
    AppActivate "edaz NCPN Upland Monitoring Database"
    MsgBox "C'EST TOUT." & vbNewLine & vbNewLine & _
        CountOfPlots & " reports rendered in " & SecondsElapsedTotal & " seconds." & _
        vbNewLine & "(For print preview option, Word Docs are minimized in the system tray.)" _
        , vbOKOnly, Title:="AllDone"
    
    If Me.optPrintOptions = 1 Then Me.btnTooManyWordDocs.Visible = True
    
Exit_GetReportz:
    DoCmd.Hourglass False
    Exit Sub

Err_GetReportz:
    MsgBox Err.Description
    Resume Exit_GetReportz
End Sub

Private Sub btnTooManyWordDocs_Click()
    Me.lblKILL.Visible = True
End Sub

Private Sub ButtonClose_Click()
On Error GoTo Err_ButtonClose_Click
    
    Me.lblKILL.Visible = False
    Me.btnTooManyWordDocs.Visible = False
    DoCmd.Close

Exit_ButtonClose_Click:
    Exit Sub

Err_ButtonClose_Click:
    MsgBox Err.Description
    Resume Exit_ButtonClose_Click
    
End Sub

Private Sub Form_Close()
    'Me.lblKILL.Visible = False
    'Me.btnTooManyWordDocs.Visible = False
End Sub

Private Sub Park_Code_AfterUpdate()
'this populates list of plots for plot dropdown, based on value of park dropdown

  Me!Plot_ID = Null
  If Not IsNull(Me!Park_Code) Then
    Me!Plot_ID.RowSource = "SELECT Plot FROM tbl_Revisit_List WHERE [PARK] = '" & Me!Park_Code & "' ORDER BY Plot"
    Me!Plot_ID.Requery
  Else
    MsgBox "You must select a park!"
  End If
    
End Sub

' ---------------------------------
' SUB:          Button_Print_Click
' Description:  Generate Plot Establishment sheets for selected park
' Assumptions:
' Parameters:   -
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
'   HT - 3/24/2015  - MArray adjustments to include missing monument tree entries for T2
'   -----------------------------------------
'   BLC - 8/10/2015 - specified recordset objects as ADODB.Recordset to avoid
'                     compiler errors "method or object not found" on rst/Mrst/RevisitComments.Open
' ---------------------------------
'Private Sub Button_Print_Click()
Private Sub ButtonPrint(prkCode As String, pltNum As Integer)
On Error GoTo Err_Button_Print

    Dim objWord As Word.Application
    Dim fld As field
    Dim rst As ADODB.Recordset
    Dim Mrst As ADODB.Recordset
    Dim cat As ADOX.Catalog
    Dim tbl_Work_Surface_Type As ADOX.table
    Dim RevisitComments As ADODB.Recordset
    Dim strSQL As String
    ' Monument tree entries for T2 were inadvertently omitted from initial version of database.
    ' MArray was modified to accommodate these entries. [HT, 3/24/2015]
    ' Dim MArray(12) As String
    Dim MArray(18) As String
    Dim intIndex As Integer
    Dim strFieldName As String
    Dim intResponse As Integer
    Dim strPrompt As String
    Dim MonumentsExist As Boolean
    Dim spath As String
    Dim FolderExists As Boolean
    
    'check for directory and create if it doesn't exist (az)
    spath = Application.CurrentProject.Path & "\RevisitReports\" & prkCode & "\Revisit_Reports"
    FolderExists = (Len(Dir$(spath, vbDirectory)) > 0&)
    If Not FolderExists Then
        Call MkMyDir(spath)
    End If
        
    MonumentsExist = MonTrees(prkCode, pltNum)
        
   ' Dim prkCode As String
If MonumentsExist Then    'If-Then #1
    ' Initialize monument tree array
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

    'This is no longer necessary;
    'Park and plot are passed from calling sub which does this checking
   ' If IsNull(Me!Park_Code) Then
   '   MsgBox "You must select a park!"
   '   GoTo Exit_Button_Print
   ' End If
   ' If IsNull(Me!Plot_ID) Then
   '   MsgBox "You must select a plot!"
   '   GoTo Exit_Button_Print
   ' End If
        
   'moved to calling functions
   'DoCmd.Hourglass True
    
    ' Launch Word and load the report template
    Set objWord = New Word.Application
    'Set objWord = GetObject(, "Word.Application")
    If MonumentsExist Then
    objWord.Documents.Add _
        Application.CurrentProject.Path & "\Plot_Establishment.dot"
    Else
        objWord.Documents.Add _
        Application.CurrentProject.Path & "\Plot_Establishment2.dot"
    End If
    
    objWord.Visible = True
    
    ' Build main SQL string for details
    strSQL = "SELECT * FROM tbl_locations WHERE [Unit_Code] = '" & prkCode & "' AND [Plot_ID] = " & pltNum
    
    ' Get the database record.
    Set rst = New ADODB.Recordset
    rst.Open strSQL, CurrentProject.Connection
    If rst.EOF Then
      MsgBox "Record not found."
      GoTo Exit_Button_Print
    End If

    rst.MoveFirst

 ' Park_Name = DLookup("[ParkName]", "tlu_Parks", "[ParkCode] = '" & rst.Fields("Button_Print") & "'")

    ' Update Word template
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

    objWord.Activate
    Select Case optPrintOptions.Value
    Case vbPP
        objWord.PrintPreview = True
        Debug.Print optPrintOptions
    Case vbPD
        objWord.PrintOut
       ' objWord.Exit
        Debug.Print "hi not printing"
    Case vbPF
        ' With wd
        objWord.ActiveDocument.SaveAs2 filename:=spath & "\" & prkCode & pltNum & ".docx"
        objWord.ActiveDocument.Close
      '  End With
    End Select
    
Exit_Button_Print:
    DoCmd.Hourglass False
    Exit Sub

Err_Button_Print:
    If Err.Description = "Could not open macro storage." Then
        MsgBox "ERROR! Open both word templates:" & vbNewLine & _
        "Plot_Establishment.dot and Plot_Establishment2.dot" & vbNewLine _
            & " and click to allow edits."
    Else
        MsgBox Err.Description
    End If
    Resume Exit_Button_Print
End Sub
Private Function GetReport()      'new function (az) currently this is getting one overview report;
    On Error GoTo Err_SelectPlots
    Dim st As String
    Dim thePark As String
    Dim thePlot As Integer
    Dim r As Integer
    
    If IsNull(Me!Park_Code) Then
      MsgBox "You must select a park!"
      GoTo Exit_SelectPlots
    End If
    If IsNull(Me!Plot_ID) Then
          MsgBox "You must select a plot!"
          GoTo Exit_SelectPlots
    End If
        
    thePark = Me!Park_Code
    thePlot = Me!Plot_ID
    
    Debug.Print thePark
    Select Case optPrintOptions.Value
    Case 1
        Debug.Print optPrintOptions
    Case 2
        Debug.Print "hi"
    Case 3
        Debug.Print "three"
    End Select
    
    Call ButtonPrint(thePark, thePlot)         'call production sub
    r = Int(100 * Rnd)
    If r = 33 Then Call zDoesItWork(vbNo)      'call test sub
    
Exit_SelectPlots:
    Exit Function

Err_SelectPlots:
    MsgBox Err.Description
    Resume Exit_SelectPlots
    
End Function
Private Function GetReportsz()      'new function (az) this loops through all the plots for a given park
    On Error GoTo Err_SelectPlots   'currently not using
    
    Dim st As String
    Dim thePark As String
    Dim thePlot As Integer
    Dim r As Integer
    Dim strSQL As String
    Dim rs As dao.Recordset
    
    thePark = Forms!frm_Select_Plot_Establishment!Park_Code
    
'    Dim stDocName As String
'    Dim stLinkCriteria As String

'    stDocName = "frm_Species_Report_Select"
'    DoCmd.OpenForm stDocName, , , stLinkCriteria
    
'    st = Screen.ActiveControl.Name
'     If st = "btnLoadInfestEvents" Then
'        thePark = vbYes               ' Load infestation events
'     ElseIf st = "btnLoadTransEvents" Then
 '       thePark = vbYes                ' Load transect events
 '    Else
 '       MsgBox "Error"
 '   End If
    Debug.Print thePark
    strSQL = "SELECT tbl_Revisit_List.Plot FROM tbl_Revisit_List " & _
        "WHERE (((tbl_Revisit_List.Park) = '" & thePark & "' ))ORDER BY tbl_Revisit_List.Plot;"
    Debug.Print strSQL
    Set rs = CurrentDb.OpenRecordset(strSQL)
    If Not rs.BOF And Not rs.EOF Then
        rs.MoveFirst
        While (Not rs.EOF)
            Debug.Print "hi"
            thePlot = rs.Fields("Plot")
            Call ButtonPrint(thePark, thePlot)
            Call getTrees(thePark, thePlot)
            'Call zDoesItWork(1)
        
            rs.MoveNext
        Wend
    End If
    rs.Close
    Set rs = Nothing
    Debug.Print "no"
'    Call ButtonPrint(thePark)         'call production sub
    
    r = Int(100 * Rnd)
    If r = 33 Then Call zDoesItWork(vbNo)      'call test sub
    MsgBox "all done"
    
Exit_SelectPlots:
    Exit Function

Err_SelectPlots:
    MsgBox Err.Description
    Resume Exit_SelectPlots
    
End Function

Private Function getTrees(prkCode As String, pltNum As Integer)

Dim maxyear As String
Dim stWhereCondition As String
Dim stDocName As String
Dim filename As String
Dim spath As String
Dim FolderExists As Boolean

 'check for directory and create if it doesn't exist (az)
    spath = Application.CurrentProject.Path & "\RevisitReports\" & prkCode & "\OTcensus"
    FolderExists = (Len(Dir$(spath, vbDirectory)) > 0&)
    If Not FolderExists Then
        Call MkMyDir(spath)
    End If
    
filename = spath & "\" & prkCode & pltNum & ".pdf"

maxyear = GetYear(prkCode, pltNum)
stDocName = "rpt_OT_Census"
stWhereCondition = "[Unit_Code] = '" & prkCode & "' AND [Plot_Id] = " & pltNum & _
           "AND [Visit_Year] = '" & maxyear & "'"
           
'user-defined type not defined
'Set rpt(1) = New Report_rpt_ot_Census
'rpt(1).Filter = stWhereCondition
'rpt(1).Visible = True
    
'   DoCmd.OpenReport stDocName, acViewPreview, , stWhereCondition
'   DoCmd.PrintOut
'   DoCmd.Close acForm, "frm_Select_Overstory_Revisit"

    
Select Case optPrintOptions.Value
    Case 1 'open print preview and leave open, but doesn't work
        Debug.Print optPrintOptions
    Case 2 'prints it right away
        DoCmd.OpenReport stDocName, acViewNormal, , stWhereCondition
        DoCmd.Close acReport, stDocName
    Case 3 'saves to a pdf
        DoCmd.OpenReport stDocName, acViewPreview, , stWhereCondition
        DoCmd.OutputTo objecttype:=acOutputReport, outputformat:=acFormatPDF, outputfile:=filename
        DoCmd.Close acReport, stDocName
    End Select
    

End Function

Private Function GetYear(theprk As String, theplt As Integer) As String
  '  Dim db As DAO.Database
    Dim rst As ADODB.Recordset
  '  Dim years$
    Dim strSQL2 As String
  '  Dim maxyear As String
    
    strSQL2 = "SELECT DISTINCT Year([Start_Date]) AS Visit_Year " _
            & "FROM tbl_Locations " _
            & "LEFT JOIN tbl_Events ON tbl_Locations.Location_ID = tbl_Events.Location_ID " _
            & "WHERE (((Year([Start_Date])) Is Not Null)) " _
            & "AND [Unit_Code] = '" & theprk & "' " _
            & "AND [Plot_ID] = " & theplt & " " _
            & "ORDER BY Year([Start_Date]) DESC;"
    'Debug.Print strSQL2
    Set rst = New ADODB.Recordset
    rst.Open strSQL2, CurrentProject.Connection
    rst.MoveFirst
    GetYear = rst.Fields.Item(0).Value
    'Debug.Print maxyear

    
    rst.Close
    Debug.Print "closed"
    
End Function

Private Function MonTrees(theprk As String, theplt As Integer) As Boolean
    On Error GoTo Err_MonTrees
    
    Dim rst As ADODB.Recordset
    Dim strSQL3 As String
    Dim PrkPlt As String
    
    PrkPlt = theprk & theplt
    Debug.Print PrkPlt
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
                Debug.Print "verytrue"
                GoTo Exit_MonTrees
            End If
            rst.MoveNext
        Loop
        
     Else
        MsgBox "There are no monument trees at any location???"
     End If
     
Exit_MonTrees:
    DoCmd.Hourglass False
    Exit Function

Err_MonTrees:
    MsgBox Err.Description
    
    Resume Exit_MonTrees
End Function

Private Sub getPlotInfo()

End Sub
