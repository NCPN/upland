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
    ItemSuffix =112
    Left =300
    Top =2730
    Right =13815
    Bottom =9585
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xe5c8ddb21374e540
    End
    RecordSource ="qry_VH_Intercept"
    Caption ="fsub_VH_Intercept"
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
            Height =735
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Top =480
                    Width =840
                    Height =240
                    FontWeight =700
                    Name ="Point_Label"
                    Caption ="Point (m)"
                    Tag ="DetachedLabel"
                    LayoutCachedTop =480
                    LayoutCachedWidth =840
                    LayoutCachedHeight =720
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =960
                    Top =480
                    Width =1799
                    Height =240
                    FontWeight =700
                    Name ="Woody_Label"
                    Caption ="Woody Species"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =960
                    LayoutCachedTop =480
                    LayoutCachedWidth =2759
                    LayoutCachedHeight =720
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2889
                    Top =480
                    Width =1350
                    Height =255
                    FontWeight =700
                    Name ="WoodHt_Label"
                    Caption ="Woody Ht (cm)"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =2889
                    LayoutCachedTop =480
                    LayoutCachedWidth =4239
                    LayoutCachedHeight =735
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3840
                    Top =60
                    Width =1500
                    Height =300
                    ForeColor =8421376
                    Name ="ButtonInitialize"
                    Caption ="Initialize Form"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =3840
                    LayoutCachedTop =60
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =360
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =180
                    Top =60
                    Width =1500
                    Height =300
                    TabIndex =1
                    Name ="ButtonLookup"
                    Caption ="Master Lookup"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =1680
                    LayoutCachedHeight =360
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2040
                    Top =60
                    Width =1500
                    Height =300
                    TabIndex =2
                    Name ="ButtonUnknown"
                    Caption ="Unknown Species"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =2040
                    LayoutCachedTop =60
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =360
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =4369
                    Top =480
                    Width =1980
                    Height =240
                    FontWeight =700
                    Name ="Herb_label"
                    Caption ="Herb Species"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =4369
                    LayoutCachedTop =480
                    LayoutCachedWidth =6349
                    LayoutCachedHeight =720
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =6479
                    Top =480
                    Width =1170
                    Height =255
                    FontWeight =700
                    Name ="HerbHt_label"
                    Caption ="Herb Ht (cm)"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =6479
                    LayoutCachedTop =480
                    LayoutCachedWidth =7649
                    LayoutCachedHeight =735
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11940
                    Top =240
                    Width =480
                    ColumnOrder =0
                    TabIndex =3
                    Name ="txtStep_Val1"
                    DefaultValue ="5"

                    LayoutCachedLeft =11940
                    LayoutCachedTop =240
                    LayoutCachedWidth =12420
                    LayoutCachedHeight =480
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =10920
                            Top =240
                            Width =900
                            Height =240
                            Name ="Label440"
                            Caption ="Step Value:"
                            LayoutCachedLeft =10920
                            LayoutCachedTop =240
                            LayoutCachedWidth =11820
                            LayoutCachedHeight =480
                        End
                    End
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =12540
                    Top =240
                    Width =840
                    Height =240
                    Name ="Label441"
                    Caption ="Meters"
                    LayoutCachedLeft =12540
                    LayoutCachedTop =240
                    LayoutCachedWidth =13380
                    LayoutCachedHeight =480
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
                    TabIndex =1
                    Name ="Intercept_ID"
                    ControlSource ="Intercept_ID"
                    StatusBarText ="Unique record identifier - primary key"

                    LayoutCachedLeft =600
                    LayoutCachedTop =60
                    LayoutCachedWidth =1020
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Top =60
                    Width =840
                    Height =255
                    ColumnWidth =2310
                    FontSize =6
                    FontWeight =700
                    Name ="Point"
                    ControlSource ="Point"
                    Format ="General Number"
                    StatusBarText ="Intercept point - increments of .5m up to 50.0"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="5"

                    LayoutCachedTop =60
                    LayoutCachedWidth =840
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =7380
                    Top =60
                    Width =1200
                    TabIndex =6
                    Name ="Transect_ID"
                    ControlSource ="Transect_ID"
                    StatusBarText ="Foreign key to tbl_Canopy_Transect"

                    LayoutCachedLeft =7380
                    LayoutCachedTop =60
                    LayoutCachedWidth =8580
                    LayoutCachedHeight =300
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =6480
                    Left =4369
                    Top =60
                    Width =1980
                    TabIndex =4
                    BackColor =62207
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    ConditionalFormat = Begin
                        0x01000000a8000000020000000100000000000000000000000f00000001000000 ,
                        0x00000000fff20000010000000000000010000000230000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0048006500720062005d0020004900730020004e0075006c006c0000000000 ,
                        0x5b0048006500720062005d0020004900730020004e006f00740020004e007500 ,
                        0x6c006c0000000000
                    End
                    Name ="Herb"
                    ControlSource ="Herb"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Q.Master_PLANT_Code, Q.LU_Code, Q.Utah_Species, Q.Lifeform FROM (SELECT D"
                        "ISTINCT qryU_Top_Canopy.Master_PLANT_Code, qryU_Top_Canopy.LU_Code, qryU_Top_Can"
                        "opy.Utah_Species, qryU_Top_Canopy.Lifeform  FROM qryU_Top_Canopy WHERE (qryU_Top"
                        "_Canopy.Utah_Species Is Not Null) AND qryU_Top_Canopy.Lifeform IN ('Forb', 'Gram"
                        "inoid')  UNION   SELECT DISTINCT tbl_Unknown_Species.Unknown_Code AS Master_PLAN"
                        "T_Code, tbl_Unknown_Species.Unknown_Code AS LU_Code, tbl_Unknown_Species.Plant_T"
                        "ype + \" - \" + tbl_Unknown_Species.Plant_Description AS Utah_Species, tbl_Unkno"
                        "wn_Species.Plant_Type AS Lifeform  FROM tbl_Unknown_Species  WHERE tbl_Unknown_S"
                        "pecies.Plant_Type IN ('Grass', 'Other') OR tbl_Unknown_Species.Plant_Type IS NUL"
                        "L  UNION  SELECT TOP 1 'NP' AS Master_PLANT_Code, 'NP' AS LU_Code, 'No Plant' AS"
                        " Utah_Species, NULL AS Lifeform FROM tlu_NCPN_Plants UNION SELECT Top 1 'NR' AS "
                        "Master_PLANT_Code, 'NR' AS LU_Code, 'Not Recorded' AS Utah_Species, NULL AS Life"
                        "form FROM tlu_NCPN_Plants UNION SELECT qryU_Top_Canopy.Master_PLANT_Code, qryU_T"
                        "op_Canopy.LU_Code, qryU_Top_Canopy.Utah_Species, qryU_Top_Canopy.Lifeform FROM q"
                        "ryU_Top_Canopy WHERE qryU_Top_Canopy.Master_PLANT_Code = \"UNK\" UNION SELECT qr"
                        "yU_Top_Canopy.Master_PLANT_Code, qryU_Top_Canopy.LU_Code, qryU_Top_Canopy.Utah_S"
                        "pecies, qryU_Top_Canopy.Lifeform FROM qryU_Top_Canopy WHERE qryU_Top_Canopy.Mast"
                        "er_PLANT_Code = \"UNKH\")  AS Q ORDER BY Q.LU_Code;"
                    ColumnWidths ="0;2160;4320"
                    AfterUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"
                    LayoutCachedLeft =4369
                    LayoutCachedTop =60
                    LayoutCachedWidth =6349
                    LayoutCachedHeight =300
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff200000e0000005b00 ,
                        0x48006500720062005d0020004900730020004e0075006c006c00000000000000 ,
                        0x00000000000000000000000000000001000000000000000100000000000000ff ,
                        0xffff00120000005b0048006500720062005d0020004900730020004e006f0074 ,
                        0x0020004e0075006c006c00000000000000000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =6480
                    Left =960
                    Top =60
                    Width =1799
                    TabIndex =2
                    BackColor =62207
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    ConditionalFormat = Begin
                        0x01000000c0000000030000000100000000000000000000000f00000001000000 ,
                        0x00000000fff20000010000000000000010000000230000000100000000000000 ,
                        0xffffff000100000000000000240000002f0000000100000000000000ba141900 ,
                        0x5b0057006f006f0064005d0020004900730020004e0075006c006c0000000000 ,
                        0x5b0057006f006f0064005d0020004900730020004e006f00740020004e007500 ,
                        0x6c006c00000000005b0050006f0069006e0074005d003d003300300000000000
                    End
                    Name ="Wood"
                    ControlSource ="Wood"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Q.Master_PLANT_Code, Q.LU_Code, Q.Utah_Species, Q.Lifeform FROM (SELECT D"
                        "ISTINCT qryU_Top_Canopy.Master_PLANT_Code, qryU_Top_Canopy.LU_Code, qryU_Top_Can"
                        "opy.Utah_Species, qryU_Top_Canopy.Lifeform  FROM qryU_Top_Canopy WHERE (qryU_Top"
                        "_Canopy.Utah_Species Is Not Null) AND qryU_Top_Canopy.Lifeform IN ('DwarfShrub',"
                        " 'Shrub', 'Tree')  UNION   SELECT DISTINCT tbl_Unknown_Species.Unknown_Code AS M"
                        "aster_PLANT_Code, tbl_Unknown_Species.Unknown_Code AS LU_Code, tbl_Unknown_Speci"
                        "es.Plant_Type + \" - \" + tbl_Unknown_Species.Plant_Description AS Utah_Species,"
                        " tbl_Unknown_Species.Plant_Type AS Lifeform  FROM tbl_Unknown_Species  WHERE tbl"
                        "_Unknown_Species.Plant_Type IN ('Shrub', 'Tree', 'Other') OR tbl_Unknown_Species"
                        ".Plant_Type IS NULL UNION  SELECT Top 1 'NP' AS Master_PLANT_Code, 'NP' AS LU_Co"
                        "de, 'No Plant' AS Utah_Species, NULL AS Lifeform FROM tlu_NCPN_Plants UNION SELE"
                        "CT Top 1 'NR' AS Master_PLANT_Code, 'NR' AS LU_Code, 'Not Recorded' AS Utah_Spec"
                        "ies, NULL AS Lifeform FROM tlu_NCPN_Plants UNION SELECT qryU_Top_Canopy.Master_P"
                        "LANT_Code, qryU_Top_Canopy.LU_Code, qryU_Top_Canopy.Utah_Species, qryU_Top_Canop"
                        "y.Lifeform FROM qryU_Top_Canopy WHERE qryU_Top_Canopy.Master_PLANT_Code = \"UNK\""
                        ")  AS Q ORDER BY Q.LU_Code;"
                    ColumnWidths ="0;2160;4320"
                    AfterUpdate ="[Event Procedure]"
                    LayoutCachedLeft =960
                    LayoutCachedTop =60
                    LayoutCachedWidth =2759
                    LayoutCachedHeight =300
                    ConditionalFormat14 = Begin
                        0x01000300000001000000000000000100000000000000fff200000e0000005b00 ,
                        0x57006f006f0064005d0020004900730020004e0075006c006c00000000000000 ,
                        0x00000000000000000000000000000001000000000000000100000000000000ff ,
                        0xffff00120000005b0057006f006f0064005d0020004900730020004e006f0074 ,
                        0x0020004e0075006c006c00000000000000000000000000000000000000000000 ,
                        0x01000000000000000100000000000000ba1419000a0000005b0050006f006900 ,
                        0x6e0074005d003d00330030000000000000000000000000000000000000000000 ,
                        0x00
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2889
                    Top =60
                    Width =1350
                    TabIndex =3
                    BackColor =62207
                    Name ="txtWHeight"
                    ControlSource ="WHeight"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    ConditionalFormat = Begin
                        0x0100000080000000030000000000000006000000000000000500000001000000 ,
                        0x00000000f9eded00000000000100000006000000080000000100000000000000 ,
                        0xfff2000000000000060000000d0000000f0000000100000000000000ffffff00 ,
                        0x3100300030003100000000003000000031003000300030000000300000000000
                    End

                    LayoutCachedLeft =2889
                    LayoutCachedTop =60
                    LayoutCachedWidth =4239
                    LayoutCachedHeight =300
                    ConditionalFormat14 = Begin
                        0x01000300000000000000060000000100000000000000f9eded00040000003100 ,
                        0x3000300031000000000000000000000000000000000000000000000000000001 ,
                        0x0000000100000000000000fff200000100000030000400000031003000300030 ,
                        0x0000000000000000000000000000000000000000000006000000010000000000 ,
                        0x0000ffffff000100000030000000000000000000000000000000000000000000 ,
                        0x00
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6479
                    Top =60
                    Width =1170
                    TabIndex =5
                    BackColor =62207
                    Name ="txtHHeight"
                    ControlSource ="HHeight"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    ConditionalFormat = Begin
                        0x0100000082000000030000000000000006000000000000000500000001000000 ,
                        0x00000000f9eded00000000000100000006000000080000000100000000000000 ,
                        0xfff2000000000000040000000d000000100000000100000000000000ffffff00 ,
                        0x31003000300031000000000030000000310030003000300000002d0031000000 ,
                        0x0000
                    End

                    LayoutCachedLeft =6479
                    LayoutCachedTop =60
                    LayoutCachedWidth =7649
                    LayoutCachedHeight =300
                    ConditionalFormat14 = Begin
                        0x01000300000000000000060000000100000000000000f9eded00040000003100 ,
                        0x3000300031000000000000000000000000000000000000000000000000000001 ,
                        0x0000000100000000000000fff200000100000030000400000031003000300030 ,
                        0x0000000000000000000000000000000000000000000004000000010000000000 ,
                        0x0000ffffff00020000002d003100000000000000000000000000000000000000 ,
                        0x000000
                    End
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
    Set Points = db.OpenRecordset("tbl_VH_Intercept")
    Veg_Type = DLookup("[Vegetation_Type]", "tbl_Locations", "[Location_ID] = '" & Me.Parent.Parent!Location_ID & "'")
    If Not IsNull(Veg_Type) And Veg_Type = "oak scrub" Then
      PointCount = 5
      PointIncrement = 5
      PointLimit = 50
    Else
      PointCount = 5
      PointIncrement = 5
      PointLimit = 50
    End If
    Do Until PointCount > PointLimit
      Points.AddNew
      Points!Intercept_ID = fxnGUIDGen  ' Generate an ID for it
      Points!Transect_ID = Forms!frm_Data_Entry!frm_VH_Transect.Form!Transect_ID
      Points!Point = PointCount
'      Points!Alive = -1
'      Points!Surface_Alive = 0
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



Private Sub Form_Current()


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
    strSQL = "SELECT [Point] FROM [tbl_VH_Intercept] WHERE [Transect_ID] = '" & Me![Transect_ID] & "'"
    Set Points = db.OpenRecordset(strSQL)
    
    If Points.EOF Or IsNull(Points!Point) Then
      Me!ButtonInitialize.ForeColor = lngDkBrtGrn '8421376
    Else
      Me!ButtonInitialize.ForeColor = lngRed '255

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


Private Sub Herb_AfterUpdate()
If (Me.Herb = "NP" Or Me.Herb = "NR") Then
Me.HHeight = "0"
End If
End Sub

Private Sub Herb_GotFocus()

End Sub

Private Sub txtHHeight_AfterUpdate()
If Me.txtHHeight > 1000 Then
MsgBox "Height over 1000 cm (10 meters). Check that height entered is correct.", vbOKOnly, "Check Height"
End If

End Sub

Private Sub txtWHeight_AfterUpdate()
If Me.txtWHeight > 1000 Then
MsgBox "Height over 1000 cm (10 meters). Check that height entered is correct.", vbOKOnly, "Check Height"
End If

End Sub


Private Sub Wood_AfterUpdate()
If (Me.Wood = "NP" Or Me.Wood = "NR") Then
Me.WHeight = "0"
End If
Refresh

End Sub
