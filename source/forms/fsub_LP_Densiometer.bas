Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    TabularFamily =124
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5760
    DatasheetFontHeight =9
    ItemSuffix =16
    Left =465
    Top =60
    Right =9285
    Bottom =3285
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x993ac9fb6e87e340
    End
    RecordSource ="qry_LP_Densiometer"
    Caption ="fsub_LP_Densiometer"
    BeforeInsert ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
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
            Height =360
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =360
                    Top =60
                    Width =540
                    Height =240
                    FontWeight =700
                    Name ="Point_Label"
                    Caption ="Point"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            Height =375
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =120
                    Top =60
                    Width =420
                    Height =255
                    ColumnWidth =2310
                    Name ="SD_ID"
                    ControlSource ="SD_ID"
                    StatusBarText ="Unique record identifier - primary key"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =660
                    Top =60
                    Width =420
                    Height =255
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Transect_ID"
                    ControlSource ="Transect_ID"
                    StatusBarText ="Foreign key to tbl_LP_Belt_Transect"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1320
                    Top =60
                    Width =600
                    Height =255
                    ColumnWidth =600
                    TabIndex =3
                    Name ="Total1"
                    ControlSource ="Total1"
                    StatusBarText ="Total count"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2160
                    Top =60
                    Width =600
                    Height =255
                    ColumnWidth =600
                    TabIndex =4
                    Name ="Total2"
                    ControlSource ="Total2"
                    StatusBarText ="Total count"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3000
                    Top =60
                    Width =600
                    Height =255
                    ColumnWidth =600
                    TabIndex =5
                    Name ="Total3"
                    ControlSource ="Total3"
                    StatusBarText ="Total count"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3840
                    Top =60
                    Width =600
                    Height =255
                    ColumnWidth =600
                    TabIndex =6
                    Name ="Total4"
                    ControlSource ="Total4"
                    StatusBarText ="Total count"
                    FontName ="Tahoma"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    ListWidth =540
                    Left =180
                    Top =60
                    Width =900
                    Height =255
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"10\";\"8\""
                    Name ="Point"
                    ControlSource ="Point"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qry_Densiometer_LU.Point FROM qry_Densiometer_LU; "
                    ColumnWidths ="540"
                    FontName ="Tahoma"
                    OnGotFocus ="[Event Procedure]"
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

Private Sub Form_BeforeInsert(Cancel As Integer)

    Dim db As DAO.Database
    Dim Points As DAO.Recordset
    Dim strSQL As String
    On Error GoTo Err_Handler

' Check for species overlap
      ' Set SQL
      Set db = CurrentDb
      strSQL = "SELECT tbl_LP_Belt_Transect.Transect_ID, Count(SD_ID) AS PointCount FROM tbl_LP_Belt_Transect LEFT JOIN tbl_LP_Densiometer ON tbl_LP_Belt_Transect.Transect_ID = tbl_LP_Densiometer.Transect_ID GROUP BY tbl_LP_Belt_Transect.Transect_ID HAVING tbl_LP_Belt_Transect.Transect_ID = '" & Me.Parent!Transect_ID & "'"
      Set Points = db.OpenRecordset(strSQL)
      If Points!PointCount > 4 Then
        MsgBox "5 points maximum!", , "Spherical Densiometer"
        DoCmd.CancelEvent
        SendKeys "{ESC}"
        GoTo Exit_Procedure
      End If
    ' Create the GUID primary key value
    If IsNull(Me!SD_ID) Then
        If GetDataType("tbl_LP_Densiometer", "SD_ID") = dbText Then
            Me.SD_ID = fxnGUIDGen
        End If
    End If
'    DoCmd.RunCommand acCmdSaveRecord  ' Save it.
Exit_Procedure:
  Points.Close
  Exit Sub
  
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

Private Sub Point_GotFocus()

    If IsNull(Me.Parent!Visit_Date) Then    ' If they didn't bother to enter a date, default to event date.
      Me.Parent!Visit_Date = Me.Parent.Parent!Start_Date
      Me.Parent.Refresh   ' Force save of transect record
    End If
 
End Sub
