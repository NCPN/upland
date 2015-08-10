Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5040
    DatasheetFontHeight =9
    ItemSuffix =11
    Left =5625
    Top =4635
    Right =11010
    Bottom =7140
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x9d2210c6b41ee340
    End
    Caption ="Select for Plot Revisit Data Sheet"
    DatasheetFontName ="Arial"
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
        End
        Begin ComboBox
            SpecialEffect =2
            FontName ="Tahoma"
        End
        Begin Section
            Height =3600
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =3600
                    Left =2280
                    Top =1080
                    Width =840
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"10\""
                    Name ="Park_Code"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Parks.ParkCode, tlu_Parks.ParkName FROM tlu_Parks; "
                    ColumnWidths ="720;2880"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =960
                            Top =1080
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
                    Left =720
                    Top =540
                    Width =3615
                    Height =360
                    FontSize =12
                    FontWeight =700
                    Name ="Label2"
                    Caption ="Overstory Revisit Data Sheet"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =600
                    Top =2580
                    Width =1395
                    Height =300
                    TabIndex =3
                    Name ="ButtonClose"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =540
                    Left =2280
                    Top =1560
                    Width =840
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="Plot_ID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Locations.Location_ID, tbl_Locations.Plot_ID FROM tbl_Locations; "
                    ColumnWidths ="540"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1575
                            Top =1560
                            Width =645
                            Height =245
                            FontWeight =700
                            Name ="Plot ID_Label"
                            Caption ="Plot ID"
                            EventProcPrefix ="Plot_ID_Label"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2880
                    Top =2580
                    Height =300
                    TabIndex =4
                    Name ="ButtonReport"
                    Caption ="Preview Report"
                    OnClick ="[Event Procedure]"
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =510
                    Left =2280
                    Top =2040
                    Width =840
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"3\";\"2\""
                    Name ="Visit_Year"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qry_sel_Visit_Year.Visit_Year FROM qry_sel_Visit_Year; "
                    ColumnWidths ="510"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1260
                            Top =2040
                            Width =960
                            Height =245
                            FontWeight =700
                            Name ="Visit_Year_Label"
                            Caption ="Visit_Year"
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

Private Sub ButtonClose_Click()
On Error GoTo Err_ButtonClose_Click

    DoCmd.Close

Exit_ButtonClose_Click:
    Exit Sub

Err_ButtonClose_Click:
    MsgBox Err.Description
    Resume Exit_ButtonClose_Click
    
End Sub

Private Sub Park_Code_AfterUpdate()

  Me!Plot_ID = Null
  If Not IsNull(Me!Park_Code) Then
    Me!Plot_ID.RowSource = "SELECT Plot_ID FROM tbl_locations WHERE [Unit_Code] = '" & Me!Park_Code & "' ORDER BY Plot_ID"
    Me!Plot_ID.Requery
  Else
    MsgBox "You must select a park!"
  End If
    
End Sub

Private Sub ButtonReport_Click()
On Error GoTo Err_ButtonReport_Click

    Dim stDocName As String
    Dim stWhereCondition As String
    If IsNull(Me!Park_Code) Or IsNull(Me!Plot_ID) Or IsNull(Me!Visit_Year) Then
      MsgBox "Park Code, Plot Number, and Visit Year are all required."
      Exit Sub
    End If
    stWhereCondition = "[Unit_Code] = '" & Me!Park_Code & "' AND [Plot_Id] = " & Me!Plot_ID & "AND [Visit_Year] = '" & Me!Visit_Year & "'"
    stDocName = "rpt_OT_Census"
    DoCmd.OpenReport stDocName, acViewPreview, , stWhereCondition
    DoCmd.Close acForm, "frm_Select_Overstory_Revisit"
Exit_ButtonReport_Click:
    Exit Sub

Err_ButtonReport_Click:
    MsgBox Err.Description
    Resume Exit_ButtonReport_Click
    
End Sub
