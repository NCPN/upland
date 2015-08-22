Version =20
VersionRequired =20
Begin Form
    Modal = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =124
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    DatasheetFontHeight =9
    ItemSuffix =11
    Left =3960
    Top =1395
    Right =11160
    Bottom =4320
    DatasheetGridlinesColor =12632256
    Filter ="[Unit_Code] = 'CURE' AND [Plot_ID] = 5 AND [VisitDate]=#3/1/2010#"
    RecSrcDt = Begin
        0xde6046048da8e340
    End
    RecordSource ="tbl_Revisit_Comments"
    Caption ="frm_Revisit_Comments"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
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
            Height =780
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2160
                    Top =180
                    Width =2880
                    Height =420
                    FontSize =14
                    FontWeight =700
                    Name ="Label8"
                    Caption ="Revisit Comments"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5580
                    Top =240
                    Width =960
                    Height =300
                    Name ="ButtonClose"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =300
                    Top =180
                    Width =1620
                    Height =540
                    FontWeight =700
                    TabIndex =1
                    ForeColor =16711680
                    Name ="ButtonAdd"
                    Caption ="Add New Comment Record"
                    OnClick ="[Event Procedure]"
                End
            End
        End
        Begin Section
            Height =2160
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1140
                    Top =120
                    Width =660
                    Height =239
                    ColumnWidth =540
                    TabIndex =1
                    Name ="Unit_Code"
                    ControlSource ="Unit_Code"
                    StatusBarText ="Park Code."
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =105
                            Top =120
                            Width =975
                            Height =240
                            FontWeight =700
                            Name ="Unit_Code_Label"
                            Caption ="Unit Code"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2940
                    Top =120
                    Width =600
                    Height =239
                    ColumnWidth =600
                    TabIndex =2
                    Name ="Plot_ID"
                    ControlSource ="Plot_ID"
                    StatusBarText ="Plot identifier"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2100
                            Top =120
                            Width =780
                            Height =239
                            FontWeight =700
                            Name ="Plot_ID_Label"
                            Caption ="Plot_ID"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4800
                    Top =120
                    Width =855
                    Height =239
                    ColumnWidth =1035
                    TabIndex =3
                    Name ="Visit_Date"
                    ControlSource ="VisitDate"
                    Format ="Short Date"
                    StatusBarText ="Date of visit corresponding to comment."
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =3825
                            Top =120
                            Width =915
                            Height =240
                            FontWeight =700
                            Name ="VisitDate_Label"
                            Caption ="Visit Date"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =180
                    Top =780
                    Width =6420
                    Height =1020
                    ColumnWidth =4245
                    Name ="Revisit_Comments"
                    ControlSource ="Revisit_Comments"
                    StatusBarText ="Comments for next revisit"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =180
                            Top =540
                            Width =1800
                            Height =240
                            FontWeight =700
                            Name ="Revisit_Comments_Label"
                            Caption ="Revisit Comments"
                        End
                    End
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

Private Sub ButtonClose_Click()
On Error GoTo Err_ButtonClose_Click

    DoCmd.Close

Exit_ButtonClose_Click:
    Exit Sub

Err_ButtonClose_Click:
    MsgBox Err.Description
    Resume Exit_ButtonClose_Click
    
End Sub

Private Sub Form_Close()
  Forms!frm_Data_Entry.visible = True
  Forms!frm_Data_Gateway.visible = True
End Sub

Private Sub Form_Load()
  Dim RevisitComments As DAO.Recordset
  Dim db As DAO.Database
  Dim strSQL As String
    
  If Not IsNull(Me.OpenArgs) Then
    DoCmd.GoToRecord , , acNewRec
    Me!Unit_Code = Forms!frm_Data_Entry!txtUnit_Code
    Me!Plot_ID = Forms!frm_Data_Entry!SiteDisplay
    Me!Visit_Date = Forms!frm_Data_Entry!txtStart_date
    strSQL = "SELECT * FROM tbl_Revisit_Comments Where [Unit_Code] = '" & Me!Unit_Code & "' AND [Plot_ID] = " & Me!Plot_ID & " ORDER BY [VisitDate] DESC"
    Set db = CurrentDb
    Set RevisitComments = db.OpenRecordset(strSQL)
    If Not RevisitComments.EOF Then
      RevisitComments.MoveFirst
      Me!Revisit_Comments = RevisitComments!Revisit_Comments
    End If
    RevisitComments.Close
    Set RevisitComments = Nothing
  End If
  Forms!frm_Data_Entry.visible = False
  Forms!frm_Data_Gateway.visible = False
  
End Sub

Private Sub Revisit_Comments_AfterUpdate()
  If IsNull(Me!Revisit_Comments) Then
    Forms!frm_Data_Entry!Comments = " "
  Else
    Forms!frm_Data_Entry!Comments = Me!Revisit_Comments
  End If
End Sub
Private Sub ButtonAdd_Click()
On Error GoTo Err_ButtonAdd_Click

    DoCmd.GoToRecord , , acNewRec
    Me!Unit_Code = Forms!frm_Data_Entry!txtUnit_Code
    Me!Plot_ID = Forms!frm_Data_Entry!SiteDisplay
    Me!Visit_Date = Forms!frm_Data_Entry!txtStart_date
    
Exit_ButtonAdd_Click:
    Exit Sub

Err_ButtonAdd_Click:
    MsgBox Err.Description
    Resume Exit_ButtonAdd_Click
    
End Sub
