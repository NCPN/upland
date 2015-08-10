Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7860
    DatasheetFontHeight =9
    ItemSuffix =22
    Top =210
    Right =8055
    Bottom =4365
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xc48e9ef32e50e340
    End
    RecordSource ="tbl_Impact_Details"
    Caption ="frm_sub_Impact_Details"
    OnCurrent ="[Event Procedure]"
    BeforeInsert ="[Event Procedure]"
    DatasheetFontName ="Arial"
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
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =3000
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =120
                    Top =120
                    Width =510
                    Height =180
                    ColumnWidth =2310
                    Name ="Impact_Details_ID"
                    ControlSource ="Impact_Details_ID"
                    StatusBarText ="Unique record identifier - primary key"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =900
                    Top =120
                    Width =510
                    Height =180
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Impact_ID"
                    ControlSource ="Impact_ID"
                    StatusBarText ="Foreign key to tbl_Site_Impact"
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =120
                    Top =780
                    Width =7560
                    TabIndex =4
                    Name ="Disturbance_Size"
                    ControlSource ="Disturbance_Size"
                    StatusBarText ="Size of disturbance.  New field for 2008."
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =120
                            Top =540
                            Width =2100
                            Height =240
                            FontWeight =700
                            Name ="Disturbance_Size_Label"
                            Caption ="Disturbance Size (m2)"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =120
                    Top =1440
                    Width =7560
                    Height =420
                    TabIndex =5
                    Name ="Disturbance_Position"
                    ControlSource ="Disturbance_Position"
                    StatusBarText ="Position of disturbance relative to transects.  New field for 2008."
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =120
                            Top =1200
                            Width =3735
                            Height =240
                            FontWeight =700
                            Name ="Disturbance_Position_Label"
                            Caption ="Disturbance Position Relative to Transects"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =120
                    Top =2280
                    Width =7620
                    Height =645
                    TabIndex =8
                    Name ="Disturbance_Description"
                    ControlSource ="Disturbance_Description"
                    StatusBarText ="Description of disturbance including potential effects on fire or erosion proces"
                        "ses.  New field for 2008."
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =120
                            Top =2040
                            Width =2160
                            Height =240
                            FontWeight =700
                            Name ="Disturbance_Description_Label"
                            Caption ="Disturbance Description"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =1260
                    Left =2580
                    Top =120
                    Width =1740
                    TabIndex =2
                    Name ="Disturbance_Location"
                    ControlSource ="Disturbance_Location"
                    RowSourceType ="Value List"
                    RowSource ="\"onsite\";\"offsite-upslope\";\"offsite-other\""
                    ColumnWidths ="1260"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =120
                            Top =120
                            Width =2400
                            Height =245
                            FontWeight =700
                            Name ="Observation Location Type_Label"
                            Caption ="Observation Location Type"
                            EventProcPrefix ="Observation_Location_Type_Label"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =2160
                    Left =6240
                    Top =120
                    Width =1500
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="Disturbance_Type"
                    ControlSource ="Disturbance_Type"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Disturbance.Disturbance FROM tlu_Disturbance; "
                    ColumnWidths ="2160"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4560
                            Top =120
                            Width =1620
                            Height =245
                            FontWeight =700
                            Name ="Disturbance Type_Label"
                            Caption ="Disturbance Type"
                            EventProcPrefix ="Disturbance_Type_Label"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =120
                    Top =1440
                    Width =3000
                    Height =420
                    TabIndex =6
                    Name ="Disturbance_Distance"
                    ControlSource ="Disturbance_Distance"
                    StatusBarText ="Position of disturbance relative to transects.  New field for 2008."
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =223
                            Left =120
                            Top =1200
                            Width =3240
                            Height =240
                            FontWeight =700
                            Name ="Distance_Label"
                            Caption ="Distance Upslope from Macroplot (m)"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =4200
                    Top =1440
                    Width =3480
                    Height =420
                    TabIndex =7
                    Name ="Disturbance_Direction"
                    ControlSource ="Disturbance_Direction"
                    StatusBarText ="Position of disturbance relative to transects.  New field for 2008."
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =223
                            Left =4200
                            Top =1200
                            Width =3480
                            Height =240
                            FontWeight =700
                            Name ="Direction_Label"
                            Caption ="Direction from Macroplot (azimuth-deg)"
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

Private Sub Disturbance_Location_AfterUpdate()
  If IsNull(Me!Disturbance_Location) Or Me!Disturbance_Location = "Onsite" Then
    Me!Disturbance_Position.Visible = True
    Me!Disturbance_Distance.Visible = False
    Me!Disturbance_Direction.Visible = False
  ElseIf Me!Disturbance_Location = "offsite-upslope" Then
    Me!Disturbance_Position.Visible = False
    Me!Disturbance_Distance.Visible = True
    Me!Distance_Label.Caption = "Distance Upslope from Macroplot (m)"
    Me!Disturbance_Direction.Visible = False
  Else
    Me!Disturbance_Position.Visible = False
    Me!Disturbance_Distance.Visible = True
    Me!Distance_Label.Caption = "Distance from Macroplot (m)"
    Me!Disturbance_Direction.Visible = True
  End If
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    On Error GoTo Err_Handler
    If IsNull(Me.Parent!Visit_Date) Then
      MsgBox "You must enter Visit Date first."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
      GoTo Exit_Procedure
    End If
    ' Create the GUID primary key value
    If IsNull(Me!Impact_Details_ID) Then
      Me!Impact_Details_ID = fxnGUIDGen
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

Private Sub Form_Current()
  If IsNull(Me!Disturbance_Location) Or Me!Disturbance_Location = "Onsite" Then
    Me!Disturbance_Position.Visible = True
    Me!Disturbance_Distance.Visible = False
    Me!Disturbance_Direction.Visible = False
  ElseIf Me!Disturbance_Location = "offsite-upslope" Then
    Me!Disturbance_Position.Visible = False
    Me!Disturbance_Distance.Visible = True
    Me!Distance_Label.Caption = "Distance Upslope from Macroplot (m)"
    Me!Disturbance_Direction.Visible = False
  Else
    Me!Disturbance_Position.Visible = False
    Me!Disturbance_Distance.Visible = True
    Me!Distance_Label.Caption = "Distance from Macroplot (m)"
    Me!Disturbance_Direction.Visible = True
  End If
End Sub
