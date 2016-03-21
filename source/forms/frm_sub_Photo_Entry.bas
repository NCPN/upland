Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    TabularFamily =127
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =13680
    DatasheetFontHeight =9
    ItemSuffix =26
    Left =1356
    Right =15012
    Bottom =6660
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x2f2916b0ec7ce340
    End
    RecordSource ="tbl_Photos"
    Caption ="frm_Photo_Entry"
    BeforeInsert ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
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
            Height =480
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =3675
                    Top =240
                    Width =1065
                    Height =240
                    FontWeight =700
                    Name ="Photo_Date_Label"
                    Caption ="Photo Date"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =60
                    Top =240
                    Width =1395
                    Height =240
                    FontWeight =700
                    Name ="Transect_Label"
                    Caption ="Location"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1920
                    Top =240
                    Width =720
                    Height =240
                    FontWeight =700
                    Name ="Roll_Label"
                    Caption ="Roll #"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2700
                    Top =60
                    Width =960
                    Height =420
                    FontWeight =700
                    Name ="Frame_Label"
                    Caption ="Frame or Photo #"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =7515
                    Top =240
                    Width =1545
                    Height =240
                    FontWeight =700
                    Name ="Digital_File_Label"
                    Caption ="Digital File Name"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =4860
                    Top =240
                    Width =1440
                    Height =240
                    FontWeight =700
                    Name ="Photographer_Label"
                    Caption ="Photographer"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    Left =6300
                    Top =60
                    Width =1200
                    Height =420
                    FontWeight =700
                    Name ="Location_Label"
                    Caption ="Location on Transect (m)"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =10980
                    Top =240
                    Width =2280
                    Height =240
                    FontWeight =700
                    Name ="Comments_Label"
                    Caption ="Comments"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =9255
                    Top =240
                    Width =1575
                    Height =240
                    FontWeight =700
                    Name ="NCPN_Image_ID_Label"
                    Caption ="NCPN Image ID"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            Height =360
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =240
                    Height =255
                    ColumnWidth =2310
                    Name ="Photo_ID"
                    ControlSource ="Photo_ID"
                    StatusBarText ="Unique record identifier - primary key"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =420
                    Top =60
                    Width =240
                    Height =255
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Event_ID"
                    ControlSource ="Event_ID"
                    StatusBarText ="Foreign key to tbl_Events"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3660
                    Top =60
                    Width =1080
                    Height =255
                    ColumnWidth =1035
                    TabIndex =5
                    Name ="Photo_Date"
                    ControlSource ="Photo_Date"
                    Format ="Short Date"
                    StatusBarText ="Date photograph taken."
                    InputMask ="99/99/0000;0;_"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1965
                    Top =60
                    Width =615
                    Height =255
                    ColumnWidth =900
                    TabIndex =3
                    Name ="Roll"
                    ControlSource ="Roll"
                    StatusBarText ="Roll for film photos"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2880
                    Top =60
                    Width =540
                    Height =255
                    ColumnWidth =600
                    TabIndex =4
                    Name ="Frame"
                    ControlSource ="Frame"
                    StatusBarText ="Frame for film photos"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7500
                    Top =60
                    Width =1545
                    Height =255
                    ColumnWidth =2310
                    TabIndex =8
                    Name ="Digital_File"
                    ControlSource ="Digital_File"
                    StatusBarText ="Digital file name"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6540
                    Top =60
                    Width =600
                    Height =255
                    ColumnWidth =600
                    TabIndex =7
                    Name ="Location"
                    ControlSource ="Location"
                    StatusBarText ="Location of photo point along transect in meters"

                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10980
                    Top =60
                    Width =2280
                    Height =255
                    ColumnWidth =3000
                    TabIndex =10
                    Name ="Comments"
                    ControlSource ="Comments"
                    StatusBarText ="Photo comments"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9240
                    Top =60
                    Width =1575
                    Height =255
                    ColumnWidth =2310
                    TabIndex =9
                    Name ="NCPN_Image_ID"
                    ControlSource ="NCPN_Image_ID"
                    StatusBarText ="Digital file name in NCPN Photo Database"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =247
                    IMESentenceMode =3
                    ListRows =9
                    ListWidth =1440
                    Left =60
                    Top =60
                    Width =1800
                    TabIndex =2
                    Name ="Transect"
                    ControlSource ="Transect"
                    RowSourceType ="Value List"
                    RowSource ="T1 - origin;T2 - origin;T3 - origin;T1 - end;T2 - end;T3 - end"
                    ColumnWidths ="1440"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =1965
                    Left =4860
                    Top =60
                    TabIndex =6
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Photographer"
                    ControlSource ="Photographer"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Contacts.Contact_ID, tlu_Contacts.Last_Name, tlu_Contacts.First_Name "
                        "FROM tlu_Contacts WHERE (((tlu_Contacts.Active)=1)); "
                    ColumnWidths ="0;975;990"

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
    ' Default to Events Start Date if photo date is null
    If IsNull(Me.Parent!Start_Date) Then
      MsgBox "Missing site visit date."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
      Exit Sub
    ElseIf IsNull(Me!Photo_Date) Then
      Me!Photo_Date = Me.Parent!Start_Date
    End If
    ' Create the GUID primary key value if necessary
    If IsNull(Me!Photo_ID) Then
        If GetDataType("tbl_Photos", "Photo_ID") = dbText Then
            Me.Photo_ID = fxnGUIDGen
        End If
    End If
End Sub

Private Sub Form_Load()
  Dim Veg_Type As Variant
  ' Display the proper tabs
    Veg_Type = DLookup("[Vegetation_Type]", "tbl_Locations", "[Location_ID] = '" & Me.Parent!Location_ID & "'")
    If Not IsNull(Veg_Type) And (Veg_Type = "forest" Or Veg_Type = "oak scrub") Then
      Me!Transect.RowSource = "T1 - origin;T2 - origin;T3 - origin;T1 - end;T2 - end;T3 - end"
    End If
End Sub
