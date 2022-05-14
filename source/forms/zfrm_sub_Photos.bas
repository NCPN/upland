Version =21
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =12240
    DatasheetFontHeight =10
    ItemSuffix =60
    Left =300
    Right =13215
    Bottom =8295
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x92aa8ac732dce240
    End
    RecordSource ="tbl_Photos"
    Caption ="Transect Photos"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
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
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =5760
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =420
                    Name ="Photo_ID"
                    ControlSource ="Photo_ID"
                    StatusBarText ="Unique record identifer"

                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =660
                    Top =1140
                    Width =480
                    TabIndex =3
                    Name ="PhotoRoll"
                    ControlSource ="Roll"
                    StatusBarText ="Reference number for film roll of photo."

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =1140
                            Width =420
                            Height =240
                            FontWeight =700
                            Name ="Label39"
                            Caption ="Roll"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =780
                    Top =1620
                    Width =480
                    TabIndex =4
                    Name ="PhotoFrame"
                    ControlSource ="Frame"
                    StatusBarText ="Frame number of photo within roll."

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =1620
                            Width =540
                            Height =240
                            FontWeight =700
                            Name ="Label40"
                            Caption ="Frame"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1620
                    Top =4500
                    TabIndex =10
                    Name ="PhotoOther"
                    ControlSource ="Other_Identifier"
                    StatusBarText ="Other unique identifier or reference number for digital photo or name of movie f"
                        "ile."

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =4500
                            Width =1380
                            Height =240
                            FontWeight =700
                            Name ="Label41"
                            Caption ="Other Identifier"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1740
                    Top =4980
                    Width =4020
                    Height =480
                    ColumnWidth =7230
                    TabIndex =11
                    Name ="PhotoComments"
                    ControlSource ="Comments"
                    StatusBarText ="Brief description of photo."

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =4980
                            Width =1500
                            Height =240
                            FontWeight =700
                            Name ="Label43"
                            Caption ="Photo Comments"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1800
                    Top =4020
                    Width =1740
                    TabIndex =9
                    Name ="Digital_File_Name"
                    ControlSource ="Digital_File"
                    StatusBarText ="File name of digital photograph."

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =4020
                            Width =1560
                            Height =240
                            FontWeight =700
                            Name ="Label44"
                            Caption ="Digital File Name"
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =6480
                    Top =960
                    Width =366
                    Height =366
                    TabIndex =12
                    Name ="ButtonPrevious"
                    Caption ="Command46"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddadadadad1dadadaadadadad11adadaddadadad111dadada ,
                        0xadadad1111adadaddadad11111dadadaadadad1111adadaddadadad111dadada ,
                        0xadadadad11adadaddadadadad1dadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Previous Record"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =8520
                    Top =960
                    Width =366
                    Height =366
                    TabIndex =13
                    Name ="ButtonNext"
                    Caption ="Command47"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddadada1adadadadaadadad11adadadaddadada111adadada ,
                        0xadadad1111adadaddadada11111adadaadadad1111adadaddadada111adadada ,
                        0xadadad11adadadaddadada1adadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Next Record"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =7020
                    Top =960
                    Width =1326
                    Height =366
                    TabIndex =14
                    Name ="ButtonAdd"
                    Caption ="Add Photograph"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Add Record"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =6600
                    Top =480
                    Width =2040
                    Height =240
                    FontWeight =700
                    Name ="Label49"
                    Caption ="Photograph Navigation"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =600
                    Top =60
                    Width =420
                    TabIndex =15
                    Name ="Event_ID"
                    ControlSource ="Event_ID"
                    StatusBarText ="Foreign key to tbl_Events"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1260
                    Top =180
                    Width =1080
                    TabIndex =1
                    Name ="Photo_Date"
                    ControlSource ="Photo_Date"
                    Format ="Short Date"
                    StatusBarText ="Date photograph taken."
                    InputMask ="99/99/0000;0;_"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =180
                            Top =180
                            Width =1020
                            Height =240
                            FontWeight =700
                            Name ="Label51"
                            Caption ="Photo Date"
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListRows =9
                    ListWidth =1080
                    Left =1140
                    Top =660
                    TabIndex =2
                    Name ="Transect"
                    ControlSource ="Transect"
                    RowSourceType ="Value List"
                    RowSource ="1;2;3;1stQuadT1;3rdQuadT2;5thQuadT3;CR1stQuadT1;CR3rdQuadT2;CR5thQuadT3"
                    ColumnWidths ="1080"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =660
                            Width =900
                            Height =245
                            FontWeight =700
                            Name ="Transect_Label"
                            Caption ="Transect"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1140
                    Top =2580
                    Width =480
                    TabIndex =6
                    Name ="Direction"
                    ControlSource ="Direction"
                    StatusBarText ="Direction of photograph"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =2580
                            Width =900
                            Height =240
                            FontWeight =700
                            Name ="Label54"
                            Caption ="Direction"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =1965
                    Left =1440
                    Top =2100
                    TabIndex =5
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Photographer"
                    ControlSource ="Photographer"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Contacts.Contact_ID, tlu_Contacts.Last_Name, tlu_Contacts.First_Name "
                        "FROM tlu_Contacts; "
                    ColumnWidths ="0;975;990"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =2100
                            Width =1200
                            Height =245
                            FontWeight =700
                            Name ="Photographer_Label"
                            Caption ="Photographer"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3900
                    Top =3060
                    Width =420
                    TabIndex =7
                    Name ="Location"
                    ControlSource ="Location"
                    StatusBarText ="Location of photo point along transect in meters"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =180
                            Top =3060
                            Width =3720
                            Height =240
                            FontWeight =700
                            Name ="Label57"
                            Caption ="Location of photo point along transect (m)"
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =600
                    Left =1320
                    Top =3540
                    Width =1080
                    TabIndex =8
                    Name ="Photo_Type"
                    ControlSource ="Photo_Type"
                    RowSourceType ="Value List"
                    RowSource ="\"film\";\"digital\";\"movie\";\"other\""
                    ColumnWidths ="600"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =3540
                            Width =1080
                            Height =245
                            FontWeight =700
                            Name ="Photo Type_Label"
                            Caption ="Photo Type"
                            EventProcPrefix ="Photo_Type_Label"
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


Private Sub ButtonPrevious_Click()
On Error GoTo Err_ButtonPrevious_Click


    DoCmd.GoToRecord , , acPrevious

Exit_ButtonPrevious_Click:
    Exit Sub

Err_ButtonPrevious_Click:
    MsgBox Err.Description
    Resume Exit_ButtonPrevious_Click
    
End Sub
Private Sub ButtonNext_Click()
On Error GoTo Err_ButtonNext_Click


    DoCmd.GoToRecord , , acNext

Exit_ButtonNext_Click:
    Exit Sub

Err_ButtonNext_Click:
    MsgBox Err.Description
    Resume Exit_ButtonNext_Click
    
End Sub
Private Sub ButtonAdd_Click()
On Error GoTo Err_ButtonAdd_Click


    DoCmd.GoToRecord , , acNewRec

Exit_ButtonAdd_Click:
    Exit Sub

Err_ButtonAdd_Click:
    MsgBox Err.Description
    Resume Exit_ButtonAdd_Click
    
End Sub
