Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =3600
    DatasheetFontHeight =9
    ItemSuffix =26
    Left =4620
    Top =1800
    Right =9915
    Bottom =4530
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x11eb041d5256e340
    End
    RecordSource ="tbl_LP_Disturbance"
    Caption ="fsub_LP_Lower_Canopy"
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
                    Left =120
                    Top =120
                    Width =1545
                    Height =240
                    FontWeight =700
                    Name ="Label13"
                    Caption ="Disturbances"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1860
                    Width =306
                    Height =306
                    Name ="ButtonFirst"
                    Caption ="Command21"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddada44dadad1dadaadad44adad11adaddada44dad111dada ,
                        0xadad44ad1111adaddada44d11111dadaadad44ad1111adaddada44dad111dada ,
                        0xadad44adad11adaddada44dadad1dadaadadadadadadadaddadadadadadadada ,
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
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="First Record"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2160
                    Width =306
                    Height =306
                    TabIndex =1
                    Name ="ButtonPrevious"
                    Caption ="Command22"
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
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="Previous Record"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2460
                    Width =306
                    Height =306
                    TabIndex =2
                    Name ="ButtonNext"
                    Caption ="Command23"
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
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="Next Record"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2760
                    Width =306
                    Height =306
                    TabIndex =3
                    Name ="ButtonLast"
                    Caption ="Command24"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddadad1dadad44adaadada11dada44daddadad111dad44ada ,
                        0xadada1111da44daddadad11111d44adaadada1111da44daddadad111dad44ada ,
                        0xadada11dada44daddadad1dadad44adaadadadadadadadaddadadadadadadada ,
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
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="Last Record"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3060
                    Width =306
                    Height =306
                    TabIndex =4
                    Name ="ButtonNew"
                    Caption ="Command25"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadadd1dadadadadadadaa11dadadada4adadd111dadad4d4d4da ,
                        0xa1111dadad444dadd11111da4444444aa1111dadad444dadd111dadad4d4d4da ,
                        0xa11dadadada4adadd1dadadadadadadaadadadadadadadaddadadadadadadada ,
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
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="Add Record"
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
                    Width =300
                    Height =239
                    ColumnWidth =2310
                    Name ="LC_ID"
                    ControlSource ="Disturbance_ID"
                    StatusBarText ="Unique record identifier - primary key"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =660
                    Top =60
                    Width =360
                    Height =239
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Intercept_ID"
                    ControlSource ="Intercept_ID"
                    StatusBarText ="Foreign key to tbl_LP_Intercept"
                End
                Begin ComboBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2745
                    Left =120
                    Top =60
                    Width =1560
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"10\""
                    Name ="Disturbance"
                    ControlSource ="Disturbance"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_LP_Disturbance.Dist_Code, tlu_LP_Disturbance.Disturbance FROM tlu_LP_"
                        "Disturbance; "
                    ColumnWidths ="495;2250"
                    AfterUpdate ="[Event Procedure]"
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

Private Sub Disturbance_AfterUpdate()
    If Not IsNull(DLookup("[Disturbance]", "tbl_LP_Disturbance", "[Intercept_ID] = '" & Me!Intercept_ID & "' AND [Disturbance] = '" & Me!Disturbance & "'")) Then
      MsgBox "This Disturbance is already recorded for this point."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
    End If
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
     
    On Error GoTo Err_Handler
    
    ' Create the GUID primary key value
    If IsNull(Me!Disturbance_ID) Then
        If GetDataType("tbl_LP_Lower_Canopy", "LC_ID") = dbText Then
            Me!Disturbance_ID = fxnGUIDGen
        End If
    End If

Exit_Procedure:

    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub


Private Sub ButtonFirst_Click()
On Error GoTo Err_ButtonFirst_Click


    DoCmd.GoToRecord , , acFirst

Exit_ButtonFirst_Click:
    Exit Sub

Err_ButtonFirst_Click:
    MsgBox Err.Description
    Resume Exit_ButtonFirst_Click
    
End Sub
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
Private Sub ButtonLast_Click()
On Error GoTo Err_ButtonLast_Click


    DoCmd.GoToRecord , , acLast

Exit_ButtonLast_Click:
    Exit Sub

Err_ButtonLast_Click:
    MsgBox Err.Description
    Resume Exit_ButtonLast_Click
    
End Sub
Private Sub ButtonNew_Click()
On Error GoTo Err_ButtonNew_Click


    DoCmd.GoToRecord , , acNewRec

Exit_ButtonNew_Click:
    Exit Sub

Err_ButtonNew_Click:
    MsgBox Err.Description
    Resume Exit_ButtonNew_Click
    
End Sub
