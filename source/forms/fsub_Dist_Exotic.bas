Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    DatasheetFontHeight =9
    ItemSuffix =31
    Left =2145
    Top =4005
    Right =9630
    Bottom =6720
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x718d0abeb2a7e340
    End
    RecordSource ="tbl_Dist_Exotic"
    Caption ="fsub_LP_Belt_Shrub"
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
            Height =840
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =240
                    Top =600
                    Width =1335
                    Height =240
                    FontWeight =700
                    Name ="Species_Label"
                    Caption ="Species Name"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2700
                    Top =600
                    Width =585
                    Height =240
                    FontWeight =700
                    Name ="Total_Label"
                    Caption ="Notes"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1440
                    Top =60
                    Width =2865
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label23"
                    Caption ="Exotic Species"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5340
                    Top =60
                    Width =1545
                    Height =300
                    Name ="ButtonMaster"
                    Caption ="Master Species"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5340
                    Top =480
                    Width =1545
                    Height =300
                    TabIndex =1
                    Name ="ButtonUnknown"
                    Caption ="Unknown Species"
                    OnClick ="[Event Procedure]"
                End
            End
        End
        Begin Section
            Height =420
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =360
                    Height =255
                    ColumnWidth =2310
                    Name ="Shrub_ID"
                    ControlSource ="Exotic_ID"
                    StatusBarText ="Unique record identifier - primary key"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =300
                    Top =60
                    Width =300
                    Height =255
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Transect_ID"
                    ControlSource ="Transect_ID"
                    StatusBarText ="Foreign key to tbl_Canopy_Transect"
                End
                Begin ComboBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =6480
                    Left =180
                    Top =60
                    Width =2304
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="Species"
                    ControlSource ="Species"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qryU_Top_Canopy.Master_PLANT_Code, qryU_Top_Canopy.LU_Code, qryU_Top_Cano"
                        "py.Utah_Species FROM qryU_Top_Canopy WHERE (((qryU_Top_Canopy.Utah_Species) Is N"
                        "ot Null)) ORDER BY qryU_Top_Canopy.LU_Code; "
                    ColumnWidths ="0;2160;4320"
                    BeforeUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6240
                    Top =60
                    Width =705
                    Height =300
                    TabIndex =5
                    ForeColor =255
                    Name ="ButtonDelete"
                    Caption ="Delete"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2700
                    Top =60
                    Width =3000
                    Height =243
                    TabIndex =3
                    Name ="Notes"
                    ControlSource ="Notes"
                    StatusBarText ="Notes on species"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5760
                    Width =306
                    Height =306
                    TabIndex =4
                    Name ="ButtonZoom"
                    Caption ="Zoom Caption"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x280000001000000010000000010004000000000080000000c40e0000c40e0000 ,
                        0x1000000000000000000000000000800000800000008080008000000080008000 ,
                        0x80800000c0c0c000808080000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00666666666666446666666666666474466666666666474446 ,
                        0x666666666474446666660000474446666600777f8444666660877777f8086666 ,
                        0x607770777f066666077770777770666607777077777066660700000007706666 ,
                        0x07f770777770666660ff707777066666608ff077780666666600777700666666 ,
                        0x6666000066666666
                    End
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0xc0c0c00080808000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="Zoom Caption"
                    Picture ="C:\\arcgis\\arcexe9x\\odetools\\Bitmaps\\zoomin.bmp"
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

Private Sub Button_Master_Species_Click()
On Error GoTo Err_Button_Master_Species_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim strOpenArg As String

    strOpenArg = "fsub_LP_Exotic"
    stDocName = "frm_Master_Species"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , , strOpenArg

Exit_Button_Master_Species_Click:
    Exit Sub

Err_Button_Master_Species_Click:
    MsgBox Err.Description
    Resume Exit_Button_Master_Species_Click
 
End Sub

Private Sub ButtonZoom_Click()
  Me!Notes.SetFocus
  SendKeys ("+{F2}")
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
  
    ' Create the GUID primary key value
    If IsNull(Me!Exotic_ID) Then
        If GetDataType("tbl_LP_Exotic", "Exotic_ID") = dbText Then
            Me.Exotic_ID = fxnGUIDGen
        End If
    End If

End Sub

Private Sub Species_BeforeUpdate(Cancel As Integer)
    If Not IsNull(DLookup("[Exotic_ID]", "tbl_Dist_Exotic", "[Impact_ID] = '" & Me!Impact_ID & "' AND [Species] = '" & Me!Species & "'")) Then
      MsgBox "This species is already recorded for this transect."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
      Me.Undo
    End If
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

Private Sub Species_GotFocus()

    If IsNull(Me.Parent!Visit_Date) Then    ' If they didn't bother to enter a date, default to event date.
      Me.Parent!Visit_Date = Me.Parent.Parent!Start_Date
      Me.Parent.Refresh   ' Force save of transect record
    End If
   
End Sub
Private Sub ButtonDelete_Click()
On Error GoTo Err_ButtonDelete_Click

  Dim intReply As Integer
  
  intReply = MsgBox("Are you sure you want to delete this record?", vbYesNo, "Delete Record")
    If intReply = vbYes Then
      DoCmd.SetWarnings False
      DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
      DoCmd.DoMenuItem acFormBar, acEditMenu, 6, , acMenuVer70
      DoCmd.SetWarnings True
      Me.Requery
    End If

Exit_ButtonDelete_Click:
    Exit Sub

Err_ButtonDelete_Click:
    MsgBox Err.Description
    Resume Exit_ButtonDelete_Click
    
End Sub
Private Sub ButtonMaster_Click()
On Error GoTo Err_ButtonMaster_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Master_Species"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_ButtonMaster_Click:
    Exit Sub

Err_ButtonMaster_Click:
    MsgBox Err.Description
    Resume Exit_ButtonMaster_Click
    
End Sub
