﻿Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    FilterOn = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    TabularFamily =127
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4680
    DatasheetFontHeight =9
    ItemSuffix =15
    Left =9165
    Top =3825
    Right =13845
    Bottom =7110
    DatasheetGridlinesColor =12632256
    Filter ="[Location_ID]='20100715151932-709037899.971008'"
    RecSrcDt = Begin
        0x5bd611c7ad13e340
    End
    RecordSource ="qfrm_Visit_Date"
    Caption ="Select a Visit"
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
            Height =1140
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =180
                    Top =900
                    Width =1035
                    Height =240
                    FontWeight =700
                    Name ="Start_Date_Label"
                    Caption ="Visit Date"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =120
                    Top =120
                    Width =540
                    Height =240
                    Name ="Unit_Code_Label"
                    Caption ="Park"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =120
                    Top =420
                    Width =600
                    Height =240
                    Name ="Plot_ID_Label"
                    Caption ="Plot ID"
                    Tag ="DetachedLabel"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3360
                    Top =780
                    Width =1020
                    Height =300
                    Name ="ButtonClose"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =720
                    Top =120
                    Width =600
                    Height =255
                    ColumnWidth =540
                    TabIndex =1
                    Name ="Unit_Code"
                    ControlSource ="Unit_Code"
                    StatusBarText ="Park Code."
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =780
                    Top =420
                    Width =600
                    Height =255
                    ColumnWidth =600
                    TabIndex =2
                    Name ="Plot_ID"
                    ControlSource ="Plot_ID"
                    StatusBarText ="Plot identifier"
                End
                Begin Label
                    OverlapFlags =85
                    Left =1680
                    Top =240
                    Width =2640
                    Height =360
                    FontSize =12
                    FontWeight =700
                    Name ="Label12"
                    Caption ="Select a Visit to Edit"
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
                    Width =660
                    Height =255
                    ColumnWidth =2310
                    Name ="Event_ID"
                    ControlSource ="Event_ID"
                    StatusBarText ="M. Event identifier (Event_ID)"
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =840
                    Top =60
                    Width =690
                    Height =255
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Location_ID"
                    ControlSource ="Location_ID"
                    StatusBarText ="M. Link to tbl_Locations (Loc_ID)"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    BackStyle =0
                    IMESentenceMode =3
                    Left =180
                    Top =60
                    Width =1035
                    Height =255
                    ColumnWidth =1035
                    TabIndex =2
                    Name ="Start_Date"
                    ControlSource ="Start_Date"
                    Format ="Short Date"
                    StatusBarText ="M. Starting date for the event (Start_Date)"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1860
                    Top =60
                    Width =1020
                    Height =300
                    TabIndex =3
                    Name ="ButtonEdit"
                    Caption ="Edit Visit"
                    OnClick ="[Event Procedure]"
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
Private Sub ButtonEdit_Click()
    On Error GoTo Err_Handler

    Dim strCriteriaLoc As String
    Dim strCriteriaEvent As String

        strCriteriaLoc = GetCriteriaString("[Location_ID]=", "tbl_Locations", "Location_ID", Me.name, "Location_ID")
        strCriteriaEvent = GetCriteriaString("[Event_ID]=", "tbl_Events", "Event_ID", Me.name, "Event_ID")
        ' Filter by location and event
        DoCmd.OpenForm "frm_Data_Entry", , , strCriteriaLoc & " AND " & strCriteriaEvent, , , strCriteriaEvent
        DoCmd.Close acForm, "frm_Visit_Date"
        DoCmd.SelectObject acForm, "frm_Data_Entry"
Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub
