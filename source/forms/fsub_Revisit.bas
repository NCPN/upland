﻿Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =127
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =7380
    DatasheetFontHeight =9
    ItemSuffix =16
    Left =5244
    Top =1716
    Right =12384
    Bottom =3000
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x1dcf8f960f51e340
    End
    RecordSource ="tbl_Events"
    Caption ="fsub_Revisit"
    OnCurrent ="[Event Procedure]"
    BeforeInsert ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
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
            Height =1560
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
                    ColumnWidth =2310
                    Name ="Event_ID"
                    ControlSource ="Event_ID"
                    StatusBarText ="M. Event identifier (Event_ID)"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =900
                    Top =60
                    Width =540
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Location_ID"
                    ControlSource ="Location_ID"
                    StatusBarText ="M. Link to tbl_Locations (Loc_ID)"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =2940
                    Top =60
                    Width =1200
                    ColumnWidth =900
                    TabIndex =2
                    Name ="version_key_number"
                    ControlSource ="version_key_number"
                    StatusBarText ="Master protocol version key"

                End
                Begin TextBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =1320
                    Top =180
                    Width =975
                    Height =255
                    ColumnWidth =1035
                    TabIndex =3
                    Name ="Start_Date"
                    ControlSource ="Start_Date"
                    Format ="Short Date"
                    StatusBarText ="M. Starting date for the event (Start_Date) - date of revisit."
                    InputMask ="99/99/0000;0;_"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =247
                            TextAlign =3
                            Left =120
                            Top =180
                            Width =1140
                            Height =240
                            FontWeight =700
                            Name ="Start_Date_Label"
                            Caption ="Revisit Date"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =120
                    Top =840
                    Width =7080
                    Height =540
                    TabIndex =4
                    Name ="Comments"
                    ControlSource ="Comments"
                    StatusBarText ="Plot revisit comments."

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =120
                            Top =600
                            Width =1140
                            Height =240
                            FontWeight =700
                            Name ="Comments_Label"
                            Caption ="Comments"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =2880
                    Left =3960
                    Top =180
                    Width =1620
                    TabIndex =5
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Observer"
                    ControlSource ="Observer"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Contacts.Contact_ID, tlu_Contacts.Last_Name, tlu_Contacts.First_Name "
                        "FROM tlu_Contacts WHERE (((tlu_Contacts.Active)=1)); "
                    ColumnWidths ="0;1440;1440"

                    LayoutCachedLeft =3960
                    LayoutCachedTop =180
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =420
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =3060
                            Top =180
                            Width =840
                            Height =245
                            FontWeight =700
                            Name ="Observer_Label"
                            Caption ="Observer"
                            LayoutCachedLeft =3060
                            LayoutCachedTop =180
                            LayoutCachedWidth =3900
                            LayoutCachedHeight =425
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6240
                    Top =480
                    Width =660
                    Height =180
                    TabIndex =6
                    Name ="Event_Save"

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
        Dim Versions As DAO.Recordset
        Dim strSQL As String
        
    On Error GoTo Err_Handler
    
    ' Set master version number on event record
    Set db = CurrentDb
    strSQL = "SELECT [version_key_number] FROM [tbl_master_version] ORDER BY [version_key_number] DESC"
    Set Versions = db.OpenRecordset(strSQL)
    Versions.MoveFirst
    Me![version_key_number] = Versions![version_key_number]
    Versions.Close

    ' Create the GUID primary key value
    If IsNull(Me!Event_ID) Then
        If GetDataType("tbl_Events", "Event_ID") = dbText Then
            Me.Event_ID = fxnGUIDGen
        End If
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

Private Sub Form_Current()
  If Not IsNull(Me!Event_Save) Then
        Me.RecordsetClone.FindFirst "Event_ID = '" & Me!Event_Save & "'"    ' Go back to new event
        If Not Me.RecordsetClone.NoMatch Then
          Me.Bookmark = Me.RecordsetClone.Bookmark
        End If
  Else
    DoCmd.GoToRecord , , acNewRec
  End If
End Sub

Private Sub Form_Open(Cancel As Integer)
'  DoCmd.GoToRecord , , acNewRec
  
End Sub
