Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =127
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7020
    DatasheetFontHeight =10
    ItemSuffix =18
    Left =3480
    Top =2430
    Right =10545
    Bottom =7905
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xf2523ab75013e340
    End
    RecordSource ="tlu_Plant_List"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
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
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin Section
            Height =5280
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2070
                    Top =180
                    Width =2865
                    Height =420
                    FontSize =14
                    Name ="Label0"
                    Caption ="Add Unknown Species"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3060
                    Top =4440
                    Width =1035
                    Height =405
                    Name ="ButtonClose"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5040
                    Top =660
                    Width =1500
                    TabIndex =1
                    ForeColor =255
                    Name ="ButtonDelete"
                    Caption ="Delete This Entry"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1320
                    Top =1080
                    Width =780
                    ColumnWidth =810
                    TabIndex =2
                    Name ="Symbol"
                    ControlSource ="Symbol"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =360
                            Top =1080
                            Width =960
                            Height =240
                            FontWeight =700
                            Name ="Label14"
                            Caption ="Plant Code"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2460
                    Top =2160
                    Width =2100
                    ColumnWidth =3045
                    TabIndex =3
                    Name ="Scientific_Name"
                    ControlSource ="Scientific name"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =2160
                            Width =2070
                            Height =240
                            FontWeight =700
                            Name ="Label15"
                            Caption ="Unknown Species Name"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =360
                    Top =1380
                    Width =5040
                    Height =420
                    Name ="Label17"
                    Caption ="Enter a unique 4-6 digit code for this plant to store in the database.  I.E. CAN"
                        "Y01"
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

Private Sub Form_Load()

  If Not IsNull(Me.OpenArgs) Then
    DoCmd.GoToRecord , , acNewRec
    Me!Scientific_name = Me.OpenArgs
  End If
  
End Sub
Private Sub ButtonClose_Click()
On Error GoTo Err_ButtonClose_Click


    DoCmd.Close

Exit_ButtonClose_Click:
    Exit Sub

Err_ButtonClose_Click:
    MsgBox Err.Description
    Resume Exit_ButtonClose_Click
    
End Sub
Private Sub ButtonDelete_Click()
On Error GoTo Err_ButtonDelete_Click


    DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
    DoCmd.DoMenuItem acFormBar, acEditMenu, 6, , acMenuVer70

Exit_ButtonDelete_Click:
    Exit Sub

Err_ButtonDelete_Click:
    MsgBox Err.Description
    Resume Exit_ButtonDelete_Click
    
End Sub
