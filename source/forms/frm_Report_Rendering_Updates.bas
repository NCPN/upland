Version =21
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5400
    DatasheetFontHeight =11
    ItemSuffix =15
    Left =-1320
    Top =-11025
    Right =5865
    Bottom =-6375
    TimerInterval =100
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x4e112fa18ecce540
    End
    Caption ="Update on Progress"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnTimer ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin Section
            CanGrow = NotDefault
            Height =2595
            BackColor =15523798
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1635
                    Top =540
                    Width =2490
                    Height =435
                    FontSize =16
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label3"
                    Caption ="Processing Plot . . ."
                    GridlineColor =10921638
                    LayoutCachedLeft =1635
                    LayoutCachedTop =540
                    LayoutCachedWidth =4125
                    LayoutCachedHeight =975
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1440
                    Top =1200
                    Width =780
                    Height =480
                    FontSize =16
                    FontWeight =600
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtCurrRec"
                    GridlineColor =10921638

                    LayoutCachedLeft =1440
                    LayoutCachedTop =1200
                    LayoutCachedWidth =2220
                    LayoutCachedHeight =1680
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3600
                    Top =1200
                    Width =720
                    Height =480
                    FontSize =16
                    FontWeight =600
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtTotRec"
                    GridlineColor =10921638

                    LayoutCachedLeft =3600
                    LayoutCachedTop =1200
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =1680
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =2580
                            Top =1200
                            Width =480
                            Height =480
                            FontSize =16
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label8"
                            Caption ="of"
                            GridlineColor =10921638
                            LayoutCachedLeft =2580
                            LayoutCachedTop =1200
                            LayoutCachedWidth =3060
                            LayoutCachedHeight =1680
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2160
                    Top =2280
                    Width =480
                    Height =255
                    FontSize =9
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtSeconds"
                    FontName ="Tahoma"
                    GridlineColor =10921638

                    LayoutCachedLeft =2160
                    LayoutCachedTop =2280
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =2535
                    ThemeFontIndex =-1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2280
                            Width =2085
                            Height =255
                            FontSize =9
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblSeconds"
                            Caption ="Elapsed Time (seconds):"
                            FontName ="Tahoma"
                            GridlineColor =10921638
                            LayoutCachedLeft =60
                            LayoutCachedTop =2280
                            LayoutCachedWidth =2145
                            LayoutCachedHeight =2535
                            ThemeFontIndex =-1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4440
                    Top =2220
                    Width =600
                    Height =315
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtMinutes"
                    GridlineColor =10921638

                    LayoutCachedLeft =4440
                    LayoutCachedTop =2220
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =2535
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3180
                            Top =2220
                            Width =1050
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblMinutes"
                            Caption ="(minutes):"
                            GridlineColor =10921638
                            LayoutCachedLeft =3180
                            LayoutCachedTop =2220
                            LayoutCachedWidth =4230
                            LayoutCachedHeight =2535
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
Option Explicit

Private Sub Form_Load()
    Dim intPos As Integer
    Dim intLeft As Integer
    Dim dblRight As Double
    
   
    If Not IsNull(Me.OpenArgs) Then
    intPos = InStr(Me.OpenArgs, "|")
    intLeft = CInt(Left(Me.OpenArgs, intPos - 1))
    dblRight = CDbl(Mid(Me.OpenArgs, intPos + 1))
    
    Me!txtCurrRec = intLeft Mod 100
    Me!txtTotRec = Int(intLeft / 100)
    Me!txtSeconds = dblRight
    Me!txtMinutes = Round(dblRight / 60, 2)
    End If
End Sub

Private Sub Form_Timer()
   ' DoCmd.Close acForm, "frm_Report_Rendering_Updates"

End Sub
