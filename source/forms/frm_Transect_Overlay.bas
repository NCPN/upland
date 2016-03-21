Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8884
    DatasheetFontHeight =11
    ItemSuffix =2
    Left =7080
    Top =4440
    Right =15960
    Bottom =11985
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x6ab456d96fb4e440
    End
    Caption ="frm_Transect_Overlay"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnClick ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
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
        Begin Section
            Height =7560
            BackColor =-2147483607
            Name ="Detail"
            OnClick ="[Event Procedure]"
            AutoHeight =1
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Top =2220
                    Width =8640
                    Height =3600
                    FontSize =125
                    FontWeight =900
                    BorderColor =8355711
                    ForeColor =2366701
                    Name ="lblTransectNumber"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedTop =2220
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =5820
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    Left =330
                    Width =8085
                    Height =2220
                    FontSize =100
                    FontWeight =900
                    BorderColor =8355711
                    Name ="lblTransect"
                    Caption ="Transect"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =330
                    LayoutCachedWidth =8415
                    LayoutCachedHeight =2220
                    ForeTint =100.0
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

' =================================
' MODULE:       frm_Transect_Overlay
' Level:        Form module
' Version:      1.01
' Description:  data functions & procedures specific to transect overlay message
'
' Source/date:  Bonnie Campbell, 2/2/2016
' Adapted:      -
' Revisions:    BLC - 2/3/2016  - 1.00 - initial version
' =================================

' ---------------------------------
' SUB:          Form_Open
' Description:  Handles form opening actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 3, 2016 - for NCPN tools
' Revisions:
'   BLC, 2/3/2016  - initial version
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler

    Dim Transp As Long
    
    'close when no transect # is passed
    If IsNull(Me.OpenArgs) Or Me.OpenArgs = 0 Then
        DoCmd.Close
        GoTo Exit_Handler
    End If
    
    'set transect #
    lblTransectNumber.Caption = Me.OpenArgs
    
    'set background color -- RGB(0, 0, 0)
    'to set color RGB, see http://www.rapidtables.com/web/color/RGB_Color.htm for Red/Green/Blue values
    Transp = RGB(153, 255, 51)
     
    Me.Detail.backcolor = Transp
     
    Me.Painting = False
    'set background opacity
    SetFormOpacity Me, 0.9, Transp
    Me.Painting = True

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[Form_frm_Transect_Overlay])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_Click
' Description:  Handles form click actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 3, 2016 - for NCPN tools
' Revisions:
'   BLC, 2/3/2016  - initial version
' ---------------------------------
Private Sub Form_Click()
On Error GoTo Err_Handler

    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Click[Form_frm_Transect_Overlay])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Detail_Click
' Description:  Handles form detail click actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 3, 2016 - for NCPN tools
' Revisions:
'   BLC, 2/3/2016  - initial version
' ---------------------------------
Private Sub Detail_Click()
On Error GoTo Err_Handler

    DoCmd.Close
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Detail_Click[Form_frm_Transect_Overlay])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          lblTransect_Click
' Description:  Handles form transect label click actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 3, 2016 - for NCPN tools
' Revisions:
'   BLC, 2/3/2016  - initial version
' ---------------------------------
Private Sub lblTransect_Click()
On Error GoTo Err_Handler

    DoCmd.Close

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblTransect_Click[Form_frm_Transect_Overlay])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          lblTransectNumber_Click
' Description:  Handles form transect number label click actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 3, 2016 - for NCPN tools
' Revisions:
'   BLC, 2/3/2016  - initial version
' ---------------------------------
Private Sub lblTransectNumber_Click()
On Error GoTo Err_Handler

    DoCmd.Close

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblTransectNumber_Click[Form_frm_Transect_Overlay])"
    End Select
    Resume Exit_Handler
End Sub
