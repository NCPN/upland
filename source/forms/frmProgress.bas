Version =21
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =0
    AllowUpdating =2
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7371
    DatasheetFontHeight =10
    ItemSuffix =146
    Left =1935
    Top =825
    Right =9810
    Bottom =9510
    TimerInterval =250
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x337b15f8934ae240
    End
    Caption =" "
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x5103000034020000510300003402000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnTimer ="[Event Procedure]"
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            Width =850
            Height =850
        End
        Begin Line
            BorderLineStyle =0
            Width =1701
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="MS Sans Serif"
            BorderLineStyle =0
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
            Width =1701
            Height =1701
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin CustomControl
            SpecialEffect =2
            Width =4536
            Height =2835
        End
        Begin ToggleButton
            Width =283
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin Tab
            Width =5103
            Height =3402
            BorderLineStyle =0
        End
        Begin Page
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =7938
            BackColor =16776960
            Name ="Detail"
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =93
                    Left =29
                    Top =29
                    Width =7314
                    Height =7881
                    BackColor =16776960
                    BorderColor =8388608
                    Name ="boxOuter"
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =284
                    Top =227
                    Width =6804
                    Height =345
                    FontSize =12
                    FontWeight =700
                    BackColor =16776960
                    BorderColor =255
                    ForeColor =255
                    Name ="lblTitle"
                    Caption ="Please wait..."
                    FontName ="Arial"
                    OnClick ="[Event Procedure]"
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =1
                    Left =624
                    Top =680
                    Width =6464
                    Height =227
                    BackColor =16776960
                    BorderColor =255
                    ForeColor =255
                    Name ="lblStep00"
                    FontName ="Arial"
                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =223
                    Left =113
                    Top =113
                    Width =7144
                    Height =7711
                    BackColor =16776960
                    BorderColor =8388608
                    Name ="boxInner"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =624
                    Top =964
                    Width =6464
                    Height =227
                    BackColor =16776960
                    BorderColor =255
                    ForeColor =255
                    Name ="lblStep01"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =624
                    Top =1247
                    Width =6464
                    Height =227
                    BackColor =16776960
                    BorderColor =255
                    ForeColor =255
                    Name ="lblStep02"
                    FontName ="Arial"
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =284
                    Top =680
                    Width =284
                    Height =227
                    FontSize =12
                    FontWeight =700
                    BackColor =16776960
                    ForeColor =255
                    Name ="lblTick00"
                    Caption ="ü"
                    FontName ="Wingdings"
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =284
                    Top =964
                    Width =284
                    Height =227
                    FontSize =12
                    FontWeight =700
                    BackColor =16776960
                    ForeColor =255
                    Name ="lblTick01"
                    Caption ="ü"
                    FontName ="Wingdings"
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =284
                    Top =1247
                    Width =284
                    Height =227
                    FontSize =12
                    FontWeight =700
                    BackColor =16776960
                    ForeColor =255
                    Name ="lblTick02"
                    Caption ="ü"
                    FontName ="Wingdings"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =624
                    Top =1531
                    Width =6464
                    Height =227
                    BackColor =16776960
                    BorderColor =255
                    ForeColor =255
                    Name ="lblStep03"
                    FontName ="Arial"
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =284
                    Top =1531
                    Width =284
                    Height =227
                    FontSize =12
                    FontWeight =700
                    BackColor =16776960
                    ForeColor =255
                    Name ="lblTick03"
                    Caption ="ü"
                    FontName ="Wingdings"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =624
                    Top =1814
                    Width =6464
                    Height =227
                    BackColor =16776960
                    BorderColor =255
                    ForeColor =255
                    Name ="lblStep04"
                    FontName ="Arial"
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =284
                    Top =1814
                    Width =284
                    Height =227
                    FontSize =12
                    FontWeight =700
                    BackColor =16776960
                    ForeColor =255
                    Name ="lblTick04"
                    Caption ="ü"
                    FontName ="Wingdings"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =624
                    Top =2098
                    Width =6464
                    Height =227
                    BackColor =16776960
                    BorderColor =255
                    ForeColor =255
                    Name ="lblStep05"
                    FontName ="Arial"
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =284
                    Top =2098
                    Width =284
                    Height =227
                    FontSize =12
                    FontWeight =700
                    BackColor =16776960
                    ForeColor =255
                    Name ="lblTick05"
                    Caption ="ü"
                    FontName ="Wingdings"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =624
                    Top =2381
                    Width =6464
                    Height =227
                    BackColor =16776960
                    BorderColor =255
                    ForeColor =255
                    Name ="lblStep06"
                    FontName ="Arial"
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =284
                    Top =2381
                    Width =284
                    Height =227
                    FontSize =12
                    FontWeight =700
                    BackColor =16776960
                    ForeColor =255
                    Name ="lblTick06"
                    Caption ="ü"
                    FontName ="Wingdings"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =624
                    Top =2665
                    Width =6464
                    Height =227
                    BackColor =16776960
                    BorderColor =255
                    ForeColor =255
                    Name ="lblStep07"
                    FontName ="Arial"
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =284
                    Top =2665
                    Width =284
                    Height =227
                    FontSize =12
                    FontWeight =700
                    BackColor =16776960
                    ForeColor =255
                    Name ="lblTick07"
                    Caption ="ü"
                    FontName ="Wingdings"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =624
                    Top =2948
                    Width =6464
                    Height =227
                    BackColor =16776960
                    BorderColor =255
                    ForeColor =255
                    Name ="lblStep08"
                    FontName ="Arial"
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =284
                    Top =2948
                    Width =284
                    Height =227
                    FontSize =12
                    FontWeight =700
                    BackColor =16776960
                    ForeColor =255
                    Name ="lblTick08"
                    Caption ="ü"
                    FontName ="Wingdings"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =624
                    Top =3232
                    Width =6464
                    Height =227
                    BackColor =16776960
                    BorderColor =255
                    ForeColor =255
                    Name ="lblStep09"
                    FontName ="Arial"
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =284
                    Top =3232
                    Width =284
                    Height =227
                    FontSize =12
                    FontWeight =700
                    BackColor =16776960
                    ForeColor =255
                    Name ="lblTick09"
                    Caption ="ü"
                    FontName ="Wingdings"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =624
                    Top =3515
                    Width =6464
                    Height =227
                    BackColor =16776960
                    BorderColor =255
                    ForeColor =255
                    Name ="lblStep10"
                    FontName ="Arial"
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =284
                    Top =3515
                    Width =284
                    Height =227
                    FontSize =12
                    FontWeight =700
                    BackColor =16776960
                    ForeColor =255
                    Name ="lblTick10"
                    Caption ="ü"
                    FontName ="Wingdings"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =624
                    Top =3799
                    Width =6464
                    Height =227
                    BackColor =16776960
                    BorderColor =255
                    ForeColor =255
                    Name ="lblStep11"
                    FontName ="Arial"
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =284
                    Top =3799
                    Width =284
                    Height =227
                    FontSize =12
                    FontWeight =700
                    BackColor =16776960
                    ForeColor =255
                    Name ="lblTick11"
                    Caption ="ü"
                    FontName ="Wingdings"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =624
                    Top =4082
                    Width =6464
                    Height =227
                    BackColor =16776960
                    BorderColor =255
                    ForeColor =255
                    Name ="lblStep12"
                    FontName ="Arial"
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =284
                    Top =4082
                    Width =284
                    Height =227
                    FontSize =12
                    FontWeight =700
                    BackColor =16776960
                    ForeColor =255
                    Name ="lblTick12"
                    Caption ="ü"
                    FontName ="Wingdings"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =624
                    Top =4366
                    Width =6464
                    Height =227
                    BackColor =16776960
                    BorderColor =255
                    ForeColor =255
                    Name ="lblStep13"
                    FontName ="Arial"
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =284
                    Top =4366
                    Width =284
                    Height =227
                    FontSize =12
                    FontWeight =700
                    BackColor =16776960
                    ForeColor =255
                    Name ="lblTick13"
                    Caption ="ü"
                    FontName ="Wingdings"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =624
                    Top =4649
                    Width =6464
                    Height =227
                    BackColor =16776960
                    BorderColor =255
                    ForeColor =255
                    Name ="lblStep14"
                    FontName ="Arial"
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =284
                    Top =4649
                    Width =284
                    Height =227
                    FontSize =12
                    FontWeight =700
                    BackColor =16776960
                    ForeColor =255
                    Name ="lblTick14"
                    Caption ="ü"
                    FontName ="Wingdings"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =624
                    Top =4933
                    Width =6464
                    Height =227
                    BackColor =16776960
                    BorderColor =255
                    ForeColor =255
                    Name ="lblStep15"
                    FontName ="Arial"
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =284
                    Top =4933
                    Width =284
                    Height =227
                    FontSize =12
                    FontWeight =700
                    BackColor =16776960
                    ForeColor =255
                    Name ="lblTick15"
                    Caption ="ü"
                    FontName ="Wingdings"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =624
                    Top =5216
                    Width =6464
                    Height =227
                    BackColor =16776960
                    BorderColor =255
                    ForeColor =255
                    Name ="lblStep16"
                    FontName ="Arial"
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =284
                    Top =5216
                    Width =284
                    Height =227
                    FontSize =12
                    FontWeight =700
                    BackColor =16776960
                    ForeColor =255
                    Name ="lblTick16"
                    Caption ="ü"
                    FontName ="Wingdings"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =624
                    Top =5500
                    Width =6464
                    Height =227
                    BackColor =16776960
                    BorderColor =255
                    ForeColor =255
                    Name ="lblStep17"
                    FontName ="Arial"
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =284
                    Top =5500
                    Width =284
                    Height =227
                    FontSize =12
                    FontWeight =700
                    BackColor =16776960
                    ForeColor =255
                    Name ="lblTick17"
                    Caption ="ü"
                    FontName ="Wingdings"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =624
                    Top =5783
                    Width =6464
                    Height =227
                    BackColor =16776960
                    BorderColor =255
                    ForeColor =255
                    Name ="lblStep18"
                    FontName ="Arial"
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =284
                    Top =5783
                    Width =284
                    Height =227
                    FontSize =12
                    FontWeight =700
                    BackColor =16776960
                    ForeColor =255
                    Name ="lblTick18"
                    Caption ="ü"
                    FontName ="Wingdings"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =624
                    Top =6067
                    Width =6464
                    Height =227
                    BackColor =16776960
                    BorderColor =255
                    ForeColor =255
                    Name ="lblStep19"
                    FontName ="Arial"
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =284
                    Top =6067
                    Width =284
                    Height =227
                    FontSize =12
                    FontWeight =700
                    BackColor =16776960
                    ForeColor =255
                    Name ="lblTick19"
                    Caption ="ü"
                    FontName ="Wingdings"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =624
                    Top =6350
                    Width =6464
                    Height =227
                    BackColor =16776960
                    BorderColor =255
                    ForeColor =255
                    Name ="lblStep20"
                    FontName ="Arial"
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =284
                    Top =6350
                    Width =284
                    Height =227
                    FontSize =12
                    FontWeight =700
                    BackColor =16776960
                    ForeColor =255
                    Name ="lblTick20"
                    Caption ="ü"
                    FontName ="Wingdings"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =624
                    Top =6634
                    Width =6464
                    Height =227
                    BackColor =16776960
                    BorderColor =255
                    ForeColor =255
                    Name ="lblStep21"
                    FontName ="Arial"
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =284
                    Top =6634
                    Width =284
                    Height =227
                    FontSize =12
                    FontWeight =700
                    BackColor =16776960
                    ForeColor =255
                    Name ="lblTick21"
                    Caption ="ü"
                    FontName ="Wingdings"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =624
                    Top =6917
                    Width =6464
                    Height =227
                    BackColor =16776960
                    BorderColor =255
                    ForeColor =255
                    Name ="lblStep22"
                    FontName ="Arial"
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =284
                    Top =6917
                    Width =284
                    Height =227
                    FontSize =12
                    FontWeight =700
                    BackColor =16776960
                    ForeColor =255
                    Name ="lblTick22"
                    Caption ="ü"
                    FontName ="Wingdings"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =624
                    Top =7201
                    Width =6464
                    Height =227
                    BackColor =16776960
                    BorderColor =255
                    ForeColor =255
                    Name ="lblStep23"
                    FontName ="Arial"
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =284
                    Top =7201
                    Width =284
                    Height =227
                    FontSize =12
                    FontWeight =700
                    BackColor =16776960
                    ForeColor =255
                    Name ="lblTick23"
                    Caption ="ü"
                    FontName ="Wingdings"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =624
                    Top =7484
                    Width =6464
                    Height =227
                    BackColor =16776960
                    BorderColor =255
                    ForeColor =255
                    Name ="lblStep24"
                    FontName ="Arial"
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextFontFamily =2
                    Left =284
                    Top =7484
                    Width =284
                    Height =227
                    FontSize =12
                    FontWeight =700
                    BackColor =16776960
                    ForeColor =255
                    Name ="lblTick24"
                    Caption ="ü"
                    FontName ="Wingdings"
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
'The frmProgress form is designed to stay visible for about 2" after it expires.
'However, the operator can cancel the delay if he clicks on the form's title.
'11/5/2006  Allows ten entries.
'15/5/2006  Resize form to handle only the number of entries required.
'           This cannot work as the form size itself never changes on screen.
'18/08/2008 Tried again using Access 2003

Private Const conMaxStep As Integer = 24    'Steps = conMaxSteps + 1 (From 0)
Private Const conDelSecs As Integer = 2     'Default delay in secs
Private Const conProgSep As String = "~"    'Separator character within strMsgs
Private Const conCross As Long = &HFB       'Wingdings cross
Private Const conTick As Long = &HFC        'Wingdings tick
Private Const conCM As Long = &H238         'Centimeter

'intPeriod 1/4"s counted after completion; intDelay 1/4"s to count;
'intLastStep is the last step used on the form
Private intPeriod As Integer, intDelay As Integer, intLastStep As Integer
Private lblTicks(0 To conMaxStep) As Label, lblSteps(0 To conMaxStep) As Label

Private Sub Form_Open(Cancel As Integer)
    Dim strStep As String
    Dim ctlThis As Control

    'Assign all labels to the arrays.  Ignore any failures.
    On Error Resume Next
    For Each ctlThis In Controls
        strStep = Right(ctlThis.Name, 2)
        Select Case Left(ctlThis.Name, 7)
        Case "lblTick"
            Set lblTicks(CInt(strStep)) = ctlThis
        Case "lblStep"
            Set lblSteps(CInt(strStep)) = ctlThis
        End Select
    Next ctlThis
    On Error GoTo 0
End Sub

'intStep = 0            Reset all and set up captions
'intStep = Positive     Operate on relevant (intStep-1) line of the display
'intStep = Negative     Close Progress form after processing -intStep

'  intState = 0         Not started yet - visible / dim
'  intState = 1         In progress     - visible / bold
'  intState = 2         Completed       - visible / ticked
'  intState = 3         Hidden          - visible / dim / crossed
'  intState = 4         In progress for intStep - Completed for previous step
'  intState = 5         In progress for intStep - Hidden for previous step
Public Sub SetStep(ByVal intStep As Integer, _
                   Optional ByVal intState As Integer = -1, _
                   Optional ByRef strMsgs As String = "", _
                   Optional ByVal intDelSecs As Integer = -1, _
                   Optional ByVal dblCM As Double = 0)
    Dim intIdx As Integer, intTop As Integer
    Dim lngSize As Long
    Dim blnClose As Boolean

    'Cancel any pending close (see Timer code)
    intPeriod = 0
    'Default intDelSecs if not set
    If intDelSecs = -1 Then intDelSecs = conDelSecs
    'Default intState depending on intStep
    If intState = -1 Then
        Select Case intStep
        Case 0              'Open - Default = 1 In progress
            intState = 1
        Case Is > 0         'Change step - Default = 4 Complete & In progress
            intState = 4
        Case Is < 0         'Close - Default = 2 Complete
            intState = 2
        End Select
    End If
    Select Case Abs(intStep)
    Case 0      'Reset all and set up captions
        intDelay = intDelSecs * 4 + Sgn(intDelSecs)
        'find number of elements in strMsgs
        intTop = UBound(Split(strMsgs, conProgSep))
        If intTop > conMaxStep Then intTop = conMaxStep
        For intIdx = 0 To conMaxStep
            If intIdx > intTop Then
                lblTicks(intIdx).Visible = False
                lblSteps(intIdx).Visible = False
            Else
                lblSteps(intIdx).Visible = True
                lblSteps(intIdx).Caption = Split(strMsgs, conProgSep)(intIdx)
                Call SetState(intStep:=intIdx, _
                              intState:=IIf(intIdx = 0, intState, 0))
            End If
        Next intIdx
        'Resize form depending on # of lines used and lngWidth passed
        With Me
            If intTop < conMaxStep Then
                lngSize = (conMaxStep - intTop) * conCM / 2
                .boxInner.Height = .boxInner.Height - lngSize
                .boxOuter.Height = .boxOuter.Height - lngSize
                .InsideHeight = .InsideHeight - lngSize
                'Following line depends on Access 2003
                Call .Move(Left:=.WindowLeft, Top:=.WindowTop + lngSize / 2)
            End If
            If dblCM > 0 Then
                lngSize = dblCM * conCM
                .lblTitle.width = .lblTitle.width - lngSize
                .boxInner.width = .boxInner.width - lngSize
                .boxOuter.width = .boxOuter.width - lngSize
                .InsideWidth = .InsideWidth - lngSize
                For intTop = intTop To 0 Step -1
                    lblSteps(intTop).width = lblSteps(intTop).width - lngSize
                Next intTop
                'Following line depends on Access 2003
                Call .Move(Left:=.WindowLeft + lngSize / 2)
            End If
        End With
    Case 1 To conMaxStep + 1
        Call SetState(Abs(intStep) - 1, intState)
    End Select
    If intStep < 0 Then     'Close Progress form
        If intDelay = 0 Then Call CloseMe
        'Otherwise start timer
        intPeriod = 1
    End If
    'Update the screen
    DoEvents
End Sub

Private Sub SetState(intStep As Integer, intState As Integer)
    lblTicks(intStep).Caption = Chr(conTick)
    lblSteps(intStep).FontBold = False
    Select Case intState
    Case 0          'Not started yet (dim)
        lblTicks(intStep).Visible = False
        lblSteps(intStep).ForeColor = vbBlue
    Case 1, 4, 5    'In progress (bold)
        lblTicks(intStep).Visible = False
        lblSteps(intStep).ForeColor = vbRed
        lblSteps(intStep).FontBold = True
        If intState > 3 And intStep > 0 Then _
            Call SetState(intStep:=intStep - 1, intState:=intState - 2)
    Case 2      'Completed (Tick)
        lblTicks(intStep).Visible = True
        lblSteps(intStep).ForeColor = vbRed
    Case 3      'Hidden (dim / cross)
        lblTicks(intStep).Caption = Chr(conCross)
        lblTicks(intStep).Visible = True
        lblSteps(intStep).ForeColor = vbBlue
    End Select
    'Always bring frmProgress to front when updating
    Call DoCmd.SelectObject(objecttype:=acForm, ObjectName:=Me.Name)
    'Update the screen
    DoEvents
End Sub

Private Sub lblTitle_Click()
    If intPeriod > 0 Then Call CloseMe
End Sub

Private Sub Form_Timer()
    Select Case intPeriod
    Case 0
        Exit Sub
    Case Is < intDelay
Debug.Print intPeriod, Now
        intPeriod = intPeriod + 1
        Call DoCmd.SelectObject(objecttype:=acForm, ObjectName:=Me.Name)
    Case Else
        Call CloseMe
    End Select
End Sub

Private Sub CloseMe()
    Call DoCmd.Close
End Sub
