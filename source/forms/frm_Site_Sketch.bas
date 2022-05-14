Version =21
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    DividingLines = NotDefault
    FilterOn = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =31080
    DatasheetFontHeight =10
    ItemSuffix =27
    Left =4035
    Top =3030
    Right =8460
    Bottom =7800
    DatasheetGridlinesColor =12632256
    OnUnload ="[Event Procedure]"
    Filter ="[Impact_ID]='20090417114512-579518616.199493'"
    RecSrcDt = Begin
        0x4e04d3557112e340
    End
    RecordSource ="tbl_Site_Impact"
    Caption ="StoreLoadJpeg"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =255
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
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
            ForeColor =-2147483630
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin OptionButton
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
            BorderLineStyle =0
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =660
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =180
                    Width =1512
                    Height =600
                    Name ="CmdLoad"
                    Caption ="Load Jpeg from disk"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1920
                    Width =1512
                    Height =600
                    TabIndex =1
                    Name ="cmdSaveBlob"
                    Caption ="Save this map in site record"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3660
                    Width =1512
                    Height =600
                    TabIndex =2
                    Name ="cmdLoadBlob"
                    Caption ="Load site map from this record"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =87
                    Left =12240
                    Top =240
                    Width =900
                    Height =300
                    TabIndex =3
                    Name ="txtDisplay"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =12240
                            Width =492
                            Height =228
                            Name ="Label4"
                            Caption ="Text3:"
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =7920
                    Top =180
                    Width =600
                    TabIndex =4
                    Name ="cmd1-8"
                    Caption ="Size"
                    EventProcPrefix ="cmd1_8"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin BoundObjectFrame
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =13320
                    Top =300
                    Width =900
                    Height =180
                    ColumnWidth =4704
                    TabIndex =5
                    Name ="Site_Sketch"
                    ControlSource ="Site_Sketch"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13260
                            Top =60
                            Width =960
                            Height =228
                            Name ="Label8"
                            Caption ="MyBLOB:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5400
                    Width =1512
                    Height =600
                    TabIndex =6
                    Name ="cmdSaveJpeg"
                    Caption ="Save this image as Jpeg disk file"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =10440
                    Top =120
                    Width =1560
                    Height =480
                    TabIndex =7
                    Name ="cmdLoadImageCtl"
                    Caption ="Load supported Picture types into Image control"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin OptionGroup
                    OverlapFlags =215
                    Left =8580
                    Top =60
                    Width =2874
                    Height =544
                    TabIndex =8
                    Name ="FrameSize"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"

                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =8700
                            Width =864
                            Height =228
                            BackColor =-2147483633
                            Name ="Label16"
                            Caption ="Image Size"
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =8766
                            Top =298
                            OptionValue =1
                            Name ="Option18"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =8996
                                    Top =270
                                    Width =300
                                    Height =228
                                    Name ="Label19"
                                    Caption ="1x"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =9420
                            Top =304
                            OptionValue =2
                            Name ="Option20"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =9650
                                    Top =276
                                    Width =372
                                    Height =228
                                    Name ="Label21"
                                    Caption ="1/2"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =10080
                            Top =304
                            OptionValue =4
                            Name ="Option22"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =10308
                                    Top =276
                                    Width =372
                                    Height =228
                                    Name ="Label23"
                                    Caption ="1/4"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =215
                            Left =10800
                            Top =304
                            OptionValue =8
                            Name ="Option24"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =11030
                                    Top =276
                                    Width =312
                                    Height =228
                                    Name ="Label25"
                                    Caption ="1/8"
                                End
                            End
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =7140
                    Width =1200
                    Height =600
                    TabIndex =9
                    Name ="ButtonClose"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
            End
        End
        Begin Section
            Height =29760
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Image
                    SpecialEffect =2
                    PictureAlignment =0
                    Left =120
                    Top =60
                    Width =28800
                    Height =28800
                    Name ="Image0"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                End
            End
        End
        Begin FormFooter
            Visible = NotDefault
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
'DEVELOPED AND TESTED UNDER MICROSOFT ACCESS 97 VBA ONLY
'
'Copyright: Stephen Lebans - Lebans Holdings 1999 Ltd.
'Copyright: Intel Corporation

'Distribution:
' You are not allowed to redistribute this code in any fashion,
' whether in print or electronic media. You may though install the source
' code on any machine running this application or your application
' developed with this source.
' Plain and simple you are free to use this source within your own
' applications without cost or obligation, other that keeping
' the copyright notices intact. You may not resell this source code
' by itself or as part of a collection.
'
' This source may be downloaded from:
' www.lebans.com
'
'
'
'Name:      StoreLoadJpeg
'
'Version:   1.0
'
'Purpose:
'
' 1)Use the Intel JPEG libary to load and display Jpeg files
' in an Access Database. Jpeg files are stored in their original compressed
' format within an OLE binary field to avoid standard OLE object bloat. At run time the
' original compressed JPEG is displayed by the Intel
' Jpeg Library by copying the field data directly to
' memory without having to save it to a temporary disk file.
'
' 2) Use the Intel JPEG library to allow the user to save Bitmap,
' Gif, Metafile or Bitmap files to the JPEG format.

'­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­­ 
'
'Author:    Stephen Lebans
'
'Email:     Stephen@lebans.com
'
'Web Site:  www.lebans.com
'
'Date:      Feb 06, 2001, 11:22:18 PM
'
'Inputs:    See inline Comments for explanation

'Output:    See inline Comments for explanation
'
'Credits:   Intel Jpeg Library
'           http://developer.intel.com/software/products/perflib/ijl/index.htm
'           http://developer.intel.com/software/products/perflib/ijl/ijllicense.htm
'
'           http://www.vbaccelerator.com/
'           Steve McMahon added 2 functions to the Intel Jpeg VB project that
'           I have modified the implementation to allow them to be used in
'           the VBA Access environment.
'           They are LoadJPGFromPtr & SaveJPGToPtr
'           http://www.vbaccelerator.com/codelib/gfx/vbjpeg.htm
'
'BUGS:      Please report any bugs to my email address.
'
'What's Missing:
'           Lots...this is just a starting point, but you have to start somewhere.

'
' Enjoy
' Stephen Lebans


Option Compare Database
Option Explicit

Private m_cDib As New cDIBSection
' Temp vars
Dim blRet As Boolean
Dim lngRet As Long



Private Sub CmdLoad_Click()
On Error GoTo Err_CmdLoad_Click


 Dim jpg_scale As Long
 Dim strfName As String
  strfName = m_cDib.FileDialog(True)
 If Len(strfName) & vbNullString = 0 Then Exit Sub
 
 ' Display at 100%
 jpg_scale = 1
  ' Read JPEG image
  If LoadJPG(m_cDib, strfName, jpg_scale) Then
   Call m_cDib.DIBtoPictureData(Me.Image0)
  Else
    MsgBox "Unable to Load Jpeg Image", vbCritical
  Exit Sub

  End If
' Enable the SaveBlob command button
   Me.cmdSaveBlob.Enabled = True
' Enable the Frame Size group
   Me.FrameSize.Enabled = True
   
Exit_CmdLoad_Click:
    Exit Sub

Err_CmdLoad_Click:
    MsgBox Err.Description
    Resume Exit_CmdLoad_Click
    
End Sub

Private Sub FrameSize_AfterUpdate()
On Error GoTo Err_Size_Click

' Redisplay current JPEG image at selected ratio
' valid ratios are 1, 1/2, 1/4 and 1/8
' Read JPEG image
If LoadJPG(m_cDib, m_cDib.CurrentJpegFileName, Me.FrameSize.Value) Then
    Call m_cDib.DIBtoPictureData(Me.Image0)
Else
    MsgBox "Unable to Load Jpeg Image", vbCritical
End If

Exit_Size_Click:
    Exit Sub

Err_Size_Click:
    MsgBox Err.Description
    Resume Exit_Size_Click
    
End Sub


Private Sub cmdLoadBlob_Click()
Dim lngPtr As Long
Dim lngSize As Long
Dim varTemp As Variant
Dim bArray() As Byte


' Is the field empty for this record?
   varTemp = Me.Site_Sketch
   If IsNull(varTemp) Then
    Me.Image0.Picture = ""
    Exit Sub
   End If
   ' Resize array to hold BLOB
   ReDim bArray(LenB(varTemp) - 1)
   ' Copy temp var to our byte array
   bArray = varTemp
   ' Get a pointer to the first byte of the array
   lngPtr = VarPtr(bArray(0))
   ' Get total size of the data
   lngSize = UBound(bArray) + 1
   ' Call the function to load the Jpeg from
   ' a buffer(byte array) instead of a file.
   If LoadJPGFromPtr(m_cDib, lngPtr, lngSize) Then
   ' Create a PictureData prop from the loaded Jpeg data.
   Call m_cDib.DIBtoPictureData(Me.Image0)
   Else
      MsgBox "Failed to load from the file: " & m_cDib.CurrentJpegFileName, vbInformation
   End If

End Sub

Private Sub cmdLoadImageCtl_Click()

' Calls the API File Dialog Window
' Returns full path to new File.
' If LoadSave = TRUE then call File Load Dialog

On Error GoTo Err_fFileDialog

' Call the File Common Dialog Window
Dim clsDialog As Object
Dim strTemp As String
Dim strfName As String

Set clsDialog = New clsCommonDialog

' Fill in our structure
clsDialog.Filter = clsDialog.Filter & "ALL (*.*)" & Chr$(0) & "*.*" & Chr$(0)

' Display the Open File Dialog
clsDialog.DialogTitle = "Please Select an Image File to Load"
clsDialog.ShowOpen

' See if user clicked Cancel or even selected
' the very same file already selected
strfName = clsDialog.filename
If Len(strfName & vbNullString) = 0 Then
Set clsDialog = Nothing
Exit Sub
' Raise the exception
 ' Err.Raise vbObjectError + 513, "frmStoreLoadJpeg.fFileDialog", _
 ' "Please type in a Name for a New File"
End If

' Load the Image control with the selected file
Me.Image0.Picture = strfName

Exit_fFileDialog:

Err.Clear
Set clsDialog = Nothing
Exit Sub

Err_fFileDialog:

MsgBox Err.Description, vbOKOnly, Err.Source & ":" & Err.Number
Resume Exit_fFileDialog


End Sub

Private Sub cmdSaveBlob_Click()
On Error GoTo Err_cmdSaveBlob_Click
Dim blRet As Boolean

Dim varTemp As Variant
' Copy original Jpeg file into byte array
If m_cDib.LoadJpegFileIntoArray Then
    ' Copy the byte array to our Blob field for this record
    Me.Site_Sketch = m_cDib.JPegAsByteArray
End If
'Save the record
Me.Dirty = False
Exit_cmdSaveBlob_Click:
    Exit Sub

Err_cmdSaveBlob_Click:
    MsgBox Err.Description
    Resume Exit_cmdSaveBlob_Click
    
End Sub


Private Sub cmdSaveJpeg_Click()
 Dim strfName As String
On Error GoTo LabelExit:


If (m_cDib.PictureDataToDIB(Me.Image0)) Then
DoEvents
'Me.Image0.Picture = ""
Call m_cDib.DIBtoPictureData(Me.Image0)
Else: MsgBox "Failed to save Jpeg"
End If
' Display the Save File dialog box

strfName = m_cDib.FileDialog(False)
If strfName <> "" Then
    Call SaveJPG(m_cDib, strfName)
End If

LabelExit:

End Sub

Private Sub Form_Current()
' Disable the SaveBlob commandButton until
' the user has loaded a Jpeg file
Me.cmdSaveBlob.Enabled = False
' Load the Blob field's contents into the Image control
Call cmdLoadBlob_Click


' Disable the Frame Size group
   Me.FrameSize.Enabled = False
End Sub




Private Sub Form_Load()
  Dim src As Long
  Dim dst As Long
  Dim strTemp As String
  Dim Major As Long, Minor As Long, build As Long
  Dim szVersion As String
  Dim Version As IJLibVersion
  
  DoCmd.Maximize
  ' initial title
  strTemp = "Intel(R) JPEG Library: "

  ' get pointer to IJLibVersion from IJL
  src = ijlGetLibVersion()
  dst = VarPtr(Version)

  ' get data from pointer
  Call CopyMemory(dst, src, Len(Version))

  Major = Version.Major
  Minor = Version.Minor
  build = Version.build

  ' prepare version string
  szVersion = "[" + CStr(Version.Major) + "." + CStr(Version.Minor) + "." + CStr(Version.build) + "]"
    Me.Caption = strTemp & szVersion
  

End Sub





Private Sub Form_Unload(Cancel As Integer)
' Release our DIB class reference
Set m_cDib = Nothing

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
