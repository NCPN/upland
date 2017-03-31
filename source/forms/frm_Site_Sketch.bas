Version =20
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
    Begin
        Begin Label
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
            ForeColor =-2147483630
            FontName ="MS Sans Serif"
        End
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
        End
        Begin BoundObjectFrame
            SpecialEffect =2
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="MS Sans Serif"
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
                End
                Begin BoundObjectFrame
                    Visible = NotDefault
                    OverlapFlags =87
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
                            OverlapFlags =93
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
                        0x000300010000000020000000400000006000000080000000a0000000c0000000 ,
                        0xe00000000020000020200000402000006020000080200000a0200000c0200000 ,
                        0xe02000000040000020400000404000006040000080400000a0400000c0400000 ,
                        0xe04000000060000020600000406000006060000080600000a0600000c0600000 ,
                        0xe06000000080000020800000408000006080000080800000a0800000c0800000 ,
                        0xe080000000a0000020a0000040a0000060a0000080a00000a0a00000c0a00000 ,
                        0xe0a0000000c0000020c0000040c0000060c0000080c00000a0c00000c0c00000 ,
                        0xe0c0000000e0000020e0000040e0000060e0000080e00000a0e00000c0e00000 ,
                        0xe0e000000000400020004000400040006000400080004000a0004000c0004000 ,
                        0xe00040000020400020204000402040006020400080204000a0204000c0204000 ,
                        0xe02040000040400020404000404040006040400080404000a0404000c0404000 ,
                        0xe04040000060400020604000406040006060400080604000a0604000c0604000 ,
                        0xe06040000080400020804000408040006080400080804000a0804000c0804000 ,
                        0xe080400000a0400020a0400040a0400060a0400080a04000a0a04000c0a04000 ,
                        0xe0a0400000c0400020c0400040c0400060c0400080c04000a0c04000c0c04000 ,
                        0xe0c0400000e0400020e0400040e0400060e0400080e04000a0e04000c0e04000 ,
                        0xe0e040000000800020008000400080006000800080008000a0008000c0008000 ,
                        0xe00080000020800020208000402080006020800080208000a0208000c0208000 ,
                        0xe02080000040800020408000404080006040800080408000a0408000c0408000 ,
                        0xe04080000060800020608000406080006060800080608000a0608000c0608000 ,
                        0xe06080000080800020808000408080006080800080808000a0808000c0808000 ,
                        0xe080800000a0800020a0800040a0800060a0800080a08000a0a08000c0a08000 ,
                        0xe0a0800000c0800020c0800040c0800060c0800080c08000a0c08000c0c08000 ,
                        0xe0c0800000e0800020e0800040e0800060e0800080e08000a0e08000c0e08000 ,
                        0xe0e080000000c0002000c0004000c0006000c0008000c000a000c000c000c000 ,
                        0xe000c0000020c0002020c0004020c0006020c0008020c000a020c000c020c000 ,
                        0xe020c0000040c0002040c0004040c0006040c0008040c000a040c000c040c000 ,
                        0xe040c0000060c0002060c0004060c0006060c0008060c000a060c000c060c000 ,
                        0xe060c0000080c0002080c0004080c0006080c0008080c000a080c000c080c000 ,
                        0xe080c00000a0c00020a0c00040a0c00060a0c00080a0c000a0a0c000c0a0c000 ,
                        0xe0a0c00000c0c00020c0c00040c0c00060c0c00080c0c000a0c0c000c0c0c000 ,
                        0xe0c0c00000e0c00020e0c00040e0c00060e0c00080e0c000a0e0c000c0e0c000 ,
                        0xe0e0c00000000000
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
If LoadJPG(m_cDib, m_cDib.CurrentJpegFileName, Me.FrameSize.value) Then
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
clsDialog.Filter = clsDialog.Filter & "ALL (*.*)" & chr$(0) & "*.*" & chr$(0)

' Display the Open File Dialog
clsDialog.DialogTitle = "Please Select an Image File to Load"
clsDialog.ShowOpen

' See if user clicked Cancel or even selected
' the very same file already selected
strfName = clsDialog.fileName
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

MsgBox Err.Description, vbOKOnly, Err.source & ":" & Err.Number
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
