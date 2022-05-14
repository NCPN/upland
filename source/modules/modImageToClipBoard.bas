Option Compare Database
Option Explicit

   
     
'*********  Code Start  ************
Private Type METAFILEPICT
 mm As Long
 xExt As Long
 yExt As Long
 hMF As Long
End Type

'ญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญญ 
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hWnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long

Private Declare Function SetEnhMetaFileBits Lib "gdi32" _
(ByVal cbBuffer As Long, lpData As Byte) As Long

Private Declare Function SetWinMetaFileBits Lib "gdi32" _
(ByVal cbBuffer As Long, lpbBuffer As Byte, _
ByVal hDCRef As Long, lpmfp As Any) As Long

Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags&, ByVal _
dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) _
As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) _
As Long

Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) _
As Long

Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) _
As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As _
Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat _
As Long, ByVal hMem As Long) As Long

Private Declare Sub apiCopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
 (Destination As Any, Source As Any, ByVal Length As Long)
 
' CONSTANTS
Private Const WM_HSCROLL = &H114
Private Const WM_VSCROLL = &H115
' Scroll Bar Commands
Private Const SB_PAGEUP = 2
Private Const SB_PAGELEFT = 2

'Global Memory Flags
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const GMEM_SHARE = &H2000

' ClipBoard Formats
Private Const CF_BITMAP = 2
Private Const CF_DIB = 8
Private Const CF_ENHMETAFILE = 14
Private Const CF_METAFILEPICT = 3

     
     
     
     '*******************************************
     'DEVELOPED AND TESTED UNDER MICROSOFT ACCESS 97 VBA ONLY
     '
     'Copyright: Lebans Holdings 1999 Ltd.
     '           May not be resold in whole or part. Please feel
     '           free to use any/all of this code within your
     '           own application without cost or obligation.
     '           Please include the one line Copyright notice
     '           if you use this function in your own code.
     '
     'Name:      FUNCTION() FPictureDataToClipBoard
     '
     'Purpose:   Provides a method to copy the contents of an
     '           Image Control to the ClipBoard. You cannot set
     '           the Focus to an Image Control in Form View in
     '           order to use RunCommand acCmdCopy.
     '
     'Author:    Stephen Lebans
     'Email:     Stephen@lebans.com
     'Web Site:  www.lebans.com
     'Date:      June 16, 2000, 09:55:11 PM
     '
     'Called by: Any
     '
     'Inputs:    Needs a reference to an Access Image Control.
     '
     'Returns:   True on success, False on failure
     '
     'Credits:
     'Anyone that wants some! :-)
     '
     'BUGS:
     'No serious bugs notices at this point in time.
     'Please report any bugs to my email address.
     '
     'What's Missing:
     'There's always something!
     '
     'HOW TO USE:
     'Simply call the function with a reference to the Image
     'control that contains the Picture you want copied
     'to the ClipBoard.
     '
     'HOW IT WORKS:
     ' The first 8 Bytes of a PictureData prop signify
     ' that the data is structured as one of the
     ' following ClipBoard Formats.
     ' CF_DIB
     ' CF_ENHMETAFILE
     ' CF_METAFILEPICT
     ' If the first 40 bytes of a PictureData prop are
     ' not a BITMAPINFOHEADER structure then we will find
     ' a ClipBoard Format structure of 8 Bytes in length
     ' signifying whether a Metafile or Enhanced Metafile is present.
     '
     ' So the first 4 bytes tell us the format of the data.
     ' The next 4 bytes point to handle for a Memory Metafile.
     ' This is not needed for our purposes.
  
     '*******************************************
    

Function FPictureDataToClipBoard(ctl As Access.Image) As Boolean
' Memory Vars
Dim hGlobalMemory As Long
Dim lpGlobalMemory As Long
Dim hClipMemory As Long

' Cf_metafilepict structure
Dim cfm As METAFILEPICT
 
' Handle to a Memory Metafile
Dim hMetafile As Long

' Which ClipBoard format is contained in the PictureData prop
Dim CBFormat As Long

' Byte array to hold the PictureData prop
Dim bArray() As Byte

' Temp var
Dim lngRet As Long

On Error GoTo Err_PtoC

' Resize to hold entire PictureData prop
ReDim bArray(LenB(ctl.PictureData) - 1)

' Copy to our array
bArray = ctl.PictureData

' Determine which ClipBoard format we are using
Select Case bArray(0)


Case 40
' This is a straight DIB.
CBFormat = CF_DIB
' MSDN states to Allocate moveable|Shared Global memory
' for ClipBoard operations.
hGlobalMemory = GlobalAlloc(GMEM_MOVEABLE Or GMEM_SHARE Or GMEM_ZEROINIT, UBound(bArray) + 1)
If hGlobalMemory = 0 Then _
Err.Raise vbObjectError + 515, "ImageToClipBoard.modImageToClipBoard", _
   "GlobalAlloc Failed..not enough memory"

' Lock this block to get a pointer we can use to this memory.
lpGlobalMemory = GlobalLock(hGlobalMemory)
If lpGlobalMemory = 0 Then _
Err.Raise vbObjectError + 516, "ImageToClipBoard.modImageToClipBoard", _
   "GlobalLock Failed"

' Copy DIB as is in its entirety
apiCopyMemory ByVal lpGlobalMemory, bArray(0), UBound(bArray) + 1

' Unlock the memory in preparation to copy to the clipboard
If GlobalUnlock(hGlobalMemory) <> 0 Then _
Err.Raise vbObjectError + 517, "ImageToClipBoard.modImageToClipBoard", _
   "GlobalUnLock Failed"


Case CF_ENHMETAFILE
' New Enhanced Metafile(EMF)
CBFormat = CF_ENHMETAFILE
' Create a Memory based Metafile we can pass to the ClipBoard
hMetafile = SetEnhMetaFileBits(UBound(bArray) + 1 - 8, bArray(8))


Case CF_METAFILEPICT
' Old Metafile format(WMF)
CBFormat = CF_METAFILEPICT
' Create a Memory based Metafile we can pass to the ClipBoard
' We need to convert from the older WMF to the new EMF format
' Copy the Metafile Header over to our Local Structure
apiCopyMemory cfm, bArray(8), Len(cfm)
' By converting the older WMF to EMF this
' allows us to have a single solution for Metafiles.
' 24 is the number of bytes in the sum of the
' METAFILEPICT structure and the 8 byte ClipBoard Format struct.
hMetafile = SetWinMetaFileBits(UBound(bArray) + 24 + 1 - 8, bArray(24), 0&, cfm)

 
Case Else
'Should not happen
Err.Raise vbObjectError + 514, "ImageToClipBoard.modImageToClipBoard", _
   "Unrecognized PictureData ClipBoard format"

End Select

 ' Can we open the ClipBoard.
If OpenClipboard(0&) = 0 Then _
Err.Raise vbObjectError + 518, "ImageToClipBoard.modImageToClipBoard", _
"OpenClipBoard Failed"

' Always empty the ClipBoard First. Not the friendliest thing
' to do if you have several programs interacting!
Call EmptyClipboard

' Now set the Image to the ClipBoard
If CBFormat = CF_ENHMETAFILE Or CBFormat = CF_METAFILEPICT Then

    ' Remember we can use this logic for both types of Metafiles
    ' because we converted the older WMF to the newer EMF.
    hClipMemory = SetClipboardData(CF_ENHMETAFILE, hMetafile)

Else
' We are dealing with a standard DIB.
hClipMemory = SetClipboardData(CBFormat, hGlobalMemory)

End If

If hClipMemory = 0 Then _
    Err.Raise vbObjectError + 519, "ImageToClipBoard.modImageToClipBoard", _
    "SetClipBoardData Failed"

' Close the ClipBoard
lngRet = CloseClipboard
If lngRet = 0 Then _
    Err.Raise vbObjectError + 520, "ImageToClipBoard.modImageToClipBoard", _
    "CloseClipBoard Failed"

  ' Signal Success!
FPictureDataToClipBoard = True
 

Exit_PtoC:
Exit Function


Err_PtoC:
FPictureDataToClipBoard = False
MsgBox Err.Description, vbOKOnly, Err.Source & ":" & Err.Number
Resume Exit_PtoC

End Function


Public Function fLoadPicture(ctl As Access.Image, Optional strfName As String = "") As Boolean

On Error GoTo Err_fLoadPicture

' Temp Vars
Dim lngRet As Long
Dim blRet As Boolean

' Were we passed the Optional FileName and Path
If Len(strfName & vbNullString) = 0 Then
 ' Call the File Common Dialog Window
 Dim clsDialog As Object
 Dim strTemp As String

 Set clsDialog = New clsCommonDialog

 ' Fill in our structure
 clsDialog.Filter = "All Files (*.*)" & Chr$(0) & "*.*" & Chr$(0)
 clsDialog.Filter = clsDialog.Filter & "JPEG (*.JPG)" & Chr$(0) & "*.JPG" & Chr$(0)
 clsDialog.Filter = clsDialog.Filter & "Bmp (*.BMP)" & Chr$(0) & "*.BMP" & Chr$(0)
 clsDialog.Filter = clsDialog.Filter & "Gif (*.GIF)" & Chr$(0) & "*.GIF" & Chr$(0)
 clsDialog.Filter = clsDialog.Filter & "EMF (*.EMF)" & Chr$(0) & "*.EMF" & Chr$(0)
 clsDialog.Filter = clsDialog.Filter & "WMF (*.WMF)" & Chr$(0) & "*.WMF" & Chr$(0)
 
 clsDialog.hdc = 0
 clsDialog.MaxFileSize = 256
 clsDialog.max = 256
 clsDialog.FileTitle = vbNullString
 clsDialog.DialogTitle = "Please Select an Image File to Load"
 clsDialog.InitDir = vbNullString
 clsDialog.DefaultExt = vbNullString
 
 ' Display the File Dialog
 clsDialog.ShowOpen
 
 ' See if user clicked Cancel or even selected
 ' the very same file already selected
 strfName = clsDialog.filename
 If Len(strfName & vbNullString) = 0 Then
 ' Raise the exception
   Err.Raise vbObjectError + 513, "CreateBitmapFromImageCtl.modStdPic", _
   "Please Select a Valid Image File"
 End If

' If we jumped to here then user supplied a FileName
End If

' It may take a few seconds to render larger JPEGs.
' Set the MousePointer to "HOURGLASS"
Application.Screen.MousePointer = 11

' Load the Picture as a StandardPicture object
ctl.Picture = strfName
If ctl.Picture <> strfName Then
 Err.Raise vbObjectError + 514, "CreateBitmapFromImageCtl.modStdPic", _
 "Please Select a Valid Image File"
End If


' Set the Dimensions of the Image Control
' to the actual size of the graphic we are displaying.
' There is a Bug/Feature in how Access handles this
' property. This prop is derived directly from the
' BITMAPINFOHEADER->biXPelsPerMeter & biYPelsPerMeter
' If this value is ZERO in the Bitmap File then an
' Application error occurs and Access fills in the
' Image Controls ImageWidth & Height props with the
' Text from the error.
' The bug is that Access will use whatever values above
' ZERO that are in these members. A lot of Bitmap graphics
' files have garbage or just plain wrong values. This will
' obviously result in incorrect values for these props at
' runtime.

Dim intImageWidth As Long
Dim intImageHeight As Long

' Could be  invalid props here - quite common
On Error Resume Next
intImageWidth = ctl.ImageWidth
intImageHeight = ctl.ImageHeight

If intImageWidth = 0 Then intImageWidth = ctl.Parent.width / 2
If intImageHeight = 0 Then intImageHeight = ctl.Parent.Detail.Height / 2

' Return to normal error handling
On Error GoTo Err_fLoadPicture

' Error check to ensure we do not exceed
' SubForm boundaries
If intImageWidth < ctl.Parent.width Then
 ctl.width = intImageWidth
Else
 ctl.width = ctl.Parent.width - 200
End If

If intImageHeight < ctl.Parent.Detail.Height Then
 ctl.Height = intImageHeight
Else
 ctl.Height = ctl.Parent.Detail.Height - 200
End If

' Scroll the Form back to X:0,Y:0
ScrollToHome ctl

' Cleanup
fLoadPicture = True

Exit_LoadPic:

' Set the MousePointer back to Default
Application.Echo True
Application.Screen.MousePointer = 0
Err.Clear
Set clsDialog = Nothing
Exit Function

Err_fLoadPicture:
fLoadPicture = False
MsgBox Err.Description, vbOKOnly, Err.Source & ":" & Err.Number
Resume Exit_LoadPic

End Function


Public Sub ScrollToHome(ctl As Control)
' Scroll the Form back to X:0,Y:0
' The Form is heavily Subclassed by Access.
' It does not seem to respond to SB_TOP or SB_LEFT
' so we have to resort to the following kludge.

' Temp var
Dim lngRet As Long

' Temp counter
Dim lngTemp As Long

' Be careful because of Echo Off
On Error Resume Next

' Stop Screen Redraws
Application.Echo False

For lngTemp = 1 To 9
lngRet = SendMessage(ctl.Parent.hWnd, WM_VSCROLL, SB_PAGEUP, 0&)
lngRet = SendMessage(ctl.Parent.hWnd, WM_HSCROLL, SB_PAGELEFT, 0&)
Next lngTemp

' Start Screen Redraws
Application.Echo True

End Sub