Option Compare Database
Option Explicit

      
        
Private Const vbPicTypeBitmap = 1

        Private Type IID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Private Type PictDesc
   Size As Long
   Type As Long
   hBmp As Long
   hPal As Long
   Reserved As Long
End Type
        
    'ннннннннннннннннннннннннннннннннннннннннннннннннна
'аPrivate Declare Function OleCreatePictureIndirect Lib _
'а   "olepro32.dll" _
'а   (PicDesc As PictDesc, RefIID As IID, _
'а    ByVal fPictureOwnsHandle As Long, _
'а    IPic As IPicture) As Long
'ннннннннннннннннннннннннннннннннннннннннннннннннна

    
    
'''Windows API Function Declarations

'Does the clipboard contain a bitmap/metafile?
Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Integer) As Long

'Open the clipboard to read
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long

'Get a pointer to the bitmap/metafile
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Integer) As Long

'Close the clipboard
Private Declare Function CloseClipboard Lib "user32" () As Long

'Create our own copy of the metafile, so it doesn't get wiped out by subsequent clipboard updates.
Declare Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As Long, ByVal lpszFile As String) As Long

'Create our own copy of the bitmap, so it doesn't get wiped out by subsequent clipboard updates.
Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

'The API format types we're interested in
Const CF_BITMAP = 2
Const CF_PALETTE = 9
Const CF_ENHMETAFILE = 14
Const IMAGE_BITMAP = 0
Const LR_COPYRETURNORG = &H4
' Addded by SL Apr/2000
Const xlPicture = CF_BITMAP
Const xlBitmap = CF_BITMAP

        
        
        
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
        'Name:      BitmapToPicture &
        '           GetClipBoard
        '
        'Purpose:   Provides a method to save the contents of a
        '           Bound or Unbound OLE Control to a Disk file.
        '           This version only handles BITMAP files.
        '           '
        'Author:    Stephen Lebans
        'Email:     Stephen@lebans.com
        'Web Site:  www.lebans.com
        'Date:      Apr 10, 2000, 05:31:18 AM
        '
        'Called by: Any
        '
        'Inputs:    Needs a Handle to a Bitmap.
        '           This must be a 24 bit bitmap for this release.
        '
        'Credits:
        'As noted directly in Source :-)
        '
        'BUGS:
        'To keep it simple this version only works with Bitmap files of 16 or 24 bits.
        'I'll go back and add the
        'code to allow any depth bitmaps and add support for
        'metafiles as well.
        'No serious bugs notices at this point in time.
        'Please report any bugs to my email address.
        '
        'What's Missing:
        '
        '
        'HOW TO USE:
        '
        '*******************************************
        
       
    



Function GetClipBoard(ImageType As Long) As Long
' Get a handle to a Bitmap object
' from the ClipBoard

' Handles for graphic Objects
Dim hClipBoard As Long
Dim hBitmap As Long
Dim hBitmap2 As Long

'Check if the clipboard contains the required format
'hPicAvail = IsClipboardFormatAvailable(lPicType)

 ' Open the ClipBoard
 hClipBoard = OpenClipboard(0&)

 If hClipBoard <> 0 Then
    ' Get a handle to the Bitmap
    hBitmap = GetClipboardData(CF_BITMAP)

    If hBitmap <> 0 Then 'GoTo exit_error
    ' Create our own copy of the image on the clipboard, in the appropriate format.
    'If lPicType = CF_BITMAP Then
        hBitmap2 = CopyImage(hBitmap, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG)
        ImageType = CF_BITMAP
        Else
        hBitmap = GetClipboardData(CF_ENHMETAFILE)
        hBitmap2 = CopyEnhMetaFile(hBitmap, vbNullString)
        ImageType = CF_ENHMETAFILE
        End If
    
        'Release the clipboard to other programs
        hClipBoard = CloseClipboard

 GetClipBoard = hBitmap2
 Exit Function
 
 End If
    
    
exit_error:
' Return False
GetClipBoard = -1
End Function