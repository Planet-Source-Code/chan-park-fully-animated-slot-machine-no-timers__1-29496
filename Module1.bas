Attribute VB_Name = "mdlMain"
Option Explicit
Public bErrorCheck As Boolean       'Has an error occurred?
Public bHelpFrom As Boolean         'Stores whether help frmHelp has been
                                        'accessed from frmStart(True) or
                                        'frmMain(False)

'******************************************
' Declare Bitblt API function for animating
' Returns a 0 if error occurs
'******************************************
'This function reduces and eliminates flickering during animation
'because it allows the loading of the image to be done 'behind the scenes'
Public Declare Function BitBlt Lib "gdi32" ( _
    ByVal hDestDC As Long, _
    ByVal xDest As Long, _
    ByVal yDest As Long, _
    ByVal dWdth As Long, _
    ByVal dHght As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, _
    ByVal ySrc As Long, _
    ByVal dwRop As Long) As Long
    
'* Sets raster operation constant*
Public Const SRCCOPY = &HCC0020         'copies source image onto destination
'************************************
'************************************

'***********************************************
' Declare LoadImageA API function to load images
' Returns the handle of the object if successful
' else, returns a 0
'***********************************************
'This function loads an image in a file into memory
Public Declare Function LoadImageA Lib "user32" ( _
    ByVal hInst As Long, _
    ByVal lpsz As String, _
    ByVal iType As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal fOptions As Long) As Long
    
    '* iType options: *
Public Const IMAGE_BITMAP = 0
    '* fOptions flags: *
Public Const LR_LOADFROMFILE = &H10
'************************************
'************************************

'********************************************
' Declare API functions for creating and
' deleting device contexts
'********************************************
Public Declare Function CreateCompatibleDC Lib "gdi32" ( _
    ByVal hdc As Long) As Long
    
Public Declare Function DeleteDC Lib "gdi32" ( _
    ByVal hdc As Long) As Long
'********************************************
'********************************************

'********************************************
' Declare API function for putting an object
' into a device context, by passing that
' object's handle and the destination device
' context
'********************************************
Public Declare Function SelectObject Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal hObject As Long) As Long
'********************************************
'********************************************


'************************************************
' Creates function that loads an image, generates
' a compatible device context, then puts the
' loaded image into the device context
' Returns the image's DC, and changes lBMPHndl's
' value to the loaded image's handle
'************************************************
Public Function GenerateDC(ByVal sBMPFile As String, ByRef lBMPHndl As Long) As Long
Dim lDC As Long
Dim lBitmapHdl As Long

'Create a Device Context, compatible with the screen
lDC = CreateCompatibleDC(0)

'Returns an error value
If lDC = 0 Then
    GenerateDC = 0
    Exit Function
End If

'Load the image
lBitmapHdl = LoadImageA(0, sBMPFile, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)

If lBitmapHdl = 0 Then 'Failure in loading bitmap
    DeleteDC lDC
    GenerateDC = 0
    Exit Function
End If

'Throw the Bitmap into the Device Context
SelectObject lDC, lBitmapHdl

'Return the device context and handle
lBMPHndl = lBitmapHdl
GenerateDC = lDC

End Function

'********************************************
' Subroutine for checking errors
'********************************************
Public Sub ErrorCheck()
If bErrorCheck = True Then
    MsgBox "An error has occurred, please make sure that all program files are in their right places and/or restart your computer", vbCritical
    End
End If
End Sub
'********************************************
'********************************************
