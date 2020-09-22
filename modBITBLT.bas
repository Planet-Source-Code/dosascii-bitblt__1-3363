Attribute VB_Name = "modBITBLT"
Option Explicit
Declare Function BitBlt Lib "gdi32" ( _
    ByVal hDestDC As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, _
    ByVal ySrc As Long, _
    ByVal dwRop As Long _
    ) As Long
'
' ================================
' PARAMETERS AND THEIR DESCRIPTION
' ================================
' ByVal hDestDC As Long ..... hDC of object to receive the .bmp
' ByVal x As Long ........... x coordinate (upper-left) destination rectangle
' ByVal y As Long ........... y coordinate (upper-left) destination rectangle
' ByVal nWidth As Long ...... width of the destination rect. and source .bmp
' ByVal nHeight As Long ..... height of the destination rect. and source .bmp
' ByVal hSrcDC As Long ...... hDC of source object that contains the .bmp
' ByVal xSrc As Long ........ x coordinate (upper-left) source .bmp
' ByVal ySrc As Long ........ y coordinate (upper-left) source .bmp
' ByVal dwRop As Long ....... specifies the raster operation to be performed as below
'
' ===========================================================
' RASTER OPERATION CONSTANTS
' ===========================================================
'
' Constants  ------------------     Description
' ===========================================================
'
Public Const SRCCOPY = &HCC0020     'Copies the source bitmap to destination bitmap.
Public Const SRCAND = &H8800C6      'Combines pixels of the destination with source bitmap
                                    'using the Boolean AND operator.
Public Const SRCINVERT = &H660046   'Combines pixels of the destination with source bitmap
                                    'using the Boolean XOR operator.
Public Const SRCPAINT = &HEE0086    'Combines pixels of the destination with source bitmap
                                    'using the Boolean OR operator.
Public Const SRCERASE = &H4400328   'Inverts the destination bitmap and then combines the
                                    'results with the source bitmap using the Boolean AND
                                    'operator.
Public Const WHITENESS = &HFF0062   'Turns all output white.
Public Const BLACKNESS = &H42       'Turn output black.


