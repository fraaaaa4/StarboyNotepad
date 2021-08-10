Attribute VB_Name = "Module9"
Option Explicit

Private Type Rect
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type CHARRANGE
  cpMin As Long
  cpMax As Long
End Type

Private Type FORMATRANGE
  hdc As Long
  hdcTarget As Long
  rc As Rect
  rcPage As Rect
  chrg As CHARRANGE
End Type

Private Const WM_USER As Long = &H400
Private Const EM_FORMATRANGE As Long = WM_USER + 57
Private Const EM_SETTARGETDEVICE As Long = WM_USER + 72
Private Const PHYSICALOFFSETX As Long = 112
Private Const PHYSICALOFFSETY As Long = 113

Private Declare Function GetDeviceCaps Lib "gdi32" _
   (ByVal hdc As Long, ByVal nIndex As Long) As Long
   
Private Declare Function SendMessage Lib "USER32" _
  Alias "SendMessageA" (ByVal hWnd As Long, ByVal msg As Long, _
  ByVal wp As Long, lp As Any) As Long
  
  Public Function PrintRTFWithMargins(Optional Category, Optional RTFControl As Object, _
   Optional ByVal LeftMargin As Single, Optional ByVal TopMargin As Single, _
   Optional ByVal RightMargin As Single, Optional ByVal BottomMargin As Single) _
   As Boolean
   
'********************************************************8
'PURPOSE: Prints Contents of RTF Control with Margins

'PARAMETERS:
    'RTFControl: RichTextBox Control For Printing
    'LeftMargin: Left Margin in Inches
    'TopMargin: TopMargin in Inches
    'RightMargin: RightMargin in Inches
    'BottomMargin: BottomMargin in Inches

'***************************************************************




'*************************************************************
'I DO THIS BECAUSE IT IS MY UNDERSTANDING THAT
'WHEN CALLING A SERVER DLL, YOU CAN RUN INTO
'PROBLEMS WHEN USING EARLY BINDING WHEN A PARAMETER
'IS A CONTROL OR A CUSTOM OBJECT.  IF YOU JUST PLUG THIS INTO
'A FORM, YOU CAN DECLARE RTFCONTROL AS RICHTEXTBOX
'AND COMMENT OUT THE FOLLOWING LINE


'**************************************************************
   
   Dim lngLeftOffset As Long
   Dim lngTopOffSet As Long
   Dim lngLeftMargin As Long
   Dim lngTopMargin As Long
   Dim lngRightMargin As Long
   Dim lngBottomMargin As Long
   
   Dim typFr As FORMATRANGE
   Dim rectPrintTarget As Rect
   Dim rectPage As Rect
   Dim lngTxtLen As Long
   Dim lngPos As Long
   Dim lngRet As Long
   Dim iTempScaleMode As Integer
   
   iTempScaleMode = Printer.ScaleMode
   
   ' needed to get a Printer.hDC
   Printer.Print ""
   Printer.ScaleMode = vbTwips

   ' Get the offsets to printable area in twips
   lngLeftOffset = Printer.ScaleX(GetDeviceCaps(Printer.hdc, _
      PHYSICALOFFSETX), vbPixels, vbTwips)
   lngTopOffSet = Printer.ScaleY(GetDeviceCaps(Printer.hdc, _
      PHYSICALOFFSETY), vbPixels, vbTwips)

   ' Get Margins in Twips
   lngLeftMargin = InchesToTwips(LeftMargin) - lngLeftOffset
   lngTopMargin = InchesToTwips(TopMargin) - lngTopOffSet
   lngRightMargin = (Printer.Width - _
       InchesToTwips(RightMargin)) - lngLeftOffset
   
    lngBottomMargin = (Printer.Height - _
         InchesToTwips(BottomMargin)) - lngTopOffSet

   ' Set printable area rect
   rectPage.Left = 0
   rectPage.Top = 0
   rectPage.Right = Printer.ScaleWidth
   rectPage.Bottom = Printer.ScaleHeight

   ' Set rect in which to print, based on margins passed in
   rectPrintTarget.Left = lngLeftMargin
   rectPrintTarget.Top = lngTopMargin
   rectPrintTarget.Right = lngRightMargin
   rectPrintTarget.Bottom = lngBottomMargin

   ' Set up the printer for this print job
   typFr.hdc = Printer.hdc 'for rendering
   typFr.hdcTarget = Printer.hdc 'for formatting
   typFr.rc = rectPrintTarget
   typFr.rcPage = rectPage
   typFr.chrg.cpMin = 0
   typFr.chrg.cpMax = -1

   ' Get length of text in the RichTextBox Control
   lngTxtLen = Len(Form1.RichTextBox1.Text)

   ' print page by page
   Do
      ' Print the page by sending EM_FORMATRANGE message
      'Allows you to range of text within a specific device
      'here, the device is the printer, which must be specified
      'as hdc and hdcTarget of the FORMATRANGE structure
      
      lngPos = SendMessage(Form1.RichTextBox1.hWnd, EM_FORMATRANGE, _
        True, typFr)
  
       If lngPos >= lngTxtLen Then Exit Do  'Done
       typFr.chrg.cpMin = lngPos ' Starting position next page
      Printer.NewPage             ' go to next page
      Printer.Print ""   'to get hDC again
      typFr.hdc = Printer.hdc
      typFr.hdcTarget = Printer.hdc
   Loop

   ' Done
   Printer.EndDoc

   ' This frees memory
   lngRet = SendMessage(Form1.RichTextBox1.hWnd, EM_FORMATRANGE, _
     False, Null)
   Printer.ScaleMode = iTempScaleMode
   PrintRTFWithMargins = True
   Exit Function
    
End Function

Private Function InchesToTwips(ByVal Inches As Single) As Single
    InchesToTwips = 1440 * Inches
End Function

Public Sub VB6IHateYou()
Call PrintRTFWithMargins
End Sub

