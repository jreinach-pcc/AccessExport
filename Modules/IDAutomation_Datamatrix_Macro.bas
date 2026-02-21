Attribute VB_Name = "IDAutomation_Datamatrix_Macro"
'NOTICE:
'The procedures located at
'http://www.idautomation.com/fonts/datamatrix/faq.html#Excel_Access_VBA
'must be followed to activate this macro.
'Copyright, IDAutomation.com, Inc. 2001-2007. All rights reserved.
Public Function EncDM(DataToEncode As String, Optional ProcTilde As Integer = 0, Optional EncMode As Integer = 0, Optional PrefFormat As Integer = 0) As String
    'Format the data to the Data Matrix Font by calling the ActiveX DLL:
    'Remove the comment for after install the 2D font to computer for2D-Barcode
    Dim DMFontEncoder As DMATRIXLib.Datamatrix
    Set DMFontEncoder = New Datamatrix
    DMFontEncoder.FontEncode DataToEncode, ProcTilde, EncMode, PrefFormat, EncDM
End Function
