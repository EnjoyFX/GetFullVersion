Attribute VB_Name = "VBA_versions"
'*----------------------------------------------------------------------------*
'|  Description of the numbering scheme for product code GUIDs in Office 2013 |
'|  Written by AndyFX, 15 Feb 2016                                            |
'*----------------------------------------------------------------------------*
'|  Based on https://support.microsoft.com/en-us/kb/2786054                   |
'*----------------------------------------------------------------------------*
' Support products year 2007 - 2013

Option Explicit

Public Type MSFullVersion
    
    Raw As String
    ReleaseVers As String
    ReleaseType As String
    MajorVersion As String
    MinorVersion As String
    ProductID As String
    LanguageID As String
    Year As String
    
    x64 As Boolean
    ForDebug  As Boolean
    IsOffice As Boolean
    
End Type

Private Type KeyPosition
    Start As String
    Size As String
End Type

' These constants coded as [start key pos]:[size of key]
Const pReleaseVers = "2:1"
Const pReleaseType = "3:1"
Const pMajorVersion = "4:2"
Const pMinorVersion = "6:4"
Const pProductID = "11:4"
Const pLanguageID = "16:4"
Const px64 = "21:1"
Const pForDebug = "26:1"
Const pIsOffice = "27:11"

Public Ver As MSFullVersion
Private Key As KeyPosition


Public Function GetFullVersion(errMessage As String) As Boolean
    Dim ProductKey As String


    GetFullVersion = False

    On Error GoTo errHandler
    ProductKey = Application.ProductCode

    With Ver
        .Raw = ProductKey
        .ReleaseVers = GetSystemValue(ProductKey, pReleaseVers)
        .ReleaseType = GetSystemValue(ProductKey, pReleaseType)
        .MajorVersion = GetSystemValue(ProductKey, pMajorVersion)
        .MinorVersion = GetSystemValue(ProductKey, pMinorVersion)
        .ProductID = GetSystemValue(ProductKey, pProductID)
        .LanguageID = GetSystemValue(ProductKey, pLanguageID)
        .x64 = GetSystemValue(ProductKey, px64)
        .ForDebug = GetSystemValue(ProductKey, pForDebug)
        .IsOffice = GetSystemValue(ProductKey, pIsOffice)
    End With

    GetFullVersion = True
    Exit Function

errHandler:
    errMessage = err.Description
End Function
Function GetSystemValue(ProductKey, ItemPtr As String) As String
    Dim StartOfKey As Integer
    Dim Size_OfKey As Integer
    Dim KeyValue As String

    GetSystemValue = ""
    StartOfKey = gStart(ItemPtr)
    Size_OfKey = gSize(ItemPtr)
    
    KeyValue = gKeyValue(ProductKey, StartOfKey, Size_OfKey)
    
    Select Case ItemPtr
        Case pReleaseVers:      GetSystemValue = ScanReleaseVersion(KeyValue)
        Case pReleaseType:      GetSystemValue = ScanReleaseType(KeyValue)
        Case pMajorVersion:     GetSystemValue = KeyValue
        Case pMinorVersion:     GetSystemValue = KeyValue
        Case pProductID:        GetSystemValue = ScanProductID(KeyValue)
        Case pLanguageID:       GetSystemValue = ScanLanguageID(KeyValue)
        Case px64:              GetSystemValue = IIf(KeyValue = "1", True, False)
        Case pForDebug:         GetSystemValue = IIf(KeyValue = "D", True, False)
        Case pIsOffice:         GetSystemValue = IIf(KeyValue = "000000FF1CE", True, False)
    End Select
    
End Function
Private Function gStart(ByVal s1 As String, Optional ByVal Delim As String = ":") As Integer
    Dim A As Integer
    A = InStr(s1, Delim): If A = 0 Then gStart = s1 Else gStart = Val(Left(s1, A - 1))
End Function
Private Function gSize(ByVal s1 As String, Optional ByVal Delim As String = ":") As Integer
    Dim A As Integer
    A = InStr(s1, Delim): If A = 0 Then gSize = s1 Else gSize = Val(Right(s1, Len(s1) - A))
End Function
Private Function gKeyValue(ProductKey, Start, Size) As String
    gKeyValue = UCase(Mid(ProductKey, Start, Size))
End Function
Function ScanReleaseVersion(KeyValue)
    Dim msg As String
    
    Select Case KeyValue
        Case "0": msg = "Release before Beta 1"
        Case "1": msg = "Beta 1"
        Case "2": msg = "Beta 2"
        Case "3": msg = "Release Candidate 0 (RC0)"
        Case "4": msg = "Release Candidate 1 (RC1)/OEM Preview release"
        Case "9": msg = "RTM. This is the first version that is shipped (the initial release)"
        Case "A": msg = "Service Pack 1 (SP1). This value is not used if the product code is not changed after the RTM version"
        Case "B": msg = "Service Pack 2 (SP2). This value is not used if the product code is not changed after the RTM version"
        Case "C": msg = "Service Pack 3 (SP3). This value is not used if the product code is not changed after the RTM version"
        Case Else:  msg = ""
    End Select
    ScanReleaseVersion = msg
    
End Function

Function ScanReleaseType(KeyValue)
    Dim msg As String
    
    Select Case KeyValue
        Case "0": msg = "Volume license"
        Case "1": msg = "Retail/OEM"
        Case "2": msg = "Trial"
        Case "5": msg = "Download"
        Case Else:  msg = ""
    End Select
    ScanReleaseType = msg
    
End Function

Function ScanProductID(KeyValue)
    Dim msg As String
     
     Select Case Ver.MajorVersion
        Case "12": Ver.Year = "2007"
        Case "14": Ver.Year = "2010"
        Case "15": Ver.Year = "2013"
        Case Else: Ver.Year = ""
    End Select
    
    Select Case KeyValue
        '2013 keys presents:
        Case "0011": msg = "Office Professional Plus"
        Case "0012": msg = "Office Standard"
        Case "0013": msg = "Office Home and Business"
        Case "0014": msg = "Office Professional"
        Case "0015": msg = "Access"
        Case "0016": msg = "Excel"
        Case "0017": msg = "SharePoint Designer"
        Case "0018": msg = "PowerPoint"
        Case "0019": msg = "Publisher"
        Case "001A": msg = "Outlook"
        Case "001B": msg = "Word"
        Case "001C": msg = "Access Runtime"
        Case "001F": msg = "Office Proofing Tools Kit Compilation"
        Case "002F": msg = "Office Home and Student"
        Case "003A": msg = "Project Standard"
        Case "003B": msg = "Project Professional"
        Case "0044": msg = "InfoPath"
        Case "0051": msg = "Visio Professional"
        Case "0053": msg = "Visio Standard"
        Case "00A1": msg = "OneNote"
        Case "00BA":
            Select Case Ver.Year
                Case "2010", "2013": msg = "Office SharePoint Workspace"
                Case "2007":         msg = "Office Groove"
            End Select
        Case "110D": msg = "Office SharePoint Server"
        Case "110F": msg = "Project Server"
        Case "012B": msg = "Lync"
        
        '2010 keys (not presents in 2013):
        Case "011D": msg = "Office Professional Plus Subscription"
        Case "007A": msg = "Outlook Connector"
        Case "008B": msg = "Office Small Business Basics"
        Case "0052": msg = "Visio Viewer"
        Case "0057": msg = "Visio"
        Case "00AF": msg = "PowerPoint Viewer"
        
        '2007 keys (not presents in 2010):
        Case "00A3": msg = "Office OneNote Home Student"
        Case "00A7": msg = "Calendar Printing Assistant for Microsoft Office Outlook"
        Case "00A9": msg = "Office InterConnect"
        Case "00B0": msg = "Save as PDF add-in"
        Case "00B1": msg = "Save as XPS add-in"
        Case "00B2": msg = "Save as PDF or XPS add-in"
        Case "00CA": msg = "Office Small Business"
        Case "10D7": msg = "Office InfoPath Forms Services"
        Case "1122": msg = "Windows SharePoint Services Developer Resources 1.2"
        Case "0010": msg = "SKU - Microsoft Software Update for Web Folders (English) 12"
        Case "0020": msg = "Office Compatibility Pack for Word, Excel, and PowerPoint File Formats"
        Case "0026": msg = "Expression Web"
        Case "002E": msg = "Office Ultimate"
        Case "0030": msg = "Office Enterprise"
        Case "0031": msg = "Office Professional Hybrid"
        Case "0033": msg = "Office Personal"
        Case "0035": msg = "Office Professional Hybrid"

        Case Else:  msg = ""
    End Select
    
    ScanProductID = "Microsoft " + msg + " " + Ver.Year
    
End Function

Function ScanLanguageID(KeyValue)
    Dim msg As String, DecID As String
    DecID = CStr("&H" + (KeyValue))
    
    Select Case DecID
        Case "1025": msg = "Arabic"
        Case "1026": msg = "Bulgarian"
        Case "2052": msg = "Chinese (Simplified)"
        Case "1028": msg = "Chinese"
        Case "1050": msg = "Croatian"
        Case "1029": msg = "Czech"
        Case "1030": msg = "Danish"
        Case "1043": msg = "Dutch"
        Case "1033": msg = "English"
        Case "1061": msg = "Estonian"
        Case "1035": msg = "Finnish"
        Case "1036": msg = "French"
        Case "1031": msg = "German"
        Case "1032": msg = "Greek"
        Case "1037": msg = "Hebrew"
        Case "1081": msg = "Hindi"
        Case "1038": msg = "Hungarian"
        Case "1057": msg = "Indonesian"
        Case "1040": msg = "Italian"
        Case "1041": msg = "Japanese"
        Case "1087": msg = "Kazakh"
        Case "1042": msg = "Korean"
        Case "1062": msg = "Latvian"
        Case "1063": msg = "Lithuanian"
        Case "1086": msg = "Malay"
        Case "1044": msg = "Norwegian (Bokm?l)"
        Case "1045": msg = "Polish"
        Case "1046": msg = "Portuguese"
        Case "2070": msg = "Portuguese"
        Case "1048": msg = "Romanian"
        Case "1049": msg = "Russian"
        Case "2074": msg = "Serbian (Latin)"
        Case "1051": msg = "Slovak"
        Case "1060": msg = "Slovenian"
        Case "3082": msg = "Spanish"
        Case "1053": msg = "Swedish"
        Case "1054": msg = "Thai"
        Case "1055": msg = "Turkish"
        Case "1058": msg = "Ukrainian"
        Case "1066": msg = "Vietnamese"
        Case Else:  msg = ""
    End Select
    ScanLanguageID = msg
    
End Function

Public Sub testMe()

    Dim result As Boolean
    Dim Answer As String, Errors As String

    Errors = ""
    result = GetFullVersion(Errors)

    If result Then
        Answer = _
        "[Release Version] " + Ver.ReleaseVers + vbCr + _
        "[Release Type] " + Ver.ReleaseType + vbCr + _
        "[Year] " + Ver.Year + vbCr + _
        "[Major Version] " + Ver.MajorVersion + vbCr + _
        "[Minor Version] " + Ver.MinorVersion + vbCr + _
        "[Product ID] " + Ver.ProductID + vbCr + _
        "[Language ID] " + Ver.LanguageID + vbCr + _
        "[x64] " + CStr(Ver.x64) + vbCr + _
        "[for Debug] " + CStr(Ver.ForDebug) + vbCr + _
        "[Office ID present] " + CStr(Ver.IsOffice)
    Else
        Answer = Errors
    End If

    MsgBox Answer

End Sub
