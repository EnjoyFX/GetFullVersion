Attribute VB_Name = "VBA_versions"
'*----------------------------------------------------------------------------*
'|  Description of the numbering scheme for product code GUIDs in Office 2013 |
'|  Written by AndyFX, 15 Feb 2016                                            |
'*----------------------------------------------------------------------------*
'|  Based on https://support.microsoft.com/en-us/kb/2786054                   |
'*----------------------------------------------------------------------------*

Option Explicit

Public Type MSFullVersion
    Raw As String
End Type

Public Ver As MSFullVersion


Public Function GetFullVersion(errMessage As String) As Boolean
    Dim ProductKey As String

    GetFullVersion = False

    On Error GoTo errHandler
    ProductKey = Application.ProductCode
    Ver.Raw = ProductKey
    GetFullVersion = True
    Exit Function

errHandler:
    errMessage = err.Description
End Function

Public Sub testMe()

    Dim result As Boolean
    Dim Answer As String, Errors As String

    Errors = ""
    result = GetFullVersion(Errors)

    If result Then Answer = Ver.Raw Else Answer = Errors

    MsgBox Answer

End Sub
