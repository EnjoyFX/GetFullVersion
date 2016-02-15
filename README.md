# GetFullVersion
 VBA code for detection full list of Microsoft Office product parameters
 
 including: Public Type MSFullVersion (Raw As String, ReleaseVers As String, ReleaseType As String, MajorVersion As String, MinorVersion As String, ProductID As String, LanguageID As String, Year As String, x64 As Boolean, ForDebug  As Boolean, IsOffice As Boolean

Usage: add Module to your project and add  to code: result = GetFullVersion(Errors).
if result = true, you can get all (or some) parameters above for your further work.
