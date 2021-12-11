# GetFullVersion
## VBA code for detection full list of Microsoft Office product parameters
 
**Included:**

```Public Type MSFullVersion (Raw As String, ReleaseVers As String, ReleaseType As String, MajorVersion As String, MinorVersion As String, ProductID As String, LanguageID As String, Year As String, x64 As Boolean, ForDebug  As Boolean, IsOffice As Boolean)```

**Usage:**

Add Module to your project and add  to code:

```result = GetFullVersion(Errors)```

If result = True, you can get all (or some) parameters above for your further work.

Example in file *GetFullVersion.xlsm*

![image](https://user-images.githubusercontent.com/12576352/145683235-79aeec09-2392-4578-b4f3-a42c175b4a91.png)

![image](https://user-images.githubusercontent.com/12576352/145683252-7ebe73be-cde7-49af-8dfe-7956ef574f6b.png)
