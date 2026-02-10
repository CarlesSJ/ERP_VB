Attribute VB_Name = "connexiodsn"

    Option Explicit

    Private Const REG_SZ = 1    'Constant for a string variable type.
    Private Const HKEY_LOCAL_MACHINE = &H80000002

    Private Declare Function RegCreateKey Lib "advapi32.dll" Alias _
       "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, _
       phkResult As Long) As Long

    Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias _
       "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
       ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal _
       cbData As Long) As Long

    Private Declare Function RegCloseKey Lib "advapi32.dll" _
       (ByVal hKey As Long) As Long
                    

Sub creardsn()

   Dim DataSourceName As String
   Dim DatabaseName As String
   Dim Description As String
   Dim DriverPath As String
   Dim DriverName As String
   Dim LastUser As String
   Dim Regional As String
   Dim Server As String

   Dim lResult As Long
   Dim hKeyHandle As Long

   'Specify the DSN parameters.

   DataSourceName = "tintes"
   DatabaseName = "InkmakerDB"
   Description = "Base de dades Tintes Inkmaker"
   DriverPath = "Sql Server"
   LastUser = "sa"
   Server = "14-00372\SQLEXPRESS"
   DriverName = "SQL Server"

   'Create the new DSN key.

   lResult = RegCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\" & _
        DataSourceName, hKeyHandle)

   'Set the values of the new DSN key.

   lResult = RegSetValueEx(hKeyHandle, "Database", 0&, REG_SZ, _
      ByVal DatabaseName, Len(DatabaseName))
   lResult = RegSetValueEx(hKeyHandle, "Description", 0&, REG_SZ, _
      ByVal Description, Len(Description))
   lResult = RegSetValueEx(hKeyHandle, "Driver", 0&, REG_SZ, _
      ByVal DriverPath, Len(DriverPath))
   lResult = RegSetValueEx(hKeyHandle, "LastUser", 0&, REG_SZ, _
      ByVal LastUser, Len(LastUser))
   lResult = RegSetValueEx(hKeyHandle, "Server", 0&, REG_SZ, _
      ByVal Server, Len(Server))

   'Close the new DSN key.

   lResult = RegCloseKey(hKeyHandle)

   'Open ODBC Data Sources key to list the new DSN in the ODBC Manager.
   'Specify the new value.
   'Close the key.

   lResult = RegCreateKey(HKEY_LOCAL_MACHINE, _
      "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", hKeyHandle)
   lResult = RegSetValueEx(hKeyHandle, DataSourceName, 0&, REG_SZ, _
      ByVal DriverName, Len(DriverName))
   lResult = RegCloseKey(hKeyHandle)
                
End Sub
