Attribute VB_Name = "modMain"
Option Explicit

'#########################################################################################
'#########################################################################################
'#########################################################################################
'########                                                                        #########
'########                                CONSTANTS                               #########
'########                                                                        #########
'#########################################################################################
'#########################################################################################
'#########################################################################################

Public Const HKEY_CURRENT_USER_KEY  As String = "Software\Microsoft\Visual Basic\6.0\RecentFiles"
Public Const HKEY_CURRENT_USER      As Long = &H80000001

Public Const KEY_QUERY_VALUE        As Long = 1
Public Const KEY_ALL_ACCESS         As Long = 63
Public Const REG_SZ                 As Long = 1

Public Const MOVE_ITEM_UP           As Integer = 0
Public Const MOVE_ITEM_DOWN         As Integer = 1

'#########################################################################################
'#########################################################################################
'#########################################################################################
'########                                                                        #########
'########                                   APIs                                 #########
'########                                                                        #########
'#########################################################################################
'#########################################################################################
'#########################################################################################

Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
                        "RegOpenKeyExA" (ByVal hKey As Long, _
                                         ByVal lpSubKey As String, _
                                         ByVal ulOptions As Long, _
                                         ByVal samDesired As Long, _
                                         phkResult As Long) _
                                         As Long
                                
Public Declare Function RegQueryValueEx Lib "advapi32" Alias _
                        "RegQueryValueExA" (ByVal hKey As Long, _
                                            ByVal lpValueName As String, _
                                            ByVal lpReserved As Long, _
                                            ByRef lpType As Long, _
                                            ByVal szData As String, _
                                            ByRef lpcbData As Long) _
                                            As Long
                                            
Public Declare Function RegSetValueEx Lib "advapi32" Alias _
                        "RegSetValueExA" (ByVal hKey As Long, _
                                          ByVal lpValueName As String, _
                                          ByVal Reserved As Long, _
                                          ByVal dwType As Long, _
                                          ByVal szData As String, _
                                          ByVal cbData As Long) _
                                          As Long
                                          
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias _
                        "RegDeleteValueA" (ByVal hKey As Long, _
                                           ByVal lpValueName As String) _
                                           As Long

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

