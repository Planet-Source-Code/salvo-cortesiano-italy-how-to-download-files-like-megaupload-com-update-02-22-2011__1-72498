Attribute VB_Name = "modCookies"
' Nome del Progetto: Megaupload Upload-Download v1.0.1
' ****************************************************************************************************
' Copyright © 2008 - 2011 Salvo Cortesiano - Società: http://www.netshadows.it/
' Tutti i diritti riservati, Indirizzo Internet: http://www.netshadows.it/
' ****************************************************************************************************
' Attenzione: Questo programma per computer è protetto dalle vigenti leggi sul copyright
' e sul diritto d'autore. Le riproduzioni non autorizzate di questo codice, la sua distribuzione
' la distribuzione anche parziale è considerata una violazione delle leggi, e sarà pertanto
' perseguita con l'estensione massima prevista dalla legge in vigore.
' ****************************************************************************************************

Option Explicit

'// --------------------------Types, consts and structures
Private Const ERROR_CACHE_FIND_FAIL As Long = 0
Private Const ERROR_CACHE_FIND_SUCCESS As Long = 1
Private Const ERROR_FILE_NOT_FOUND As Long = 2
Private Const ERROR_ACCESS_DENIED As Long = 5
Private Const ERROR_INSUFFICIENT_BUFFER As Long = 122
Private Const MAX_CACHE_ENTRY_INFO_SIZE As Long = 4096

Private Const LMEM_FIXED As Long = &H0
Private Const LMEM_ZEROINIT As Long = &H40

Public Enum eCacheType
    eNormal = &H1&
    eEdited = &H8&
    eTrackOffline = &H10&
    eTrackOnline = &H20&
    eSticky = &H40&
    eSparse = &H10000
    eCookie = &H100000
    eURLHistory = &H200000
    eURLFindDefaultFilter = 0&
End Enum

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type INTERNET_CACHE_ENTRY_INFO
    dwStructSize As Long
    lpszSourceUrlName As Long
    lpszLocalFileName As Long
    CacheEntryType  As Long         '// Type of entry returned
    dwUseCount As Long
    dwHitRate As Long
    dwSizeLow As Long
    dwSizeHigh As Long
    LastModifiedTime As FILETIME
    ExpireTime As FILETIME
    LastAccessTime As FILETIME
    LastSyncTime As FILETIME
    lpHeaderInfo As Long
    dwHeaderInfoSize As Long
    lpszFileExtension As Long
    dwExemptDelta  As Long
End Type

'// --------------------------Internet Cache API
Private Declare Function FindFirstUrlCacheEntry Lib "Wininet.dll" Alias "FindFirstUrlCacheEntryA" (ByVal lpszUrlSearchPattern As String, lpFirstCacheEntryInfo As Any, lpdwFirstCacheEntryInfoBufferSize As Long) As Long
Private Declare Function FindNextUrlCacheEntry Lib "Wininet.dll" Alias "FindNextUrlCacheEntryA" (ByVal hEnumHandle As Long, lpNextCacheEntryInfo As Any, lpdwNextCacheEntryInfoBufferSize As Long) As Long
Private Declare Function FindCloseUrlCache Lib "Wininet.dll" (ByVal hEnumHandle As Long) As Long
Private Declare Function DeleteUrlCacheEntry Lib "Wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long

'// --------------------------Memory API
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Function lstrcpyA Lib "kernel32" (ByVal RetVal As String, ByVal Ptr As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal Ptr As Any) As Long

Private Function InternetDeleteCache(sCacheFile As String) As Boolean
    InternetDeleteCache = CBool(DeleteUrlCacheEntry(sCacheFile))
End Function

Private Function InternetCacheList(Optional eFilterType As eCacheType = eNormal) As Variant
    Dim ICEI As INTERNET_CACHE_ENTRY_INFO
    Dim lhFile As Long, lBufferSize As Long, lptrBuffer As Long
    Dim sCacheFile As String
    Dim asURLs() As String, lNumEntries As Long

    'Determine required buffer size
    lBufferSize = 0
    lhFile = FindFirstUrlCacheEntry(0&, ByVal 0&, lBufferSize)
   
    If (lhFile = ERROR_CACHE_FIND_FAIL) And (Err.LastDllError = ERROR_INSUFFICIENT_BUFFER) Then
    
        'Allocate memory for ICEI structure
        lptrBuffer = LocalAlloc(LMEM_FIXED, lBufferSize)
    
        If lptrBuffer Then
         
            'Set a Long pointer to the memory location
            CopyMemory ByVal lptrBuffer, lBufferSize, 4
            
            'Call first find API passing it the pointer to the allocated memory
            lhFile = FindFirstUrlCacheEntry(vbNullString, ByVal lptrBuffer, lBufferSize)        '1 = success
            
            If lhFile <> ERROR_CACHE_FIND_FAIL Then
            
                'Loop through the cache
                Do
                    'Copy data back to structure
                    CopyMemory ICEI, ByVal lptrBuffer, Len(ICEI)
                
                    If ICEI.CacheEntryType And eFilterType Then
                        sCacheFile = StrFromPtrA(ICEI.lpszSourceUrlName)
                        lNumEntries = lNumEntries + 1
                        If lNumEntries = 1 Then
                            ReDim asURLs(1 To 1)
                        Else
                            ReDim Preserve asURLs(1 To lNumEntries)
                        End If
                        asURLs(lNumEntries) = sCacheFile
                    End If
                
                    'Free memory associated with the last-retrieved file
                    Call LocalFree(lptrBuffer)
                    
                    'Call FindNextUrlCacheEntry with buffer size set to 0.
                    'Call will fail and return required buffer size.
                    lBufferSize = 0
                    Call FindNextUrlCacheEntry(lhFile, ByVal 0&, lBufferSize)
                    
                    'Allocate and assign the memory to the pointer
                    lptrBuffer = LocalAlloc(LMEM_FIXED, lBufferSize)
                    CopyMemory ByVal lptrBuffer, lBufferSize, 4&
                
                Loop While FindNextUrlCacheEntry(lhFile, ByVal lptrBuffer, lBufferSize)
            
            End If
        
        End If
    
    End If
   
    'Free memory
    Call LocalFree(lptrBuffer)
    Call FindCloseUrlCache(lhFile)
    InternetCacheList = asURLs
End Function

Private Function StrFromPtrA(ByVal lptrString As Long) As String
    'Create buffer
    StrFromPtrA = String$(lstrlenA(ByVal lptrString), 0)
    'Copy memory
    Call lstrcpyA(ByVal StrFromPtrA, ByVal lptrString)
End Function

Public Sub DeleteCookiesEntry(ByVal silent As Boolean)
    Dim avURLs As Variant, vThisValue As Variant
    
    On Error Resume Next
    'Return an array of all internet cache files
    avURLs = InternetCacheList
    For Each vThisValue In avURLs
        'Print files
        Debug.Print CStr(vThisValue)
    Next
    
    'Return the an array of all cookies
    avURLs = InternetCacheList(eCookie)
    
    If silent = True Then
    If MsgBox("Delete cookies?", vbQuestion + vbYesNo) = vbYes Then
        For Each vThisValue In avURLs
            'Delete cookies
            InternetDeleteCache CStr(vThisValue)
            'Debug.Print "Deleted " & vThisValue
        Next
    Else
        'For Each vThisValue In avURLs
            'Print cookie; Files
            'Debug.Print vThisValue
        'Next
    End If
    Else
    For Each vThisValue In avURLs
            '// Delete cookies
            InternetDeleteCache CStr(vThisValue)
            'Debug.Print "Deleted " & vThisValue
        Next
    End If
End Sub
