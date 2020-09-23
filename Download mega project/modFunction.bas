Attribute VB_Name = "modFunction"
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

'/// Get File size
'/// *****************************************************************
Private Const K_B = 1024#
Private Const M_B = (K_B * 1024#) ' MegaBytes
Private Const G_B = (M_B * 1024#) ' GigaBytes
Private Const T_B = (G_B * 1024#) ' TeraBytes
Private Const P_B = (T_B * 1024#) ' PetaBytes
Private Const E_B = (P_B * 1024#) ' ExaBytes
Private Const Z_B = (E_B * 1024#) ' ZettaBytes
Private Const Y_B = (Z_B * 1024#) ' YottaBytes

Public Enum DISP_BYTES_FORMAT
    DISP_BYTES_LONG
    DISP_BYTES_SHORT
    DISP_BYTES_ALL
End Enum

'/// BrowserForFolders
'/// *****************************************************************
Private Type BrowseInfo
    hwndOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Const BIF_BROWSEINCLUDEURLS = 128
Private Const BIF_EDITBOX = 16
Private Const BIF_NEWDIALOGSTYLE = 64
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_STATUSTEXT = 4
Private Const BIF_VALIDATE = 32
Public Const MAX_PATH = 260

Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal hMem As Long)
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SendMessageA Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_VALIDATEFAILEDA = 3
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BFFM_SETSELECTION = (WM_USER + 102)

Dim m_StartFolder As String
Dim bValidateFailed As Boolean

'/// Play Sound Resource
'/// *****************************************************************
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function PlaySoundData Lib "winmm.dll" Alias "PlaySoundA" (lpData As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long

Private Const SND_SYNC = &H0
Private Const SND_ASYNC = &H1
Private Const SND_NODEFAULT = &H2
Private Const SND_MEMORY = &H4
Private Const SND_ALIAS = &H10000
Private Const SND_FILENAME = &H20000
Private Const SND_RESOURCE = &H40004
Private Const SND_ALIAS_ID = &H110000
Private Const SND_ALIAS_START = 0
Private Const SND_LOOP = &H8
Private Const SND_NOSTOP = &H10
Private Const SND_VALID = &H1F
Private Const SND_NOWAIT = &H2000
Private Const SND_VALIDFLAGS = &H17201F
Private Const SND_RESERVED = &HFF000000
Private Const SND_TYPE_MASK = &H170007

Private Const WAVERR_BASE = 32
Private Const WAVERR_BADFORMAT = (WAVERR_BASE + 0)
Private Const WAVERR_STILLPLAYING = (WAVERR_BASE + 1)
Private Const WAVERR_UNPREPARED = (WAVERR_BASE + 2)
Private Const WAVERR_SYNC = (WAVERR_BASE + 3)
Private Const WAVERR_LASTERROR = (WAVERR_BASE + 3)

Private m_snd() As Byte
Public Function FormatFileSize(ByVal dblFileSize As Double, Optional ByVal strFormatMask As String) As String
On Error Resume Next
Select Case dblFileSize
    Case 0 To 1023 ' Bytes
        FormatFileSize = Format(dblFileSize) & " bytes"
    Case 1024 To 1048575 ' KB
        If strFormatMask = Empty Then strFormatMask = "###0"
        FormatFileSize = Format(dblFileSize / 1024#, strFormatMask) & " KB"
    Case 1024# ^ 2 To 1073741823 ' MB
        If strFormatMask = Empty Then strFormatMask = "###0.0"
        FormatFileSize = Format(dblFileSize / (1024# ^ 2), strFormatMask) & " MB"
    Case Is > 1073741823# ' GB
        If strFormatMask = Empty Then strFormatMask = "###0.0"
        FormatFileSize = Format(dblFileSize / (1024# ^ 3), strFormatMask) & " GB"
End Select
End Function

Public Function FormatPercentage(ByVal dblFileSize As Double, Optional ByVal strFormatMask As String) As String
On Error Resume Next
Select Case dblFileSize
    Case 0 To 1023
        FormatPercentage = Format(dblFileSize)
    Case 1024 To 1048575
        If strFormatMask = Empty Then strFormatMask = "###0"
        FormatPercentage = Format(dblFileSize / 1024#, strFormatMask)
    Case 1024# ^ 2 To 1073741823
        If strFormatMask = Empty Then strFormatMask = "###0.0"
        FormatPercentage = Format(dblFileSize / (1024# ^ 2), strFormatMask)
    Case Is > 1073741823#
        If strFormatMask = Empty Then strFormatMask = "###0.0"
        FormatPercentage = Format(dblFileSize / (1024# ^ 3), strFormatMask)
End Select
End Function

Public Function FormatTime(ByVal sglTime As Single) As String
On Error Resume Next
Select Case sglTime
    Case 0 To 59
        FormatTime = Format(Int(sglTime / 60), "#0") & ":" & Format(sglTime Mod 60, "#0")
    Case 60 To 3599
        FormatTime = "0" & Format(Int(sglTime / 60), "#0") & ":" & Format(sglTime Mod 60, "#00") & ":" & Format(sglTime Mod 60, "#00")
    Case Else
        FormatTime = "0" & Format(Int(sglTime / 3600), "#0") & ":" & Format(sglTime / 60 Mod 60, "#00") & ":" & Format(Int(sglTime / 60), "#00")
End Select
End Function

Public Function GetSizeBytes(Dec As Variant, Optional DispBytesFormat As DISP_BYTES_FORMAT = DISP_BYTES_ALL) As String
    Dim DispLong As String: Dim DispShort As String: Dim s As String
    If DispBytesFormat <> DISP_BYTES_SHORT Then DispLong = FormatNumber(Dec, 0) & " bytes" Else DispLong = ""
    If DispBytesFormat <> DISP_BYTES_LONG Then
        If Dec > Y_B Then
            DispShort = FormatNumber(Dec / Y_B, 2) & " Yb"
        ElseIf Dec > Z_B Then
            DispShort = FormatNumber(Dec / Z_B, 2) & " Zb"
        ElseIf Dec > E_B Then
            DispShort = FormatNumber(Dec / E_B, 2) & " Eb"
        ElseIf Dec > P_B Then
            DispShort = FormatNumber(Dec / P_B, 2) & " Pb"
        ElseIf Dec > T_B Then
            DispShort = FormatNumber(Dec / T_B, 2) & " Tb"
        ElseIf Dec > G_B Then
            DispShort = FormatNumber(Dec / G_B, 2) & " Gb"
        ElseIf Dec > M_B Then
            DispShort = FormatNumber(Dec / M_B, 2) & " Mb"
        ElseIf Dec > K_B Then
            DispShort = FormatNumber(Dec / K_B, 2) & " Kb"
        Else
            DispShort = FormatNumber(Dec, 0) & " bytes"
        End If
    Else
        DispShort = ""
    End If
    Select Case DispBytesFormat
        Case DISP_BYTES_SHORT:
            GetSizeBytes = DispShort
        Case DISP_BYTES_LONG:
            GetSizeBytes = DispLong
        Case Else:
            GetSizeBytes = DispLong & " (" & DispShort & ")"
    End Select
End Function

Public Function MakeDirectory(szDirectory As String) As Boolean
Dim strFolder As String
Dim szRslt As String
On Error GoTo IllegalFolderName
If Right$(szDirectory, 1) <> "\" Then szDirectory = szDirectory & "\"
strFolder = szDirectory
szRslt = Dir(strFolder, 63)
While szRslt = ""
    DoEvents
    szRslt = Dir(strFolder, 63)
    strFolder = Left$(strFolder, Len(strFolder) - 1)
    If strFolder = "" Then GoTo IllegalFolderName
Wend
If Right$(strFolder, 1) <> "\" Then strFolder = strFolder & "\"
While strFolder <> szDirectory
    strFolder = Left$(szDirectory, Len(strFolder) + 1)
    If Right$(strFolder, 1) = "\" Then MkDir strFolder
Wend
MakeDirectory = True
Exit Function
IllegalFolderName:
        MakeDirectory = False
    Err.Clear
End Function

Private Function BrowseCallbackProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
    Dim lpIDList As Long
    Dim Ret As Long
    Dim sBuffer As String
    On Local Error Resume Next
    Select Case uMsg
        Case BFFM_INITIALIZED
            SendMessageA hwnd, BFFM_SETSELECTION, 1, m_StartFolder
        Case BFFM_SELCHANGED
            sBuffer = Space(MAX_PATH)
            Ret = SHGetPathFromIDList(lp, sBuffer)
            If Ret = 1 Then
                SendMessageA hwnd, BFFM_SETSTATUSTEXT, 0, sBuffer
            End If
        Case BFFM_VALIDATEFAILEDA
            bValidateFailed = True
    End Select
    BrowseCallbackProc = 0
End Function

Public Function BrowseForFolder(ByVal hwndOwner As Long, ByVal Prompt As String, Optional ByVal StartFolder) As String
    Dim lNull As Long
    Dim lpIDList As Long
    Dim lResult As Long
    Dim sPath As String
    Dim udtBI As BrowseInfo
    On Local Error Resume Next
    With udtBI
        .hwndOwner = hwndOwner
        .lpszTitle = Prompt
        .ulFlags = BIF_BROWSEINCLUDEURLS Or BIF_NEWDIALOGSTYLE Or BIF_EDITBOX Or BIF_VALIDATE Or BIF_RETURNONLYFSDIRS Or BIF_STATUSTEXT
        If Not IsMissing(StartFolder) Then
            m_StartFolder = StartFolder
            If Right$(m_StartFolder, 1) <> Chr$(0) Then m_StartFolder = m_StartFolder & Chr$(0)
            .lpfnCallback = GetAddressofFunction(AddressOf BrowseCallbackProc)
        End If
    End With
    bValidateFailed = False
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList And Not bValidateFailed Then
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        CoTaskMemFree lpIDList
        lNull = InStr(sPath, vbNullChar)
        If lNull Then
            sPath = Left$(sPath, lNull - 1)
        End If
    End If
    BrowseForFolder = sPath
End Function

Private Function GetAddressofFunction(Add As Long) As Long
    On Local Error Resume Next
    GetAddressofFunction = Add
End Function

Public Function SileNtSound()
    On Error Resume Next
    sndPlaySound ByVal vbNullString, 0&
End Function

Public Function PlaySoundResource(ByVal SndID As Long, Optional sndType As String = "WAVE") As Long
   Const Flags = SND_MEMORY Or SND_ASYNC Or SND_NODEFAULT
   '/// USAGE: Call PlaySoundResource(101)
   '/// *****************************************************************
   On Error GoTo ErrorHandler
   DoEvents
   m_snd = LoadResData(SndID, sndType)
   PlaySoundResource = PlaySoundData(m_snd(0), 0, Flags)
Exit Function
ErrorHandler:
    Err.Clear
End Function
