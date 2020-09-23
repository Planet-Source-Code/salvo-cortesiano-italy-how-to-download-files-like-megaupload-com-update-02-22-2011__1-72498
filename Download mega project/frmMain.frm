VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "##"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10470
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   10470
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBrowserFolders 
      Caption         =   "..."
      Height          =   285
      Left            =   9765
      TabIndex        =   42
      ToolTipText     =   "Browser Folder Downloads..."
      Top             =   6375
      Width           =   510
   End
   Begin VB.TextBox txtPathDownloads 
      Height          =   300
      Left            =   1740
      Locked          =   -1  'True
      TabIndex        =   40
      Text            =   "n.a"
      Top             =   6375
      Width           =   7935
   End
   Begin VB.CommandButton cmdAllURLs 
      Caption         =   "Get All URL's"
      Height          =   345
      Left            =   6555
      TabIndex        =   37
      ToolTipText     =   "Get All URL of the list..."
      Top             =   7575
      Width           =   1800
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   7935
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.ListBox lsts 
      Height          =   270
      Index           =   2
      Left            =   9150
      TabIndex        =   36
      Top             =   330
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.ListBox lsts 
      Height          =   270
      Index           =   1
      Left            =   9060
      TabIndex        =   35
      Top             =   360
      Visible         =   0   'False
      Width           =   1170
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   8490
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.ComboBox cLanguage 
      Height          =   330
      ItemData        =   "frmMain.frx":23D2
      Left            =   9750
      List            =   "frmMain.frx":23F7
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   45
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.ListBox lsts 
      Height          =   270
      Index           =   3
      Left            =   9270
      TabIndex        =   29
      Top             =   30
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Close Program"
      Height          =   360
      Left            =   8580
      TabIndex        =   28
      ToolTipText     =   "Close this Application..."
      Top             =   7575
      Width           =   1800
   End
   Begin VB.Frame Frame8 
      Caption         =   "Status Download File(s)"
      Height          =   2235
      Left            =   4500
      TabIndex        =   21
      Top             =   3420
      Width           =   5925
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   1995
         Left            =   30
         ScaleHeight     =   1995
         ScaleWidth      =   5820
         TabIndex        =   22
         Top             =   195
         Width           =   5820
         Begin ComctlLib.ProgressBar PB 
            Height          =   210
            Left            =   60
            TabIndex        =   23
            Top             =   105
            Width           =   4965
            _ExtentX        =   8758
            _ExtentY        =   370
            _Version        =   327682
            Appearance      =   1
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Of:"
            Height          =   270
            Left            =   90
            TabIndex        =   46
            Top             =   885
            Width           =   1800
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Download byte's:"
            Height          =   270
            Left            =   90
            TabIndex        =   45
            Top             =   600
            Width           =   1800
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "File Size:"
            Height          =   270
            Left            =   75
            TabIndex        =   44
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label lblWorkStatus 
            Caption         =   "Ready..."
            Height          =   750
            Left            =   90
            TabIndex        =   31
            Top             =   1215
            Width           =   5685
         End
         Begin VB.Line Line5 
            BorderColor     =   &H80000010&
            BorderStyle     =   2  'Dash
            X1              =   90
            X2              =   5760
            Y1              =   1170
            Y2              =   1170
         End
         Begin VB.Label lblPercentage 
            Caption         =   "0%"
            Height          =   270
            Left            =   5160
            TabIndex        =   27
            Top             =   90
            Width           =   615
         End
         Begin VB.Label lblSaved 
            Caption         =   "##"
            ForeColor       =   &H000000C0&
            Height          =   270
            Left            =   2025
            TabIndex        =   26
            Top             =   600
            Width           =   3735
         End
         Begin VB.Label lblSize 
            Caption         =   "##"
            ForeColor       =   &H000000C0&
            Height          =   270
            Left            =   2025
            TabIndex        =   25
            Top             =   360
            Width           =   3720
         End
         Begin VB.Label lblOf 
            Caption         =   "##"
            ForeColor       =   &H000000C0&
            Height          =   270
            Left            =   2040
            TabIndex        =   24
            Top             =   900
            Width           =   3735
         End
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "File Download URL"
      Height          =   630
      Left            =   45
      TabIndex        =   15
      Top             =   5685
      Width           =   10395
      Begin VB.TextBox txtDownloadURL 
         Height          =   300
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "n.a"
         Top             =   210
         Width           =   7005
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   7200
         ScaleHeight     =   465
         ScaleWidth      =   3165
         TabIndex        =   16
         Top             =   135
         Width           =   3165
         Begin VB.CommandButton cmdDownload 
            Caption         =   "Download"
            Height          =   315
            Left            =   75
            TabIndex        =   19
            ToolTipText     =   "Start the Download..."
            Top             =   60
            Width           =   1125
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save URL"
            Height          =   315
            Left            =   1305
            TabIndex        =   18
            ToolTipText     =   "Save this URL into List..."
            Top             =   60
            Width           =   1125
         End
         Begin VB.CommandButton cmdCopy 
            Caption         =   "(C)"
            Height          =   315
            Left            =   2520
            TabIndex        =   17
            ToolTipText     =   "Copy URL to Clipboard..."
            Top             =   60
            Width           =   570
         End
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Info of File"
      Height          =   2220
      Left            =   45
      TabIndex        =   9
      Top             =   3435
      Width           =   4410
      Begin VB.TextBox txtDownloadLink 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   1635
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "n.a"
         Top             =   1095
         Width           =   2640
      End
      Begin VB.TextBox txtFileName 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "n.a"
         Top             =   540
         Width           =   4170
      End
      Begin VB.TextBox txtSize 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "n.a"
         Top             =   1095
         Width           =   1440
      End
      Begin VB.TextBox txtDscription 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   630
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   10
         Text            =   "frmMain.frx":2434
         Top             =   1500
         Width           =   4170
      End
      Begin VB.Label lbls 
         Caption         =   "Link di Download:"
         Height          =   255
         Index           =   3
         Left            =   1650
         TabIndex        =   39
         Top             =   870
         Width           =   2550
      End
      Begin VB.Label lbls 
         Caption         =   "File Name:"
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   14
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label lbls 
         Caption         =   "Size:"
         Height          =   255
         Index           =   4
         Left            =   135
         TabIndex        =   13
         Top             =   870
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Multy URL's"
      Height          =   2025
      Left            =   45
      TabIndex        =   7
      Top             =   1335
      Width           =   10380
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   60
         ScaleHeight     =   375
         ScaleWidth      =   4575
         TabIndex        =   47
         Top             =   1605
         Width           =   4575
         Begin VB.CommandButton cmdClearList 
            Caption         =   "Clear"
            Height          =   315
            Left            =   1620
            TabIndex        =   50
            ToolTipText     =   "Clear the List URL's"
            Top             =   30
            Width           =   1095
         End
         Begin VB.CommandButton cmdGetThisURL 
            Caption         =   "Get this URL"
            Height          =   315
            Left            =   2820
            TabIndex        =   49
            ToolTipText     =   "Get Info of Selected URL"
            Top             =   30
            Width           =   1635
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Paste URL's"
            Height          =   315
            Left            =   60
            TabIndex        =   48
            ToolTipText     =   "Paste URL's from Clipboard..."
            Top             =   30
            Width           =   1485
         End
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1830
         Left            =   4665
         ScaleHeight     =   1830
         ScaleWidth      =   5670
         TabIndex        =   32
         Top             =   135
         Width           =   5670
         Begin VB.Frame Frame10 
            Caption         =   "Downloads List Links"
            Height          =   1785
            Left            =   0
            TabIndex        =   33
            Top             =   0
            Width           =   5625
            Begin VB.ListBox lsts 
               Height          =   1320
               Index           =   0
               Left            =   105
               TabIndex        =   34
               Top             =   285
               Width           =   5430
            End
         End
      End
      Begin VB.ListBox lstURLs 
         Height          =   1320
         Left            =   75
         TabIndex        =   8
         Top             =   255
         Width           =   4545
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "URL"
      Height          =   615
      Left            =   45
      TabIndex        =   2
      Top             =   705
      Width           =   10380
      Begin VB.TextBox txtURL 
         Height          =   300
         Left            =   120
         TabIndex        =   5
         Top             =   225
         Width           =   7110
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7215
         ScaleHeight     =   420
         ScaleWidth      =   3120
         TabIndex        =   3
         Top             =   150
         Width           =   3120
         Begin VB.CommandButton cmdGO 
            Caption         =   "Get this URL"
            Height          =   345
            Left            =   1230
            TabIndex        =   6
            ToolTipText     =   "Get this URL of megaupload..."
            Top             =   45
            Width           =   1800
         End
         Begin VB.CommandButton cmdPaste 
            Caption         =   "Paste"
            Height          =   315
            Left            =   165
            TabIndex        =   4
            Top             =   45
            Width           =   915
         End
      End
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Delete Cookies"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   75
      TabIndex        =   43
      ToolTipText     =   "Clear ALL Cookies..."
      Top             =   7755
      Width           =   1575
   End
   Begin VB.Label lbls 
      Caption         =   "Downloads Path:"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   41
      Top             =   6405
      Width           =   1695
   End
   Begin VB.Image imgs 
      Height          =   240
      Index           =   0
      Left            =   0
      Picture         =   "frmMain.frx":2438
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgs 
      Height          =   240
      Index           =   1
      Left            =   0
      Picture         =   "frmMain.frx":29C2
      Top             =   180
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgs 
      Height          =   240
      Index           =   2
      Left            =   0
      Picture         =   "frmMain.frx":2F4C
      Top             =   405
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "##"
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   4965
      TabIndex        =   1
      Top             =   75
      Width           =   5445
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "##"
      ForeColor       =   &H00808080&
      Height          =   225
      Left            =   4920
      TabIndex        =   0
      Top             =   435
      Width           =   5505
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   -45
      Picture         =   "frmMain.frx":34D6
      Top             =   -45
      Width           =   5010
   End
   Begin VB.Shape sp1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000003&
      Height          =   720
      Left            =   -15
      Top             =   -45
      Width           =   10785
   End
   Begin VB.Menu mnufile 
      Caption         =   "file"
      Visible         =   0   'False
      Begin VB.Menu mnuList 
         Caption         =   "Delete selected URL"
         Index           =   0
      End
      Begin VB.Menu mnuList 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuList 
         Caption         =   "Delete all URL's of this List"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private fln As String
Private str_Folder As String
Private Const sAppTitle As String = "Megaupload Download v1.0.1"

Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private Declare Function InternetCloseHandle Lib "wininet" (ByVal hInet As Long) As Long
Private Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)
Private Declare Function DeleteUrlCacheEntry Lib "Wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long

Private problemFile As Boolean
Dim returnString As String

Private Enum LANGUAGE
    Italian = 0
    England = 1
    French = 2
    Spanish = 3
    Dutch = 4
    [Fraktur (Old German)] = 5
    German = 6
    Bangla = 7
    Basque = 8
    [Portuguese (Brazilian)] = 9
    Vietnamese = 10
End Enum

Public Enum SW_SHOW_MODE
    SW_HIDE = 0
    SW_MAXIMIZE = 3
    SW_MINIMIZE = 6
    SW_NORMAL = 1
End Enum

'/// Verify if the File exist
'/// *****************************************************************
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Const GENERIC_READ As Long = &H80000000
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const OPEN_EXISTING As Long = 3
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const MAX_PATH = 260

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
Private Function BrowseFolder(ByVal strTitle As String, Optional strPath As String = "") As String
    Dim Folder As String
    On Local Error GoTo ErrorHandler
    If strPath = "" Then
        strPath = App.Path + "\"
    Else
        If Right$(strPath, 1) <> "\" Then strPath = strPath + "\"
    End If
    Folder = BrowseForFolder(Me.hwnd, strTitle, strPath)
    If Folder <> "" Then BrowseFolder = Folder Else BrowseFolder = ""
Exit Function
ErrorHandler:
        BrowseFolder = "Error!"
    Err.Clear
End Function

Private Sub DelayTime(ByVal Second As Long, Optional ByVal Refresh As Boolean = False)
    On Error Resume Next
    Dim Start As Date
    Start = Now
    Do
    If Refresh Then DoEvents
    Loop Until DateDiff("s", Start, Now) >= Second
End Sub

Private Function DownloadFile(Link As String, Filename As String) As Boolean
    On Error GoTo ErrorHandler
    Dim FileSize As Long
    Dim sz As Double
    Dim FileRemaining As Long
    Dim FileNumber As Integer
    Dim FileData() As Byte
    Dim FileSize_Current As Long
    Dim PBValue As Integer
    
    Inet.Execute Trim(Link), "GET"
    Do While Inet.StillExecuting
        DoEvents
    Loop
    
    fln = Filename
    
    FileSize = Inet.GetHeader("Content-Length")
    lblSize.Caption = GetSizeBytes(FileSize, DISP_BYTES_SHORT)
    FileRemaining = FileSize
    FileSize_Current = 0
    
    FileNumber = FreeFile
    Open Filename For Binary Access Write As #FileNumber
    
    PB.Max = 100
    
    Do Until FileRemaining = 0
        DoEvents
        If frmMain.Tag = "Cancel" Then
                Inet.Cancel
                        MsgBox "Download of the file, abort by user!", vbInformation, App.Title
                    GoTo Reset
            Exit Function
        End If
        
        If FileRemaining > 1024 Then
            FileData = Inet.GetChunk(1024, icByteArray)
            FileRemaining = FileRemaining - 1024
        Else
            FileData = Inet.GetChunk(FileRemaining, icByteArray)
            FileRemaining = 0
        End If
        
        FileSize_Current = FileSize - FileRemaining
        PBValue = CInt((100 / FileSize) * FileSize_Current)
        lblSaved.Caption = GetSizeBytes(FileSize_Current, DISP_BYTES_SHORT)
        lblOf.Caption = GetSizeBytes(FileSize - FileSize_Current, DISP_BYTES_SHORT)
        lblPercentage.Caption = PBValue & " % "
        PB.Value = PBValue
        Put #FileNumber, , FileData
        DoEvents
    Loop
    
    Close #FileNumber
    
    DownloadFile = True
    frmMain.Tag = ""
    lblPercentage.Caption = "0%"
    cmdDownload.Caption = "Download"
    On Error Resume Next
    PB.Value = 0
    
    Exit Function

Reset:
cmdDownload.Caption = "Download"
On Error Resume Next
    If FileExists(Filename) Then
        Close #FileNumber
        Call Kill(Filename)
    End If
    PB.Value = 0
Exit Function
ErrorHandler:
    DownloadFile = False
    GoTo Reset
    Err.Clear
Exit Function
End Function

Private Function FileExists(sSource As String) As Boolean
   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   On Error Resume Next
   hFile = FindFirstFile(sSource, WFD)
   FileExists = hFile <> INVALID_HANDLE_VALUE
   Call FindClose(hFile)
End Function

Private Function FileExtensionFromPath(strPath As String) As String
    On Local Error GoTo ErrorHandler
    FileExtensionFromPath = Right$(strPath, (Len(strPath) - InStrRev(strPath, ".")) + 1)
Exit Function
ErrorHandler:
        FileExtensionFromPath = ""
    Err.Clear
End Function

Private Function FileNameFromPath(strPath As String) As String
    On Local Error GoTo ErrorHandler
    FileNameFromPath = Right$(strPath, Len(strPath) - InStrRev(strPath, "\"))
Exit Function
ErrorHandler:
        FileNameFromPath = ""
    Err.Clear
End Function

Private Sub GrabInfoFile(bodyText As String)
    Dim IX As Integer
    Dim htmlBody As String
    Dim ssT As String
    Dim pos1, pos2 As Long
    Dim TempArray(20) As String
    On Local Error GoTo ErrorHandler
    
    htmlBody = bodyText
    
    problemFile = False
    
    '/// Invalid Link
    If InStr(htmlBody, "Unfortunately, the link you have clicked is not available") > 0 Then
        lblWorkStatus.Caption = "Unfortunately, the link you have clicked is not available! Reasons for this may include: -Invalid link or the file has been deleted because it was violating our Terms of service!"
        problemFile = True
    ElseIf InStr(htmlBody, "Purtroppo il link che hai scelto") > 0 Then
        lblWorkStatus.Caption = "Unfortunately, the link you have clicked is not available! Reasons for this may include: -Invalid link or the file has been deleted because it was violating our Terms of service!"
        problemFile = True
    End If
    
    '/// File > 1GB Account Premium Required
    If InStr(htmlBody, "Il file che si sta tentando di scaricare") > 0 Then
        lblWorkStatus.Caption = "Il file che si sta tentando di scaricare è maggiore di 1GB" _
        & ". Solo gli utenti Premium possono scaricare tale file!"
        problemFile = True
    Else
        TempArray(6) = "n.a"
    End If
    
    If TempArray(6) = "n.a" Then
    '/// File > 1GB Account Premium Required
    If InStr(htmlBody, "The file that you're trying to download is larger") > 0 Then
        lblWorkStatus.Caption = "The file that you're trying to download is larger" _
        & ". Only a Premium member Download this file!"
        problemFile = True
    Else
        TempArray(6) = "n.a"
    End If
    End If
    
        '/// Retrive File Name
        If InStr(htmlBody, "Nome file:") > 0 Then
            pos1 = InStr(pos1 + 1, htmlBody, "Nome file:</span> <span class=""" & "down_txt2", vbTextCompare) + Len("Nome file:</span> <span class=""" & "down_txt2")
            pos2 = InStr(pos1 + 1, htmlBody, "<", vbTextCompare)
            TempArray(3) = Mid$(htmlBody, pos1, pos2 - pos1)
            TempArray(3) = Replace(TempArray(3), ">", "")
            TempArray(3) = Replace(TempArray(3), """", "")
        Else
            TempArray(3) = "n.a"
        End If
        
        If TempArray(3) = "n.a" Then
        If InStr(htmlBody, "File name:") > 0 Then
            pos1 = InStr(pos1 + 1, htmlBody, "File name:</span> <span class=""" & "down_txt2", vbTextCompare) + Len("File name:</span> <span class=""" & "down_txt2")
            pos2 = InStr(pos1 + 1, htmlBody, "<", vbTextCompare)
            TempArray(3) = Mid$(htmlBody, pos1, pos2 - pos1)
            TempArray(3) = Replace(TempArray(3), ">", "")
            TempArray(3) = Replace(TempArray(3), """", "")
        Else
            TempArray(3) = "n.a"
        End If
        End If
        
        txtFileName.Text = TempArray(3)
    
    '/// Retrive File Description
    If InStr(htmlBody, "Descrizione file:") > 0 Then
        pos1 = InStr(pos1 + 1, htmlBody, "Descrizione file:</strong>", vbTextCompare) _
                                + Len("Descrizione file:</strong>")
        pos2 = InStr(pos1 + 1, htmlBody, "<", vbTextCompare)
        TempArray(4) = Mid$(htmlBody, pos1, pos2 - pos1)
    Else
        TempArray(4) = "n.a"
    End If
    
    If TempArray(4) = "n.a" Then
    '/// Retrive File Description
    If InStr(htmlBody, "File description:") > 0 Then
        pos1 = InStr(pos1 + 1, htmlBody, "File description:</strong>", vbTextCompare) _
                                + Len("File description:</strong>")
        pos2 = InStr(pos1 + 1, htmlBody, "<", vbTextCompare)
        TempArray(4) = Mid$(htmlBody, pos1, pos2 - pos1)
    Else
        TempArray(4) = "n.a"
    End If
    End If
    
    txtDscription.Text = TempArray(4)
    
    '/// Retrive File Size
    If InStr(htmlBody, "Dimensione file:") > 0 Then
        pos1 = InStr(pos1 + 1, htmlBody, "Dimensione file:</strong>", vbTextCompare) _
                                + Len("Dimensione file:</strong>")
        pos2 = InStr(pos1 + 1, htmlBody, "<", vbTextCompare)
        TempArray(5) = Mid$(htmlBody, pos1, pos2 - pos1)
    Else
        TempArray(5) = "n.a"
    End If
    
    If TempArray(5) = "n.a" Then
    '/// Retrive File Size
    If InStr(htmlBody, "File size:") > 0 Then
        pos1 = InStr(pos1 + 1, htmlBody, "File size:</strong>", vbTextCompare) _
                                + Len("File size:</strong>")
        pos2 = InStr(pos1 + 1, htmlBody, "<", vbTextCompare)
        TempArray(5) = Mid$(htmlBody, pos1, pos2 - pos1)
    Else
        TempArray(5) = "n.a"
    End If
    End If
    
    txtSize.Text = TempArray(5)
    
    '/// Retrive Download Link
    If InStr(htmlBody, "Download link:") > 0 Then
        pos1 = InStr(pos1 + 1, htmlBody, "Download link:</span> <a href=", vbTextCompare) _
                                + Len("Download link:</span> <a href=")
        pos2 = InStr(pos1 + 1, htmlBody, """", vbTextCompare)
        TempArray(6) = Mid$(htmlBody, pos1, pos2 - pos1)
        TempArray(6) = Replace(TempArray(6), """", "")
    Else
        TempArray(6) = "n.a"
    End If
    
    If TempArray(6) = "n.a" Then
    If InStr(htmlBody, "Link download:") > 0 Then
        pos1 = InStr(pos1 + 1, htmlBody, "Link download:</span> <a href=", vbTextCompare) _
                                + Len("Link download:</span> <a href=")
        pos2 = InStr(pos1 + 1, htmlBody, """", vbTextCompare)
        TempArray(6) = Mid$(htmlBody, pos1, pos2 - pos1)
        TempArray(6) = Replace(TempArray(6), """", "")
    Else
        TempArray(6) = "n.a"
    End If
    End If
    
    txtDownloadLink.Text = TempArray(6)
    
Exit Sub
ErrorHandler:
        TempArray(3) = "n.a": TempArray(4) = "n.a": TempArray(5) = "n.a"
        problemFile = True
    Err.Clear
End Sub

Private Sub ImportDownloadLinks()
    Dim strLines() As String
    Dim strLine As String
    Dim i As Integer
    Dim FSO As New FileSystemObject
    Dim fsoStream As TextStream
    Dim Fs As File
    Dim tmpString As String
    On Local Error GoTo ErrorHandler
    If FileExists(App.Path + "\Downloads_List_Links.txt") Then
    Set Fs = FSO.GetFile(App.Path + "\Downloads_List_Links.txt")
        Set fsoStream = Fs.OpenAsTextStream(ForReading)
        '/// Read the file line by line
        Do While Not fsoStream.AtEndOfStream
            strLine = fsoStream.ReadLine
            If strLine <> "" Then
                strLines = Split(strLine, "|")
                    For i = 0 To UBound(strLines)
                        'tmpString = tmpString & StripLeft(Mid$(strLines(i), 1, Len(strLines(i))), "|", True) & vbCrLf
                        lsts(i).AddItem StripLeft(Mid$(strLines(i), 1, Len(strLines(i))), "*", False)
                        
                        If Mid$(StripLeft(Mid$(strLines(i), 1, Len(strLines(i))), "*", False), 1, 27) _
                        = "http://www.megaupload.com/?" Then lstURLs.AddItem StripLeft(Mid$(strLines(i), 1, Len(strLines(i))), "*", False)

                        DoEvents
                    Next i
            End If
        Loop
            fsoStream.Close
        End If
    Set fsoStream = Nothing
    Set Fs = Nothing
    Set FSO = Nothing
Exit Sub
ErrorHandler:
        MsgBox "Error #" & Err.Number & ". " & Err.Description, vbExclamation, App.Title
    Err.Clear
End Sub


Private Function StripLeft(strString As String, strChar As String, Optional sLeftsRight As Boolean = True) As String
  On Local Error Resume Next
  Dim i As Integer
    If sLeftsRight Then
        For i = 1 To Len(strString)
            If Mid$(strString, i, 1) = strChar Then
                    StripLeft = Mid$(strString, 1, i - 1)
                Exit For
            End If
        Next
    Else
        For i = (Len(strString)) To 1 Step -1
        If Mid$(strString, i, 1) = strChar Then
                StripLeft = Mid$(strString, i + 1, Len(strString) - i + 1)
            Exit For
        End If
    Next
End If
End Function

Private Function StripString(ByVal sString As String, sChar As String) As String
    Dim i As Integer
    Dim sTmp As String
    On Error Resume Next
    sTmp = Mid(sString, i + 1, Len(sString))
    For i = 1 To Len(sTmp)
      If Mid(sTmp, i, 1) = sChar Then
        Exit For
    Else
        returnString = Mid(sString, i + 2, Len(sString))
    End If
Next
     StripString = Left(sTmp, i - 1)
End Function


Private Sub cmdAdd_Click()
    Dim ClpFmt
    Dim lstEmpty As Boolean
    Dim s() As String
    Dim tmpText As String
    Dim strURLs As String
    Dim i, x As Long
    
    On Local Error Resume Next
    
    If Clipboard.GetFormat(vbCFText) Then ClpFmt = ClpFmt + 1
    If Clipboard.GetFormat(vbCFBitmap) Then ClpFmt = ClpFmt + 2
    If Clipboard.GetFormat(vbCFDIB) Then ClpFmt = ClpFmt + 4
    If Clipboard.GetFormat(vbCFRTF) Then ClpFmt = ClpFmt + 8
    
    tmpText = Clipboard.GetText()
    
   Select Case ClpFmt
      Case 1 '/// Only TXT
        SetFocusField txtURL, True
        s = Split(tmpText, vbNewLine)
            For x = LBound(s) To UBound(s)
                If s(x) <> "" Then strURLs = s(x)
                If Mid$(strURLs, 1, 27) = "http://www.megaupload.com/?" Then
                    lstURLs.AddItem strURLs
                Else
                    
                End If
            Next x
            Call ChkLst(lstURLs)
      Case 2, 4, 6 '/// Only PICTURE
      Case 3, 5, 7 '/// TXT and PICTURE
      Case 8, 9 '/// TXT RTF
            SetFocusField txtURL, True
            
            s = Split(tmpText, vbNewLine)
            For x = LBound(s) To UBound(s)
                If s(x) <> "" Then strURLs = s(x)
                If Mid$(strURLs, 1, 27) = "http://www.megaupload.com/?" Then
                    lstURLs.AddItem strURLs
                Else
                    
                End If
            Next x
            Call ChkLst(lstURLs)
      Case Else '/// CLIPBOARD EMPTY
        MsgBox "The Clipboard is empty!", vbExclamation, App.Title
   End Select
    
End Sub

Private Sub cmdAllURLs_Click()
    Dim Captcha As String
    Dim i As Integer
    Dim e_xit As Integer
    Dim htmlBody As String
    Dim pos1, pos2 As Long
    Dim TempArray(20) As String
    
    On Local Error GoTo ErrorHandler
    
    If lstURLs.ListCount = 0 Then
            MsgBox "Nothing link's to Download in the list {URL's}...", vbExclamation, App.Title
        Exit Sub
    End If
    
    e_xit = 0
    i = 0
    
    '/// Reset All
    txtDownloadURL.Text = "n.a"
    txtFileName.Text = "n.a"
    txtSize.Text = "n.a"
    txtDscription.Text = "n.a"
    txtDownloadLink.Text = "n.a"
    lblSize.Caption = "##"
    lblSaved.Caption = "##"
    lblOf.Caption = "##"
    lblPercentage.Caption = "0%"
    lblWorkStatus.Caption = "Ready..."
    
    lblWorkStatus.Caption = "Please wait, i'm working now..."
    
    For i = 0 To lstURLs.ListCount
        lstURLs.Selected(i) = True
        
    DoEvents
    
    txtURL.Text = lstURLs.List(i)
    
    lblWorkStatus.Caption = "Get and {Save} the page URL..."
    
    '// Get the HTML Code
    htmlBody = Inet1.OpenURL(lstURLs.List(i))
    
    '// Wait to load complete page of this URL
    While Inet1.StillExecuting
        DoEvents
    Wend
    
    '// Save the Code of the HTML page
    '''Open App.Path + "\tmp.html" For Output As #1
    '''    Print #1, htmlBody
    '''Close #1
    
    lblWorkStatus.Caption = "Get and {Navigate} the page URL..."
    
    '// Get the INFO of the File
    Call GrabInfoFile(htmlBody)
    
    Do Until Mid$(TempArray(0), 1, 10) = "http://www"
    
    '/// Retrive the Download Link meybe baypass the COUNTDOWN?
    If InStr(htmlBody, "id=""" & "downloadlink") > 0 Then
        pos1 = InStr(pos1 + 1, htmlBody, "id=""" & "downloadlink""" & "><a href=", vbTextCompare) + Len("id=""" & "downloadlink""" & "><a href=")
        pos2 = InStr(pos1 + 1, htmlBody, """", vbTextCompare) '- 1
        TempArray(0) = Mid$(htmlBody, pos1, pos2 - pos1)
        TempArray(0) = Replace(TempArray(0), "href=", "")
        TempArray(0) = Replace(TempArray(0), """", "")
    Else
        TempArray(0) = "n.a"
    End If
    
    DoEvents
        If problemFile = True Then Exit Do
    Loop
    
    lblWorkStatus.Caption = "Start {Downloading} the file..." & txtFileName.Text
    
    '// Display the Download URL
    txtDownloadURL.Text = TempArray(0)
    
   '// Download the files
    cmdDownload_Click
    
        If i >= lstURLs.ListCount - 1 Then Exit For
        
        DoEvents
        
    '/// Clear the Cache entry
    Call DeleteUrlCacheEntry(lstURLs.List(i))
        
        DelayTime 5, False
        
        htmlBody = Empty
        TempArray(0) = Empty
        
    Next i
    
    htmlBody = Empty
    
    lblSize.Caption = "##"
    lblSaved.Caption = "##"
    lblOf.Caption = "##"
    lblPercentage.Caption = "0%"
    
    Exit Sub
    
ErrorHandler:
    If Err.Number <> 91 Then lblWorkStatus.Caption = "Error #" & Err.Number & ". " & Err.Description
    Err.Clear
End Sub

Private Sub cmdBrowserFolders_Click()
str_Folder = BrowseFolder("Select Downloads Folder:", str_Folder)
    If str_Folder = Empty Or str_Folder = "Error!" Then Exit Sub
    txtPathDownloads.Text = str_Folder & "\"
End Sub

Private Sub cmdClearList_Click()
    If lstURLs.ListCount = 0 Then Exit Sub
    If MsgBox("Clear all dato from the List?" & vbLf & vbLf & "Are you sure?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        lstURLs.Clear
    End If
End Sub

Private Sub cmdCopy_Click()
If txtDownloadURL.Text = Empty Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText txtDownloadURL
    MsgBox "URL Link copyed to Clipboard!", vbInformation, App.Title
End Sub

Private Sub cmdDownload_Click()
Dim FSO As New FileSystemObject
    If cmdDownload.Caption = "Download" Then
        If txtPathDownloads.Text = "n.a" Then cmdBrowserFolders_Click: Exit Sub
        
        If Not FSO.FolderExists(txtPathDownloads.Text) Then
            cmdBrowserFolders_Click
                Exit Sub
            Set FSO = Nothing
        Else
            Set FSO = Nothing
        End If
        
        frmMain.Tag = ""
        cmdDownload.Caption = "Abort"
    
    If DownloadFile(txtDownloadURL.Text, txtPathDownloads.Text + StripLeft(txtDownloadURL.Text, "/", False)) Then
        cmdDownload.Caption = "Download"
        'MsgBox "File 'Downloaded' success...", vbInformation, App.Title
    Else
        If frmMain.Tag = "" Then: 'MsgBox "Error to 'Download' the file!", vbExclamation, App.Title
    End If
    
    ElseIf cmdDownload.Caption = "Abort" Then
        frmMain.Tag = "Cancel"
        cmdDownload.Caption = "Download"
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdGetThisURL_Click()
    Dim htmlBody As String
    Dim pos1, pos2 As Long
    Dim TempArray(20) As String
    
    On Local Error GoTo ErrorHandler
    
    If lstURLs.ListCount = 0 Then Exit Sub
    
    '/// Clear the Cache entry
    If txtURL.Text <> Empty Then Call DeleteUrlCacheEntry(txtURL.Text)
    
    '/// Reset All
    txtDownloadURL.Text = "n.a"
    txtFileName.Text = "n.a"
    txtSize.Text = "n.a"
    txtDscription.Text = "n.a"
    txtDownloadLink.Text = "n.a"
    lblWorkStatus.Caption = "Ready..."
    
    '// Get the HTML Code
    htmlBody = Inet1.OpenURL(lstURLs.List(lstURLs.ListIndex))
    
    '// Wait to load complete page of this URL
    While Inet1.StillExecuting
        DoEvents
    Wend
    
    '/// Retrive the Download Link meybe baypass the COUNTDOWN?
    If InStr(htmlBody, "id=""" & "downloadlink") > 0 Then
        pos1 = InStr(pos1 + 1, htmlBody, "id=""" & "downloadlink""" & "><a href=", vbTextCompare) + Len("id=""" & "downloadlink""" & "><a href=")
        pos2 = InStr(pos1 + 1, htmlBody, """", vbTextCompare) '- 1
        TempArray(0) = Mid$(htmlBody, pos1, pos2 - pos1)
        TempArray(0) = Replace(TempArray(0), "href=", "")
        TempArray(0) = Replace(TempArray(0), """", "")
    Else
        TempArray(0) = "n.a"
    End If
    
    '// Get the INFO of the File
    Call GrabInfoFile(htmlBody)
    
    If Mid$(TempArray(0), 1, 10) <> "http://www" Then
                MsgBox "Invalid URL's Download! Try again please..." & vbCrLf _
            & lblWorkStatus.Caption, vbExclamation, App.Title
        Exit Sub
    End If
    
    '// Display the Download URL
    txtDownloadURL.Text = TempArray(0)
    
    htmlBody = Empty
    
    Exit Sub
    
ErrorHandler:
    If Err.Number <> 91 Then lblWorkStatus.Caption = "Error #" & Err.Number & ". " & Err.Description
    Err.Clear
End Sub

Private Sub cmdGO_Click()
    Dim Captcha As String
    Dim htmlBody As String
    Dim pos1, pos2 As Long
    Dim TempArray(20) As String
    
    On Local Error GoTo ErrorHandler
    
    '/// Clear the Cache entry
    Call DeleteUrlCacheEntry(txtURL.Text)
    
    lblWorkStatus.Caption = "Please wait, i'm working now..."
    
    '/// Reset All
    txtDownloadURL.Text = "n.a"
    txtFileName.Text = "n.a"
    txtSize.Text = "n.a"
    txtDscription.Text = "n.a"
    txtDownloadLink.Text = "n.a"
    lblWorkStatus.Caption = "Ready..."
    
    lblWorkStatus.Caption = "Get the page URL..."
    
    '// Get the HTML Code
    htmlBody = Inet1.OpenURL(txtURL.Text)
    
    '// Wait to load complete page of this URL
    While Inet1.StillExecuting
        DoEvents
    Wend
    
    '// Save the Code of the HTML page
    '''Open App.Path + "\tmp.html" For Output As #1
    '''    Print #1, htmlBody
    '''Close #1
    
    '/// Retrive the Download Link meybe baypass the COUNTDOWN?
    If InStr(htmlBody, "id=""" & "downloadlink") > 0 Then
        pos1 = InStr(pos1 + 1, htmlBody, "id=""" & "downloadlink""" & "><a href=", vbTextCompare) + Len("id=""" & "downloadlink""" & "><a href=")
        pos2 = InStr(pos1 + 1, htmlBody, """", vbTextCompare) '- 1
        TempArray(0) = Mid$(htmlBody, pos1, pos2 - pos1)
        TempArray(0) = Replace(TempArray(0), "href=", "")
        TempArray(0) = Replace(TempArray(0), """", "")
    Else
        TempArray(0) = "n.a"
    End If
    
    '// Get the INFO of the File
    Call GrabInfoFile(htmlBody)
    
    If Mid$(TempArray(0), 1, 10) <> "http://www" Then
            MsgBox "Invalid URL's Download! Try again please..." & vbCrLf _
            & lblWorkStatus.Caption, vbExclamation, App.Title
            cmdGO.Enabled = True
        Exit Sub
    End If
    
    '// Display the Download URL
    txtDownloadURL.Text = TempArray(0)
    
    cmdGO.Enabled = True
    Exit Sub
    
    '// Download the files
    cmdDownload_Click
    
    htmlBody = Empty
    
    Exit Sub
    
ErrorHandler:
    If Err.Number <> 91 Then lblWorkStatus.Caption = "Error #" & Err.Number & ". " & Err.Description
    Err.Clear
End Sub

Private Sub cmdPaste_Click()
Dim ClpFmt
    Dim tmpText As String
    On Error Resume Next
    If Clipboard.GetFormat(vbCFText) Then ClpFmt = ClpFmt + 1
    If Clipboard.GetFormat(vbCFBitmap) Then ClpFmt = ClpFmt + 2
    If Clipboard.GetFormat(vbCFDIB) Then ClpFmt = ClpFmt + 4
    If Clipboard.GetFormat(vbCFRTF) Then ClpFmt = ClpFmt + 8
   Select Case ClpFmt
      Case 1 '/// Only TXT
            SetFocusField txtURL, True
            tmpText = Clipboard.GetText()
            If Mid$(tmpText, 1, 27) <> "http://www.megaupload.com/?" Then
                    MsgBox "Invalid URL...", vbCritical, App.Title
                Exit Sub
            End If
            txtURL.Text = Clipboard.GetText()
      Case 2, 4, 6 '/// Only PICTURE
      Case 3, 5, 7 '/// TXT and PICTURE
      Case 8, 9 '/// TXT RTF
            SetFocusField txtURL, True
            tmpText = Clipboard.GetText()
            If Mid$(tmpText, 1, 27) <> "http://www.megaupload.com/?" Then
                    MsgBox "Invalid URL...", vbCritical, App.Title
                Exit Sub
            End If
            txtURL.Text = Clipboard.GetText()
      Case Else '/// CLIPBOARD EMPTY
        MsgBox "The Clipboard is empty!", vbExclamation, App.Title
   End Select
End Sub

Private Sub cmdSave_Click()
Dim i As Integer
    Dim x As Integer
    Dim tmpString As String
    On Local Error Resume Next
    
    If lsts(3).ListCount > 0 Then
        For i = 0 To lsts(3).ListCount
            If lsts(3).List(i) <> "" And LCase(lsts(3).List(i)) = LCase(txtURL.Text) Then
                    lsts(0).Selected(i) = True
                Exit Sub
            End If
        Next i
    Else
        lsts(0).AddItem txtFileName.Text
        If txtSize.Text <> "" Then lsts(1).AddItem txtSize.Text Else lsts(1).AddItem "0 KB"
        If txtDscription.Text <> "" Then lsts(2).AddItem txtDscription.Text Else lsts(2).AddItem "No description"
        lsts(3).AddItem txtURL.Text
        GoTo AddingLine
    End If
    
    lsts(0).AddItem txtFileName.Text
    If txtSize.Text <> "" Then lsts(1).AddItem txtSize.Text Else lsts(1).AddItem "0 KB"
    If txtDscription.Text <> "" Then lsts(2).AddItem txtDscription.Text Else lsts(2).AddItem "No description"
    lsts(3).AddItem txtURL.Text
    
    GoTo AddingLine

    Exit Sub

AddingLine:
    
    tmpString = "Title*" & txtFileName.Text & "|Size*" & txtSize.Text & "|Uploaded*" _
                        & txtDscription.Text & "|Link*" & txtURL.Text
    
    If FileExists(App.Path + "\Downloads_List_Links.txt") Then _
    Open App.Path + "\Downloads_List_Links.txt" For Append As #1 _
    Else Open App.Path + "\Downloads_List_Links.txt" For Output As #1
        Print #1, tmpString
    Close #1
    
    x = lsts(0).ListCount - 1
    lsts(0).Selected(x) = True
    
Exit Sub
End Sub

Private Sub Form_Initialize()
    Call InitCommonControls
    Me.Caption = sAppTitle
    lblTop.Caption = sAppTitle
    lblInfo.Caption = "© 2008/" & Format(Now, "yyyy") & " by Salvo Cortesiano."
End Sub

Private Sub Form_Load()
    
    Dim FSO As New FileSystemObject
    If Not FSO.FolderExists(App.Path + "\Downloads") Then
        ' .... Create folder
        If MakeDirectory(App.Path + "\Downloads") = False Then:
        ' .... Until display the Error now, because if the Folder exist return a Error ;)
        Else
            str_Folder = App.Path + "\Downloads\": txtPathDownloads.Text = str_Folder
    End If
    Set FSO = Nothing
    
    Call ImportDownloadLinks
    
    If lsts(0).ListCount > 0 Then lsts(0).Selected(0) = True
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MsgBox("This will End the Program." & vbLf & vbLf & "Are you sure?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Cancel = False
        Set frmMain = Nothing
    Else
        Cancel = True
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    End
End Sub


Private Sub Label8_Click()
    Call DeleteCookiesEntry(True)
End Sub

Private Sub lsts_Click(index As Integer)
 On Error Resume Next
    If lsts(0).ListCount > 0 Then
        txtFileName.Text = lsts(0).List(lsts(0).ListIndex)
        txtSize.Text = lsts(1).List(lsts(0).ListIndex)
        txtDscription.Text = lsts(2).List(lsts(0).ListIndex)
        txtURL.Text = lsts(3).List(lsts(0).ListIndex)
        txtDownloadLink.Text = lsts(3).List(lsts(0).ListIndex)
    End If
    If txtDownloadURL.Text <> "" Then
    End If
End Sub

Private Sub lsts_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Local Error Resume Next
    If lsts(0).ListCount > 0 Then
        If lsts(0).Selected(lsts(0).ListIndex) = True Then
            If Button = 2 Then
                PopupMenu mnufile
            End If
        End If
    End If
End Sub

Private Sub lstURLs_Click()
    On Error Resume Next
    txtURL.Text = lstURLs.List(lstURLs.ListIndex)
End Sub

Private Sub lstURLs_DblClick()
    
End Sub


Private Sub mnuList_Click(index As Integer)
Dim i As Integer
    Dim x As Integer
    Dim tmpString As String
    On Local Error GoTo ErrorHandler
    Select Case index
        Case 0 '/// Remuve Selected URL
            If MsgBox("Remuve the selected URL from this List?" & vbCrLf & vbCrLf & lsts(0).List(lsts(0).ListIndex), vbYesNo + vbQuestion, App.Title) = vbYes Then
                lsts(1).RemoveItem (lsts(0).ListIndex): lsts(2).RemoveItem (lsts(0).ListIndex): lsts(3).RemoveItem (lsts(0).ListIndex): lsts(0).RemoveItem (lsts(0).ListIndex)
                If lsts(0).ListCount > 0 Then
                    x = lsts(0).ListCount - 1
                    Open App.Path + "\Downloads_List_Links.txt" For Output As #1
                    For i = 0 To lsts(0).ListCount
                        If lsts(0).List(i) <> "" Then
                            tmpString = tmpString & "Title*" & lsts(0).List(i) & "|Size*" & lsts(1).List(i) _
                            & "|Uploaded*" & lsts(2).List(i) & "|Link*" & lsts(3).List(i) & vbCrLf
                        End If
                    Next i
                        Print #1, Mid$(tmpString, 1, Len(tmpString) - Len(vbCrLf))
                        Close #1
                        lsts(0).Selected(x) = True
                    Else
                        If FileExists(App.Path + "\Downloads_List_Links.txt") Then _
                        Call Kill(App.Path + "\Downloads_List_Links.txt")
                        txtURL.Text = Empty: txtFileName.Text = "n.a": txtSize.Text = "n.a": txtDscription.Text = "n.a"
                End If
            End If
        Case 2 '/// Remuve All URL of the List
            If MsgBox("Remuve all URL of the List?", vbYesNo + vbQuestion, App.Title) = vbYes Then
                For i = 0 To 3: lsts(i).Clear: Next i
                txtURL.Text = Empty: txtFileName.Text = "n.a": txtSize.Text = "n.a": txtDscription.Text = "n.a"
                If FileExists(App.Path + "\Downloads_List_Links.txt") Then _
                Call Kill(App.Path + "\Downloads_List_Links.txt")
            End If
    End Select
Exit Sub
ErrorHandler:
        MsgBox "Error #" & Err.Number & ". " & Err.Description, vbExclamation, App.Title
    Err.Clear
End Sub

Private Sub SetFocusField(setTXT As TextBox, Optional sFocus As Boolean = True)
    On Error Resume Next
    If sFocus Then
        setTXT.SelStart = 0
        setTXT.SelLength = Len(setTXT.Text)
        setTXT.SetFocus
    Else
        setTXT.SelStart = 0
        setTXT.SetFocus
    End If
End Sub

Private Function ChkLst(iList As ListBox)
    On Local Error Resume Next
    Dim i As Integer, x As Integer
        For i = 0 To iList.ListCount - 1
            For x = 0 To iList.ListCount - 1
                If (iList.List(i) = iList.List(x)) And x <> i Then
                    lstURLs.RemoveItem (i)
                End If
            Next x
        Next i
End Function
