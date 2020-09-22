VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   2880
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      Height          =   1452
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   1080
      Width           =   8772
   End
   Begin VB.TextBox Text4 
      Height          =   288
      Left            =   4560
      TabIndex        =   5
      Top             =   465
      Width           =   1176
   End
   Begin VB.TextBox Text3 
      Height          =   252
      Left            =   1560
      TabIndex        =   2
      Top             =   465
      Width           =   2652
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   6120
      TabIndex        =   1
      Top             =   465
      Width           =   2604
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Read TOC"
      Height          =   624
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1056
   End
   Begin VB.Label Label2 
      Caption         =   "CD-ROM"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Device Location:"
      Height          =   225
      Left            =   6120
      TabIndex        =   7
      Top             =   210
      Width           =   1425
   End
   Begin VB.Label Label4 
      Caption         =   "ASPI detected?"
      Height          =   225
      Left            =   4560
      TabIndex        =   4
      Top             =   210
      Width           =   1305
   End
   Begin VB.Label Label3 
      Caption         =   "O/S version:"
      Height          =   225
      Left            =   1590
      TabIndex        =   3
      Top             =   210
      Width           =   1410
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'(c) by Irnchen
'This example will show you how to use ASPI in your app.
'This example will look for a CD-ROM device
'Then it will read the length of the TOC in mins, secs and frms.
'The info will be displayed in the Debug-Window


'INFO: ***TOC (table of contents) is the index of a cd-rom***


Option Explicit

'some variables
Dim OpSys As Integer
Dim ASPI As Boolean

Dim CDROM_HA As Integer
Dim CDROM_ID As Integer



Private Sub Command1_Click()
    'some variables
    Dim bRet As Boolean
    Dim nRet As Long
    Dim i As Integer
    Dim j As Integer
    Dim cnt As Integer
    Dim Inquiry As SRB_HAInquiry
    Dim DevType As SRB_GetDevType
    Dim ExecIO As SRB_ExecuteIO
    Dim DataBuffer As TOC

    Dim mins As Long, secs As Long, frms As Long
    Dim sts As Long, offst As Long, s As String

    'get version and qual the access modes...
    GetOSVersion

''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ASPI is installed properly if the following is TRUE
''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'check if ASPI is available
    ASPI = AspiCheck
    If ASPI = False Then
        Debug.Print "error - no ASPI layer"
        Exit Sub
    End If

    'get adapter count
    cnt = AspiGetNumAdapters()
    If cnt = 0 Then
        Debug.Print "error - no adapters"
        Exit Sub
    End If

    'scan adapters for CDROMS
    For i = 0 To cnt
    
        'get inquiry data
        Inquiry.SRB_Cmd = SC_HA_INQUIRY
        Inquiry.SRB_HaID = i
        Inquiry.SRB_Flags = 0
        Inquiry.SRB_Hdr_Rsvd = 0

        nRet = SendASPI32InquiryEx(Inquiry)
        If (Inquiry.SRB_Status <> SS_COMP) Then GoTo skipAdapter

        'info here can be shown from inquiry...


        'scan for CDROM's
        For j = 0 To 7
            'scan dev types
            DevType.SRB_Cmd = SC_GET_DEV_TYPE
            DevType.SRB_HaID = i
            DevType.SRB_Flags = 0
            DevType.SRB_Hdr_Rsvd = 0
    
            DevType.SRB_Target = j
            DevType.SRB_Lun = 0
    
            nRet = SendASPI32DevTypeEx(DevType)
            If (DevType.SRB_Status <> SS_COMP) Then GoTo skipDevice
            'If CD-ROM was found then read the TOC
            If DevType.DEV_DeviceType = DTYPE_CDROM Then
                CDROM_HA = i
                CDROM_ID = j
                List1.AddItem "CDROM" & "  HA:" & i & "  ID:" & j & "  LU:" & "0"
                GoTo RdToc
            End If
skipDevice:
        Next j
    
skipAdapter:
    Next i


RdToc:
    '***************************************
    '*********** A S P I *******************
    '***************************************
    'got HA and ID and LU, so ** READ TOC**
    ExecIO.SRB_Cmd = SC_EXEC_SCSI_CMD
    ExecIO.SRB_HaID = CDROM_HA
    ExecIO.SRB_Flags = SRB_DIR_IN
    ExecIO.SRB_Hdr_Rsvd = 0

    ExecIO.SRB_Target = CDROM_ID
    ExecIO.SRB_Lun = 0
    ExecIO.SRB_SenseLen = 14
    'Get the length of the Buffer
    ExecIO.SRB_BufLen = &H324
    ExecIO.SRB_BufPointer = VarPtr(DataBuffer)

    ExecIO.SRB_CDBLen = &HA
    ExecIO.SRB_CDBByte(0) = &H43    'read TOC command
    ExecIO.SRB_CDBByte(1) = &H2     'MSF mode
    ExecIO.SRB_CDBByte(7) = &H3     'high-order byte of buffer len
    ExecIO.SRB_CDBByte(8) = &H24    'low-order byte of buffer len

    nRet = SendASPI32ExecIOEx(ExecIO)
    While ExecIO.SRB_Status = SS_PENDING
        DoEvents
    Wend
    'If we can't read the CD-Rom device, print an error
    If (ExecIO.SRB_Status <> SS_COMP) Then Debug.Print "error IO"

    'info on I/O
    Debug.Print "CmdStatus=" + Hex(ExecIO.SRB_Status) + "H", _
                "HaStat=" + Hex(ExecIO.SRB_HaStat) + "H", _
                "TargetStat=" + Hex(ExecIO.SRB_TargStat) + "H"

    Debug.Print "Sense Key=" + Hex(ExecIO.SRB_SenseData(2)) + "H", _
                "Sense Code=" + Hex(ExecIO.SRB_SenseData(12)) + "H"

    'gen TOC - we just want offsets
    s = ""
    For i = 0 To (DataBuffer.LastTrack - DataBuffer.FirstTrack + 1)
        mins = DataBuffer.TocTrack(i).Addr(1)
        secs = DataBuffer.TocTrack(i).Addr(2)
        frms = DataBuffer.TocTrack(i).Addr(3)
        
        offst = (mins * 60 * 75) + (secs * 75) + frms
        s = s & " " & Format$(offst)
    Next

    'Print TOC
    Debug.Print s
    
    'show TOC
    Text1.Text = "Host Adapter: " & CDROM_HA & "     Device ID: " & CDROM_ID
    'Get the TOC number
    Text2.Text = "TOC: " & s
End Sub

'This will check, if your system has APSI installed
Function AspiCheck() As Boolean
    Dim hLoad As Long
    
    'load the error messages to parse...
    hLoad = LoadLibrary("WNASPI32.DLL")

    'check for ASPI driver
    If (GetProcAddress(hLoad, "GetASPI32SupportInfo") <> 0 And _
        GetProcAddress(hLoad, "SendASPI32Command") <> 0) Then
        AspiCheck = True
    End If

    If (hLoad <> 0) Then Call FreeLibrary(hLoad)

    If (AspiCheck = True) Then
        'Yes, your System has ASPI installed
        Text4.Text = "yes"
    Else
        'No, your System has not ASPI installed
        Text4.Text = "no"
    End If

End Function

Function AspiGetNumAdapters() As Integer
    Dim nRet As Long
    Dim sts As Integer
    Dim cnt As Integer

    'query ASPI for info on transport
    nRet = GetASPI32SupportInfoEx()
    sts = (nRet / 256)
    cnt = nRet And &HF

    If (sts = SS_COMP) Then AspiGetNumAdapters = cnt
End Function

Sub GetOSVersion()
    Dim osVer As OSVERSIONINFO

    'get OS version
    osVer.dwOSVersionInfoSize = Len(osVer)
    GetVersionEx osVer
    'If your Windows is a NT...
    If (osVer.dwPlatformId = VER_PLATFORM_WIN32_NT) Then
        If (osVer.dwMajorVersion = 3 And osVer.dwMinorVersion >= 50) Then
            OpSys = OS_WINNT35
            Text3.Text = osVer.dwMajorVersion & "." & osVer.dwMinorVersion & ", OS_WINNT35"
        ElseIf (osVer.dwMajorVersion = 4) Then
            OpSys = OS_WINNT4
            Text3.Text = osVer.dwMajorVersion & "." & osVer.dwMinorVersion & ", OS_WINNT4"
        ElseIf (osVer.dwMajorVersion = 5) Then
            OpSys = OS_WIN2K
            Text3.Text = osVer.dwMajorVersion & "." & osVer.dwMinorVersion & ", OS_WIN2K"
        End If
    'If your Windows is a Win32...
    ElseIf (osVer.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS) Then
        If (osVer.dwMinorVersion = 0) Then
            OpSys = OS_WIN95
            Text3.Text = osVer.dwMajorVersion & "." & osVer.dwMinorVersion & ", OS_WIN95"
        Else
            OpSys = OS_WIN98
            Text3.Text = osVer.dwMajorVersion & "." & osVer.dwMinorVersion & ", OS_WIN98"
        End If
    Else
        OpSys = OS_UNKNOWN
    End If
End Sub



'**********************NOW*********************
'common, do some experiments with it and upload the code on PSC!
'**********************************************
'                    Irnchen
