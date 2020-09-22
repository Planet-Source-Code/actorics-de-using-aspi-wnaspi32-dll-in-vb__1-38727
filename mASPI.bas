Attribute VB_Name = "mASPI"
Option Explicit

'*******************************************************************
'** ASPI command definitions
'*******************************************************************
Public Const SC_HA_INQUIRY = &H0        'Host adapter inquiry
Public Const SC_GET_DEV_TYPE = &H1      'Get device type
Public Const SC_EXEC_SCSI_CMD = &H2     'Execute SCSI command
Public Const SC_ABORT_SRB = &H3         'Abort an SRB
Public Const SC_RESET_DEV = &H4         'SCSI bus device reset
Public Const SC_SET_HA_PARMS = &H5      'Set HA parameters
Public Const SC_GET_DISK_INFO = &H6     'Get Disk
Public Const SC_RESCAN_SCSI_BUS = &H7   'Rebuild SCSI device map
Public Const SC_GETSET_TIMEOUTS = &H8   'Get/Set target timeouts

'*******************************************************************
'** SRB Status
'*******************************************************************
Public Const SS_PENDING = &H0           'SRB being processed
Public Const SS_COMP = &H1              'SRB completed without error
Public Const SS_ABORTED = &H2           'SRB aborted                    */
Public Const SS_ABORT_FAIL = &H3        'Unable to abort SRB
Public Const SS_ERR = &H4               'SRB completed with error
Public Const SS_INVALID_CMD = &H80      'Invalid ASPI command
Public Const SS_INVALID_HA = &H81       'Invalid host adapter number
Public Const SS_NO_DEVICE = &H82        'SCSI device not installed
Public Const SS_INVALID_SRB = &HE0      'Invalid parameter set in SRB
Public Const SS_OLD_MANAGER = &HE1      'ASPI manager doesn't support windows
Public Const SS_BUFFER_ALIGN = &HE1     'Buffer not aligned (SS_OLD_MANAGER in Win32)
Public Const SS_ILLEGAL_MODE = &HE2     'Unsupported Windows mode
Public Const SS_NO_ASPI = &HE3          'No ASPI managers
Public Const SS_FAILED_INIT = &HE4      'ASPI for windows failed init
Public Const SS_ASPI_IS_BUSY = &HE5     'No resources available to execute command
Public Const SS_BUFFER_TOO_BIG = &HE6   'Buffer size too big to handle
Public Const SS_MISMATCH_FILES = &HE7   'The DLLs/EXEs of ASPI don't version check
Public Const SS_NO_ADAPTERS = &HE8      'No host adapters located
Public Const SS_SHORT_RESOURCES = &HE9  'Couldn't allocate resources  needed to init
Public Const SS_ASPI_IS_SHUTDOWN = &HEA 'Call came to ASPI after PROCESS_DETACH
Public Const SS_BAD_INSTALL = &HEB      'The DLL or other components are installed wrong

'*******************************************************************
'** ASPI Command Packets
'*******************************************************************
'** SRB - COMMAND HEADER COMMON
Public Type SRB
    SRB_Cmd As Byte             '00h/00 ASPI command code
    SRB_Status As Byte          '01h/01 ASPI command status byte
    SRB_HaID As Byte            '02h/02 ASPI host adapter number
    SRB_Flags As Byte           '03h/03 ASPI request flags
    SRB_Hdr_Rsvd As Long        '04h/04 Reserved, must = 0
End Type

'** SRB - HOST ADAPTER INQUIRIY - SC_HA_INQUIRY (0)
Public Type SRB_HAInquiry
    SRB_Cmd As Byte             '00h/00 ASPI command code == SC_HA_INQUIRY
    SRB_Status As Byte          '01h/01 ASPI command status byte
    SRB_HaID As Byte            '02h/02 ASPI host adapter number
    SRB_Flags As Byte           '03h/03 ASPI request flags
    SRB_Hdr_Rsvd As Long        '04h/04 Reserved, must = 0
    HA_Count As Byte            '08h/08 Number of host adapters present
    HA_Id As Byte               '09h/09 SCSI ID of host adapter
    HA_MgrId As String * 16     '0ah/10 String describing the manager
    HA_Ident As String * 16     '1ah/26 String describing the host adapter
    HA_Unique(15) As Byte       '2ah/42 Host Adapter Unique parameters
    HA_Rsvd As Integer          '3ah/58 Reserved, must = 0
    HA_Pad(19) As Byte          '3eh/62 padding
End Type

'** SRB - GET DEVICE TYPE - SC_GET_DEV_TYPE (1)
Public Type SRB_GetDevType
    SRB_Cmd As Byte             '00h/00 ASPI command code == SC_HA_INQUIRY
    SRB_Status As Byte          '01h/01 ASPI command status byte
    SRB_HaID As Byte            '02h/02 ASPI host adapter number
    SRB_Flags As Byte           '03h/03 ASPI request flags
    SRB_Hdr_Rsvd As Long        '04h/04 Reserved, must = 0
    SRB_Target As Byte          '08h/08 Target's SCSI ID
    SRB_Lun As Byte             '09h/09 Target's LUN number
    DEV_DeviceType As Byte      '0ah/10 Target's peripheral device type
    DEV_Rsvd1 As Byte           '0bh/11 Reserved, must = 0
    DEV_Pad(67) As Byte         '0ch/12 padding
End Type

'** SRB - EXECUTE SCSI COMMAND - SC_EXEC_SCSI_CMD (2)
Public Type SRB_ExecuteIO
    SRB_Cmd As Byte             '00h/00 ASPI command code == SC_HA_INQUIRY
    SRB_Status As Byte          '01h/01 ASPI command status byte
    SRB_HaID As Byte            '02h/02 ASPI host adapter number
    SRB_Flags As Byte           '03h/03 ASPI request flags
    SRB_Hdr_Rsvd As Long        '04h/04 Reserved, must = 0
    SRB_Target As Byte          '08h/08 Target's SCSI ID
    SRB_Lun As Byte             '09h/09 Target's LUN number
    SRB_Rsvd1 As Integer        '0ah/10 Reserved for alignment
    SRB_BufLen As Long          '0ch/12 Data Allocation Length
    SRB_BufPointer As Long      '10h/16 Data Buffer Pointer
    SRB_SenseLen As Byte        '14h/20 Sense Allocation Length
    SRB_CDBLen As Byte          '15h/21 CDB Length
    SRB_HaStat As Byte          '16h/22 Host Adapter Status
    SRB_TargStat As Byte        '17h/23 Target Status
    SRB_PostProc As Long        '18h/24 Post routine
    SRB_Rsvd2(19) As Byte       '1ch/28 Reserved, must = 0
    SRB_CDBByte(15) As Byte     '30h/48 SCSI CDB
    SRB_SenseData(15) As Byte   '40h/64 Request Sense buffer
End Type

'*******************************************************************
'** PERIPHERAL DEVICE TYPE DEFINITIONS
'*******************************************************************
Public Const DTYPE_DASD = 0         'Disk Device
Public Const DTYPE_SEQD = 1         'Tape Device
Public Const DTYPE_PRNT = 2         'Printer
Public Const DTYPE_PROC = 3         'Processor
Public Const DTYPE_WORM = 4         'Write-once read-multiple
Public Const DTYPE_CROM = 5         'CD-ROM device
Public Const DTYPE_CDROM = 5        'CD-ROM device
Public Const DTYPE_SCAN = 6         'Scanner device
Public Const DTYPE_OPTI = 7         'Optical memory device
Public Const DTYPE_JUKE = 8         'Medium Changer device
Public Const DTYPE_COMM = 9         'Communications device
Public Const DTYPE_RESL = &HA       'Reserved (low)
Public Const DTYPE_RESH = &H1E      'Reserved (high)
Public Const DTYPE_UNKNOWN = &H1F   'Unknown or no device type

'*******************************************************************
'** Misc constants used by SCSI I/O commands
'*******************************************************************
Public Const SENSE_LEN = 14         'Default sense buffer length.
Public Const SRB_DIR_IN = &H8       'Transfer from SCSI target to host.
Public Const SRB_DIR_OUT = &H10     'Transfer from host to SCSI target.
Public Const SRB_POSTING = &H1      'Enable ASPI posting.
Public Const SRB_EVENT_NOTIFY = &H40    'Enable ASPI event notification.
Public Const SRB_ENABLE_RESIDUAL = &H4  'Enable residual byte count reporting.

'*******************************************************************
'** Host Adapter Status Values
'*******************************************************************
Public Const HASTAT_OK = &H0            'Host adapter did not detect an error.
Public Const HASTAT_TIMEOUT = &H9       'Timed out while SRB was waiting to be processed.
Public Const HASTAT_CMD_TIMEOUT = &HB   'While processing the SRB, adapter timed out.
Public Const HASTAT_MSG_REJECT = &HD    'While processing SRB, the adapter received a MESSAGE REJECT.
Public Const HASTAT_BUS_RESET = &HE     'A bus reset was detected.
Public Const HASTAT_PARITY_ERROR = &HF  'A parity error was detected.
Public Const HASTAT_REQ_SENSE_FAIL = &H10 'The adapter failed in issuing REQUEST SENSE.
Public Const HASTAT_SEL_TO = &H11       'Selection Timeout.
Public Const HASTAT_DO_DU = &H12        'Data overrun / data underrun.
Public Const HASTAT_BUS_FREE = &H13     'Unexpected bus free.
Public Const HASTAT_PHASE_ERR = &H14    'Target bus phase sequence failure.

'*******************************************************************
'** Target Status Values
'*******************************************************************
Public Const STATUS_GOOD = &H0          'Status Good.
Public Const STATUS_CHKCOND = &H2       'Check Condition.
Public Const STATUS_CONDMET = &H4       'Condition Met.
Public Const STATUS_BUSY = &H8          'Busy.
Public Const STATUS_INTERM = &H10       'Intermediate.
Public Const STATUS_INTCDMET = &H14     'Intermediate-condition met.
Public Const STATUS_RESCONF = &H18      'Reservation conflict.
Public Const STATUS_CMD_TERM = &H22     'Command Terminated.
Public Const STATUS_QFULL = &H28        'Queue full.

'*******************************************************************
'** Sense Codes
'*******************************************************************
Public Const SENSE_CURRENT = &H70       'Sense data is from current command.
Public Const SENSE_DEFFERED = &H71      'Sense data is from a previous command.

'*******************************************************************
'** Sense Key Values
'*******************************************************************
Public Const KEY_NOSENSE = &H0          'No Sense.
Public Const KEY_RECERROR = &H1         'Recovered Error.
Public Const KEY_NOTREADY = &H2         'Not Ready.
Public Const KEY_MEDIUMERROR = &H3      'Medium Error.
Public Const KEY_HARDERROR = &H4        'Hardware Error.
Public Const KEY_ILLGLREQ = &H5         'Illegal Request.
Public Const KEY_UNITATT = &H6          'Unit Attention.
Public Const KEY_DATAPROT = &H7         'Data Protection.
Public Const KEY_BLANKCHK = &H8         'Blank Check.
Public Const KEY_VENDSPEC = &H9         'Vendor Specific.
Public Const KEY_COPYABORT = &HA        'Copy Aborted.
Public Const KEY_ABORTCMD = &HB         'Aborted Command.
Public Const KEY_EQUAL = &HC            'Equal (Search).
Public Const KEY_VOLOVRFLW = &HD        'Volume Overflow.
Public Const KEY_MISCOMP = &HE          'Miscompare (Search).
Public Const KEY_RSVD = &HF             'Reserved.

'*******************************************************************
'** SCSI Commands for all Device Types
'*******************************************************************
Public Const SCSI_TST_U_RDY = &H0       'Test Unit Ready (Mandatory)
Public Const SCSI_REQ_SENSE = &H3       'Request Sense (Mandatory)
Public Const SCSI_READ = &H8            'Read (Mandatory)
Public Const SCSI_WRITE = &HA           'Write (Mandatory)
Public Const SCSI_INQUIRY = &H12        'Inquiry (Mandatory)
Public Const SCSI_MODE_SEL6 = &H15      'Mode Select 6-byte (Device Specific)
Public Const SCSI_MODE_SEN6 = &H1A      'Mode Sense 6-byte (Device Specific)
Public Const SCSI_MODE_SEL10 = &H55     'Mode Select 10-byte (Device Specific)
Public Const SCSI_MODE_SEN10 = &H5A     'Mode Sense 10-byte (Device Specific)

'*******************************************************************
'** ASPI DLL Declarations
'*******************************************************************
Public Declare Function GetASPI32SupportInfoEx Lib "ASPI32VB" _
    () As Long

Public Declare Function SendASPI32InquiryEx Lib "ASPI32VB" _
    Alias "SendASPI32CommandEx" (hSRB As SRB_HAInquiry) As Long

Public Declare Function SendASPI32DevTypeEx Lib "ASPI32VB" _
    Alias "SendASPI32CommandEx" (hSRB As SRB_GetDevType) As Long

Public Declare Function SendASPI32ExecIOEx Lib "ASPI32VB" _
    Alias "SendASPI32CommandEx" (hSRB As SRB_ExecuteIO) As Long


