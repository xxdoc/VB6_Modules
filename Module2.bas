Attribute VB_Name = "Module2"
Option Explicit
 
' �v�����^�A�N�Z�X�����`����\���̂̐錾
Type PRINTER_DEFAULTS
    pDatatype     As Long
    pDevMode      As Long
    DesiredAccess As Long
End Type
 
' �W���I�Ȍ�����v�����邱�Ƃ������萔�̐錾
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
' �v�����^�A�N�Z�X���̊Ǘ��Ҍ����������萔�̐錾
Public Const PRINTER_ACCESS_ADMINISTER = &H4&
' �v�����^�A�N�Z�X���̃��[�U�[�����������萔�̐錾
Public Const PRINTER_ACCESS_USE = &H8&
' �v�����^�A�N�Z�X�����ׂĂ̌����������萔�̐錾
Public Const PRINTER_ALL_ACCESS = _
               (STANDARD_RIGHTS_REQUIRED Or _
                PRINTER_ACCESS_ADMINISTER Or _
                PRINTER_ACCESS_USE)
 
 
' ���t�Ǝ������`����\���̂̐錾
Type SYSTEMTIME
    wYear         As Integer
    wMonth        As Integer
    wDayOfWeek    As Integer
    wDay          As Integer
    wHour         As Integer
    wMinute       As Integer
    wSecond       As Integer
    wMilliseconds As Integer
End Type
 
' ����W���u�����`����\���̂̐錾
Type JOB_INFO_1
    JobId        As Long
    pPrinterName As Long
    pMachineName As Long
    pUserName    As Long
    pDocument    As Long
    pDatatype    As Long
    pStatus      As Long
    Status       As Long
    Priority     As Long
    Position     As Long
    TotalPages   As Long
    PagesPrinted As Long
    Submitted    As SYSTEMTIME
End Type
 
' �ꎞ��~�������萔�̐錾
Public Const JOB_STATUS_PAUSED = &H1
' �G���[�������萔�̐錾
Public Const JOB_STATUS_ERROR = &H2
' �폜���������萔�̐錾
Public Const JOB_STATUS_DELETING = &H4
' �X�v�[���������萔�̐錾
Public Const JOB_STATUS_SPOOLING = &H8
' ������������萔�̐錾
Public Const JOB_STATUS_PRINTING = &H10
' �I�t���C���������萔�̐錾
Public Const JOB_STATUS_OFFLINE = &H20
' �p���؂�������萔�̐錾
Public Const JOB_STATUS_PAPEROUT = &H40
' �����������萔�̐錾
Public Const JOB_STATUS_PRINTED = &H80
' �폜�������萔�̐錾
Public Const JOB_STATUS_DELETED = &H100
' �h���C�o�G���[�������萔�̐錾
Public Const JOB_STATUS_BLOCKED_DEVQ = &H200
' �@��s�ǂ������萔�̐錾
Public Const JOB_STATUS_USER_INTERVENTION = &H400
' �ċN�����������萔�̐錾
Public Const JOB_STATUS_RESTART = &H800
 
' �v�����^�̈���W���u��񋓂���֐��̐錾
Declare Function EnumJobs Lib "winspool.drv" _
    Alias "EnumJobsA" _
   (ByVal hPrinter As Long, _
    ByVal FirstJob As Long, _
    ByVal NoJobs As Long, _
    ByVal Level As Long, _
    pJob As Any, _
    ByVal cdBuf As Long, _
    pcbNeeded As Long, _
    pcReturned As Long) As Long
 
' ����ʒu����ʂ̈ʒu�Ƀ������u���b�N���ړ�����֐��̐錾
Declare Sub MoveMemory Lib "kernel32.dll" _
    Alias "RtlMoveMemory" _
   (Destination As Any, _
    Source As Any, _
    ByVal Length As Long)
 
' ��������o�b�t�@�ɃR�s�[����֐��̐錾
Declare Function lstrcpy Lib "kernel32.dll" _
    Alias "lstrcpyA" _
   (lpString1 As Any, _
    lpString2 As Any) As Long
 
 
