Attribute VB_Name = "Module2"
Option Explicit
 
' プリンタアクセス権を定義する構造体の宣言
Type PRINTER_DEFAULTS
    pDatatype     As Long
    pDevMode      As Long
    DesiredAccess As Long
End Type
 
' 標準的な権利を要求することを示す定数の宣言
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
' プリンタアクセス権の管理者権限を示す定数の宣言
Public Const PRINTER_ACCESS_ADMINISTER = &H4&
' プリンタアクセス権のユーザー権限を示す定数の宣言
Public Const PRINTER_ACCESS_USE = &H8&
' プリンタアクセス権すべての権限を示す定数の宣言
Public Const PRINTER_ALL_ACCESS = _
               (STANDARD_RIGHTS_REQUIRED Or _
                PRINTER_ACCESS_ADMINISTER Or _
                PRINTER_ACCESS_USE)
 
 
' 日付と時刻を定義する構造体の宣言
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
 
' 印刷ジョブ情報を定義する構造体の宣言
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
 
' 一時停止を示す定数の宣言
Public Const JOB_STATUS_PAUSED = &H1
' エラーを示す定数の宣言
Public Const JOB_STATUS_ERROR = &H2
' 削除中を示す定数の宣言
Public Const JOB_STATUS_DELETING = &H4
' スプールを示す定数の宣言
Public Const JOB_STATUS_SPOOLING = &H8
' 印刷中を示す定数の宣言
Public Const JOB_STATUS_PRINTING = &H10
' オフラインを示す定数の宣言
Public Const JOB_STATUS_OFFLINE = &H20
' 用紙切れを示す定数の宣言
Public Const JOB_STATUS_PAPEROUT = &H40
' 完了を示す定数の宣言
Public Const JOB_STATUS_PRINTED = &H80
' 削除を示す定数の宣言
Public Const JOB_STATUS_DELETED = &H100
' ドライバエラーを示す定数の宣言
Public Const JOB_STATUS_BLOCKED_DEVQ = &H200
' 機器不良を示す定数の宣言
Public Const JOB_STATUS_USER_INTERVENTION = &H400
' 再起動中を示す定数の宣言
Public Const JOB_STATUS_RESTART = &H800
 
' プリンタの印刷ジョブを列挙する関数の宣言
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
 
' ある位置から別の位置にメモリブロックを移動する関数の宣言
Declare Sub MoveMemory Lib "kernel32.dll" _
    Alias "RtlMoveMemory" _
   (Destination As Any, _
    Source As Any, _
    ByVal Length As Long)
 
' 文字列をバッファにコピーする関数の宣言
Declare Function lstrcpy Lib "kernel32.dll" _
    Alias "lstrcpyA" _
   (lpString1 As Any, _
    lpString2 As Any) As Long
 
 
