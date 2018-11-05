Attribute VB_Name = "Module1"
Rem *///////////////////////////////////////////////////////////////////
Rem *///////////////////////////////////////////////////////////////////
Rem *///////////////////////////////////////////////////////////////////
Rem *
Rem * OPOSALL.BAS
Rem *
Rem *   Includes all OPOS Device Classes.
Rem *
Rem * Modification history
Rem * ------------------------------------------------------------------
Rem * 96-03-18 OPOS Release 1.01                                    CRM
Rem * 96-04-22 OPOS Release 1.1                                     CRM
Rem * 97-06-04 OPOS Release 1.2                                     CRM
Rem * 98-03-06 OPOS Release 1.3                                     CRM
Rem * 98-05-06 OPOS Release 1.3 Upward compatibility from 1.2   OPOS-J
Rem * 98-09-02 OPOS Release 1.4                                 OPOS-J
Rem *
Rem *///////////////////////////////////////////////////////////////////
Rem *///////////////////////////////////////////////////////////////////
Rem *///////////////////////////////////////////////////////////////////



Rem *///////////////////////////////////////////////////////////////////
Rem *
Rem * From Opos.h
Rem *
Rem *   General header file for OPOS Applications.
Rem *
Rem * Modification history
Rem * ------------------------------------------------------------------
Rem * 95-12-08 OPOS Release 1.0                                     CRM
Rem * 97-06-04 OPOS Release 1.2                                     CRM
Rem *   Add OPOS_FOREVER.
Rem *   Add BinaryConversion values.
Rem * 98-03-06 OPOS Release 1.3                                     CRM
Rem *   Add CapPowerReporting, PowerState, and PowerNotify values.
Rem *   Add power reporting values for StatusUpdateEvent.
Rem *
Rem *///////////////////////////////////////////////////////////////////


Rem *///////////////////////////////////////////////////////////////////
Rem * OPOS "State" Property Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const OPOS_S_CLOSED& = 1
Public Const OPOS_S_IDLE& = 2
Public Const OPOS_S_BUSY& = 3
Public Const OPOS_S_ERROR& = 4


Rem *///////////////////////////////////////////////////////////////////
Rem * OPOS "ResultCode" Property Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const OPOSERR& = 100
Public Const OPOSERREXT& = 200

Public Const OPOS_SUCCESS& = 0
Public Const OPOS_E_CLOSED& = 1 + OPOSERR
Public Const OPOS_E_CLAIMED& = 2 + OPOSERR
Public Const OPOS_E_NOTCLAIMED& = 3 + OPOSERR
Public Const OPOS_E_NOSERVICE& = 4 + OPOSERR
Public Const OPOS_E_DISABLED& = 5 + OPOSERR
Public Const OPOS_E_ILLEGAL& = 6 + OPOSERR
Public Const OPOS_E_NOHARDWARE& = 7 + OPOSERR
Public Const OPOS_E_OFFLINE& = 8 + OPOSERR
Public Const OPOS_E_NOEXIST& = 9 + OPOSERR
Public Const OPOS_E_EXISTS& = 10 + OPOSERR
Public Const OPOS_E_FAILURE& = 11 + OPOSERR
Public Const OPOS_E_TIMEOUT& = 12 + OPOSERR
Public Const OPOS_E_BUSY& = 13 + OPOSERR
Public Const OPOS_E_EXTENDED& = 14 + OPOSERR


Rem *///////////////////////////////////////////////////////////////////
Rem * OPOS "BinaryConversion" Property Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const OPOS_BC_NONE& = 0
Public Const OPOS_BC_NIBBLE& = 1
Public Const OPOS_BC_DECIMAL& = 2


Rem *///////////////////////////////////////////////////////////////////
Rem * "CheckHealth" Method: "Level" Parameter Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const OPOS_CH_INTERNAL& = 1
Public Const OPOS_CH_EXTERNAL& = 2
Public Const OPOS_CH_INTERACTIVE& = 3


Rem *///////////////////////////////////////////////////////////////////
Rem * OPOS "CapPowerReporting", "PowerState", "PowerNotify" Property
Rem *   Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const OPOS_PR_NONE& = 0
Public Const OPOS_PR_STANDARD& = 1
Public Const OPOS_PR_ADVANCED& = 2

Public Const OPOS_PN_DISABLED& = 0
Public Const OPOS_PN_ENABLED& = 1

Public Const OPOS_PS_UNKNOWN& = 2000
Public Const OPOS_PS_ONLINE& = 2001
Public Const OPOS_PS_OFF& = 2002
Public Const OPOS_PS_OFFLINE& = 2003
Public Const OPOS_PS_OFF_OFFLINE& = 2004


Rem *///////////////////////////////////////////////////////////////////
Rem * "ErrorEvent" Event: "ErrorLocus" Parameter Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const OPOS_EL_OUTPUT& = 1
Public Const OPOS_EL_INPUT& = 2
Public Const OPOS_EL_INPUT_DATA& = 3


Rem *///////////////////////////////////////////////////////////////////
Rem * "ErrorEvent" Event: "ErrorResponse" Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const OPOS_ER_RETRY& = 11
Public Const OPOS_ER_CLEAR& = 12
Public Const OPOS_ER_CONTINUEINPUT& = 13


Rem *///////////////////////////////////////////////////////////////////
Rem * "StatusUpdateEvent" Event: Common "Status" Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const OPOS_SUE_POWER_ONLINE& = 2001
Public Const OPOS_SUE_POWER_OFF& = 2002
Public Const OPOS_SUE_POWER_OFFLINE& = 2003
Public Const OPOS_SUE_POWER_OFF_OFFLINE& = 2004


Rem *///////////////////////////////////////////////////////////////////
Rem * General Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const OPOS_FOREVER& = -1



Rem *///////////////////////////////////////////////////////////////////
Rem *
Rem * From OposBb.h
Rem *
Rem *   Bump Bar header file for OPOS Applications.
Rem *
Rem * Modification history
Rem * ------------------------------------------------------------------
Rem * 98-03-06 OPOS Release 1.3                                     BB
Rem *
Rem *///////////////////////////////////////////////////////////////////


Rem *///////////////////////////////////////////////////////////////////
Rem * "CurrentUnitID" and "UnitsOnline" Properties
Rem *  and "Units" Parameter Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const BB_UID_1& = &H1
Public Const BB_UID_2& = &H2
Public Const BB_UID_3& = &H4
Public Const BB_UID_4& = &H8
Public Const BB_UID_5& = &H10
Public Const BB_UID_6& = &H20
Public Const BB_UID_7& = &H40
Public Const BB_UID_8& = &H80
Public Const BB_UID_9& = &H100
Public Const BB_UID_10& = &H200
Public Const BB_UID_11& = &H400
Public Const BB_UID_12& = &H800
Public Const BB_UID_13& = &H1000
Public Const BB_UID_14& = &H2000
Public Const BB_UID_15& = &H4000
Public Const BB_UID_16& = &H8000
Public Const BB_UID_17& = &H10000
Public Const BB_UID_18& = &H20000
Public Const BB_UID_19& = &H40000
Public Const BB_UID_20& = &H80000
Public Const BB_UID_21& = &H100000
Public Const BB_UID_22& = &H200000
Public Const BB_UID_23& = &H400000
Public Const BB_UID_24& = &H800000
Public Const BB_UID_25& = &H1000000
Public Const BB_UID_26& = &H2000000
Public Const BB_UID_27& = &H4000000
Public Const BB_UID_28& = &H8000000
Public Const BB_UID_29& = &H10000000
Public Const BB_UID_30& = &H20000000
Public Const BB_UID_31& = &H40000000
Public Const BB_UID_32& = &H80000000


Rem *///////////////////////////////////////////////////////////////////
Rem * "DataEvent" Event: "Status" Parameter Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const BB_DE_KEY& = &H1



Rem *///////////////////////////////////////////////////////////////////
Rem *
Rem * From OposCash.h
Rem *
Rem *   Cash Drawer header file for OPOS Applications.
Rem *
Rem * Modification history
Rem * ------------------------------------------------------------------
Rem * 95-12-08 OPOS Release 1.0                                     CRM
Rem * 98-03-06 OPOS Release 1.3                                     CRM
Rem *
Rem *///////////////////////////////////////////////////////////////////


Rem *///////////////////////////////////////////////////////////////////
Rem * "StatusUpdateEvent" Event Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const CASH_SUE_DRAWERCLOSED& = 0
Public Const CASH_SUE_DRAWEROPEN& = 1



Rem *///////////////////////////////////////////////////////////////////
Rem *
Rem * From OposCat.h
Rem *
Rem *   CAT header file for OPOS Applications.
Rem *
Rem * Modification history
Rem *------------------------------------------------------------------
Rem * 98-06-01 OPOS Release 1.4                                  OPOS-J
Rem *
Rem *///////////////////////////////////////////////////////////////////


Rem *///////////////////////////////////////////////////////////////////
Rem * Payment Condition Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const CAT_PAYMENT_LUMP& = 10
Public Const CAT_PAYMENT_BONUS_1& = 21
Public Const CAT_PAYMENT_BONUS_2& = 22
Public Const CAT_PAYMENT_BONUS_3& = 23
Public Const CAT_PAYMENT_BONUS_4& = 24
Public Const CAT_PAYMENT_BONUS_5& = 25
Public Const CAT_PAYMENT_INSTALLMENT_1& = 61
Public Const CAT_PAYMENT_INSTALLMENT_2& = 62
Public Const CAT_PAYMENT_INSTALLMENT_3& = 63
Public Const CAT_PAYMENT_BONUS_COMBINATION_1& = 31
Public Const CAT_PAYMENT_BONUS_COMBINATION_2& = 32
Public Const CAT_PAYMENT_BONUS_COMBINATION_3& = 33
Public Const CAT_PAYMENT_BONUS_COMBINATION_4& = 34
Public Const CAT_PAYMENT_REVOLVING& = 80


Rem *///////////////////////////////////////////////////////////////////
Rem * Transaction Type Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const CAT_TRANSACTION_SALES& = 10
Public Const CAT_TRANSACTION_VOID& = 20
Public Const CAT_TRANSACTION_REFUND& = 21
Public Const CAT_TRANSACTION_VOIDPRESALES& = 29
Public Const CAT_TRANSACTION_COMPLETION& = 30
Public Const CAT_TRANSACTION_PRESALES& = 40
Public Const CAT_TRANSACTION_CHECKCARD& = 41


Rem *///////////////////////////////////////////////////////////////////
Rem * ResultCodeExtended Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const OPOS_ECAT_CENTERERROR& = 1
Public Const OPOS_ECAT_COMMANDERROR& = 90
Public Const OPOS_ECAT_RESET& = 91
Public Const OPOS_ECAT_COMMUNICATIONERROR& = 92
Public Const OPOS_ECAT_DAILYLOGOVERFLOW& = 200


Rem *///////////////////////////////////////////////////////////////////
Rem * "Daily Log" Property  & Argument Constants
Rem *///////////////////////////////////////////////////////////////////
Public Const CAT_DL_NONE& = 0                           ' None of them
Public Const CAT_DL_REPORTING& = 1                      ' Only Reporting
Public Const CAT_DL_SETTLEMENT& = 2                     ' Only Settlement
Public Const CAT_DL_REPORTING_SETTLEMENT& = 3           ' Both of them


Rem *///////////////////////////////////////////////////////////////////
Rem *
Rem * From OposChan.h
Rem *
Rem *   Cash Changer header file for OPOS Applications.
Rem *
Rem * Modification history
Rem * ------------------------------------------------------------------
Rem * 97-06-04 OPOS Release 1.2                                     CRM
Rem *
Rem *///////////////////////////////////////////////////////////////////


Rem *///////////////////////////////////////////////////////////////////
Rem * "DeviceStatus" and "FullStatus" Property Constants
Rem * "StatusUpdateEvent" Event Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const CHAN_STATUS_OK& = 0                       ' DeviceStatus, FullStatus

Public Const CHAN_STATUS_EMPTY& = 11                   ' DeviceStatus, StatusUpdateEvent
Public Const CHAN_STATUS_NEAREMPTY& = 12               ' DeviceStatus, StatusUpdateEvent
Public Const CHAN_STATUS_EMPTYOK& = 13                 ' StatusUpdateEvent

Public Const CHAN_STATUS_FULL& = 21                    ' FullStatus, StatusUpdateEvent
Public Const CHAN_STATUS_NEARFULL& = 22                ' FullStatus, StatusUpdateEvent
Public Const CHAN_STATUS_FULLOK& = 23                  ' StatusUpdateEvent

Public Const CHAN_STATUS_JAM& = 31                     ' DeviceStatus, StatusUpdateEvent
Public Const CHAN_STATUS_JAMOK& = 32                   ' StatusUpdateEvent

Public Const CHAN_STATUS_ASYNC& = 91                   ' StatusUpdateEvent


Rem *///////////////////////////////////////////////////////////////////
Rem * "ResultCodeExtended" Property Constants for Cash Changer
Rem *///////////////////////////////////////////////////////////////////


Public Const OPOS_ECHAN_OVERDISPENSE& = 1 + OPOSERREXT



Rem *///////////////////////////////////////////////////////////////////
Rem *
Rem * From OposCoin.h
Rem *
Rem *   Coin Dispenser header file for OPOS Applications.
Rem *
Rem * Modification history
Rem * ------------------------------------------------------------------
Rem * 95-12-08 OPOS Release 1.0                                     CRM
Rem *
Rem *///////////////////////////////////////////////////////////////////


Rem *///////////////////////////////////////////////////////////////////
Rem * "DispenserStatus" Property Constants
Rem * "StatusUpdateEvent" Event: "Data" Parameter Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const COIN_STATUS_OK& = 1
Public Const COIN_STATUS_EMPTY& = 2
Public Const COIN_STATUS_NEAREMPTY& = 3
Public Const COIN_STATUS_JAM& = 4



Rem *///////////////////////////////////////////////////////////////////
Rem *
Rem * From OposDisp.h
Rem *
Rem *   Line Display header file for OPOS Applications.
Rem *
Rem * Modification history
Rem * ------------------------------------------------------------------
Rem * 95-12-08 OPOS Release 1.0                                     CRM
Rem * 96-03-18 OPOS Release 1.01                                    CRM
Rem *   Add DISP_MT_INIT constant and MarqueeFormat constants.
Rem * 96-04-22 OPOS Release 1.1                                     CRM
Rem *   Add CapCharacterSet values for Kana and Kanji.
Rem *
Rem *///////////////////////////////////////////////////////////////////


Rem *///////////////////////////////////////////////////////////////////
Rem * "CapBlink" Property Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const DISP_CB_NOBLINK& = 0
Public Const DISP_CB_BLINKALL& = 1
Public Const DISP_CB_BLINKEACH& = 2


Rem *///////////////////////////////////////////////////////////////////
Rem * "CapCharacterSet" Property Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const DISP_CCS_NUMERIC& = 0
Public Const DISP_CCS_ALPHA& = 1
Public Const DISP_CCS_ASCII& = 998
Public Const DISP_CCS_KANA& = 10
Public Const DISP_CCS_KANJI& = 11


Rem *///////////////////////////////////////////////////////////////////
Rem * "CharacterSet" Property Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const DISP_CS_ASCII& = 998
Public Const DISP_CS_WINDOWS& = 999


Rem *///////////////////////////////////////////////////////////////////
Rem * "MarqueeType" Property Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const DISP_MT_NONE& = 0
Public Const DISP_MT_UP& = 1
Public Const DISP_MT_DOWN& = 2
Public Const DISP_MT_LEFT& = 3
Public Const DISP_MT_RIGHT& = 4
Public Const DISP_MT_INIT& = 5


Rem *///////////////////////////////////////////////////////////////////
Rem * "MarqueeFormat" Property Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const DISP_MF_WALK& = 0
Public Const DISP_MF_PLACE& = 1


Rem *///////////////////////////////////////////////////////////////////
Rem * "DisplayText" Method: "Attribute" Property Constants
Rem * "DisplayTextAt" Method: "Attribute" Property Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const DISP_DT_NORMAL& = 0
Public Const DISP_DT_BLINK& = 1


Rem *///////////////////////////////////////////////////////////////////
Rem * "ScrollText" Method: "Direction" Parameter Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const DISP_ST_UP& = 1
Public Const DISP_ST_DOWN& = 2
Public Const DISP_ST_LEFT& = 3
Public Const DISP_ST_RIGHT& = 4


Rem *///////////////////////////////////////////////////////////////////
Rem * "SetDescriptor" Method: "Attribute" Parameter Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const DISP_SD_OFF& = 0
Public Const DISP_SD_ON& = 1
Public Const DISP_SD_BLINK& = 2



Rem *///////////////////////////////////////////////////////////////////
Rem *
Rem * From OposFptr.h
Rem *
Rem *   Fiscal Printer header file for OPOS Applications.
Rem *
Rem * Modification history
Rem * ------------------------------------------------------------------
Rem * 98-03-06 OPOS Release 1.3                                     PDU
Rem *
Rem *///////////////////////////////////////////////////////////////////


Rem *///////////////////////////////////////////////////////////////////
Rem * Fiscal Printer Station Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const FPTR_S_JOURNAL& = 1
Public Const FPTR_S_RECEIPT& = 2
Public Const FPTR_S_SLIP& = 4

Public Const FPTR_S_JOURNAL_RECEIPT& = FPTR_S_JOURNAL Or FPTR_S_RECEIPT


Rem *///////////////////////////////////////////////////////////////////
Rem * "CountryCode" Property Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const FPTR_CC_BRAZIL& = 1
Public Const FPTR_CC_GREECE& = 2
Public Const FPTR_CC_HUNGARY& = 3
Public Const FPTR_CC_ITALY& = 4
Public Const FPTR_CC_POLAND& = 5
Public Const FPTR_CC_TURKEY& = 6


Rem *///////////////////////////////////////////////////////////////////
Rem * "ErrorLevel" Property Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const FPTR_EL_NONE& = 1
Public Const FPTR_EL_RECOVERABLE& = 2
Public Const FPTR_EL_FATAL& = 3
Public Const FPTR_EL_BLOCKED& = 4


Rem *///////////////////////////////////////////////////////////////////
Rem * "ErrorState", "PrinterState" Property Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const FPTR_PS_MONITOR& = 1
Public Const FPTR_PS_FISCAL_RECEIPT& = 2
Public Const FPTR_PS_FISCAL_RECEIPT_TOTAL& = 3
Public Const FPTR_PS_FISCAL_RECEIPT_ENDING& = 4
Public Const FPTR_PS_FISCAL_DOCUMENT& = 5
Public Const FPTR_PS_FIXED_OUTPUT& = 6
Public Const FPTR_PS_ITEM_LIST& = 7
Public Const FPTR_PS_LOCKED& = 8
Public Const FPTR_PS_NONFISCAL& = 9
Public Const FPTR_PS_REPORT& = 10


Rem *///////////////////////////////////////////////////////////////////
Rem * "SlipSelection" Property Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const FPTR_SS_FULL_LENGTH& = 1
Public Const FPTR_SS_VALIDATION& = 2


Rem *///////////////////////////////////////////////////////////////////
Rem * "GetData" Method Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const FPTR_GD_CURRENT_TOTAL& = 1
Public Const FPTR_GD_DAILY_TOTAL& = 2
Public Const FPTR_GD_RECEIPT_NUMBER& = 3
Public Const FPTR_GD_REFUND& = 4
Public Const FPTR_GD_NOT_PAID& = 5
Public Const FPTR_GD_MID_VOID& = 6
Public Const FPTR_GD_Z_REPORT& = 7
Public Const FPTR_GD_GRAND_TOTAL& = 8
Public Const FPTR_GD_PRINTER_ID& = 9
Public Const FPTR_GD_FIRMWARE& = 10
Public Const FPTR_GD_RESTART& = 11


Rem *///////////////////////////////////////////////////////////////////
Rem * "AdjustmentType" arguments in diverse methods
Rem *///////////////////////////////////////////////////////////////////

Public Const FPTR_AT_AMOUNT_DISCOUNT& = 1
Public Const FPTR_AT_AMOUNT_SURCHARGE& = 2
Public Const FPTR_AT_PERCENTAGE_DISCOUNT& = 3
Public Const FPTR_AT_PERCENTAGE_SURCHARGE& = 4


Rem *///////////////////////////////////////////////////////////////////
Rem * "ReportType" argument in "PrintReport" method
Rem *///////////////////////////////////////////////////////////////////

Public Const FPTR_RT_ORDINAL& = 1
Public Const FPTR_RT_DATE& = 2


Rem *///////////////////////////////////////////////////////////////////
Rem * "StatusUpdateEvent" Event: "Data" Parameter Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const FPTR_SUE_COVER_OPEN& = 11
Public Const FPTR_SUE_COVER_OK& = 12

Public Const FPTR_SUE_JRN_EMPTY& = 21
Public Const FPTR_SUE_JRN_NEAREMPTY& = 22
Public Const FPTR_SUE_JRN_PAPEROK& = 23

Public Const FPTR_SUE_REC_EMPTY& = 24
Public Const FPTR_SUE_REC_NEAREMPTY& = 25
Public Const FPTR_SUE_REC_PAPEROK& = 26

Public Const FPTR_SUE_SLP_EMPTY& = 27
Public Const FPTR_SUE_SLP_NEAREMPTY& = 28
Public Const FPTR_SUE_SLP_PAPEROK& = 29

Public Const FPTR_SUE_IDLE& = 1001


Rem *///////////////////////////////////////////////////////////////////
Rem * "ResultCodeExtended" Property Constants for Fiscal Printer
Rem *///////////////////////////////////////////////////////////////////

Public Const OPOS_EFPTR_COVER_OPEN& = 1 + OPOSERREXT     ' (Several)
Public Const OPOS_EFPTR_JRN_EMPTY& = 2 + OPOSERREXT      ' (Several)
Public Const OPOS_EFPTR_REC_EMPTY& = 3 + OPOSERREXT      ' (Several)
Public Const OPOS_EFPTR_SLP_EMPTY& = 4 + OPOSERREXT      ' (Several)
Public Const OPOS_EFPTR_SLP_FORM& = 5 + OPOSERREXT       ' EndRemoval
Public Const OPOS_EFPTR_MISSING_DEVICES& = 6 + OPOSERREXT  ' (Several)
Public Const OPOS_EFPTR_WRONG_STATE& = 7 + OPOSERREXT    ' (Several)
Public Const OPOS_EFPTR_TECHNICAL_ASSISTANCE& = 8 + OPOSERREXT  ' (Several)
Public Const OPOS_EFPTR_CLOCK_ERROR& = 9 + OPOSERREXT    ' (Several)
Public Const OPOS_EFPTR_FISCAL_MEMORY_FULL& = 10 + OPOSERREXT ' (Several)
Public Const OPOS_EFPTR_FISCAL_MEMORY_DISCONNECTED& = 11 + OPOSERREXT ' (Several)
Public Const OPOS_EFPTR_FISCAL_TOTALS_ERROR& = 12 + OPOSERREXT ' (Several)
Public Const OPOS_EFPTR_BAD_ITEM_QUANTITY& = 13 + OPOSERREXT ' (Several)
Public Const OPOS_EFPTR_BAD_ITEM_AMOUNT& = 14 + OPOSERREXT ' (Several)
Public Const OPOS_EFPTR_BAD_ITEM_DESCRIPTION& = 15 + OPOSERREXT ' (Several)
Public Const OPOS_EFPTR_RECEIPT_TOTAL_OVERFLOW& = 16 + OPOSERREXT ' (Several)
Public Const OPOS_EFPTR_BAD_VAT& = 17 + OPOSERREXT       ' (Several)
Public Const OPOS_EFPTR_BAD_PRICE& = 18 + OPOSERREXT     ' (Several)
Public Const OPOS_EFPTR_BAD_DATE& = 19 + OPOSERREXT      ' (Several)
Public Const OPOS_EFPTR_NEGATIVE_TOTAL& = 20 + OPOSERREXT ' (Several)
Public Const OPOS_EFPTR_WORD_NOT_ALLOWED& = 21 + OPOSERREXT ' (Several)



Rem *///////////////////////////////////////////////////////////////////
Rem *
Rem * From OposKbd.h
Rem *
Rem *   POS Keyboard header file for OPOS Applications.
Rem *
Rem * Modification history
Rem * ------------------------------------------------------------------
Rem * 96-04-22 OPOS Release 1.1                                     CRM
Rem * 97-06-04 OPOS Release 1.2                                     CRM
Rem *   Add "EventTypes" and "POSKeyEventType" values.
Rem *
Rem *///////////////////////////////////////////////////////////////////


Rem *///////////////////////////////////////////////////////////////////
Rem * "EventTypes" Property Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const KBD_ET_DOWN& = 1
Public Const KBD_ET_DOWN_UP& = 2


Rem *///////////////////////////////////////////////////////////////////
Rem * "POSKeyEventType" Property Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const KBD_KET_KEYDOWN& = 1
Public Const KBD_KET_KEYUP& = 2



Rem *///////////////////////////////////////////////////////////////////
Rem *
Rem * From OposLock.h
Rem *
Rem *   Keylock header file for OPOS Applications.
Rem *
Rem * Modification history
Rem * ------------------------------------------------------------------
Rem * 95-12-08 OPOS Release 1.0                                     CRM
Rem *
Rem *///////////////////////////////////////////////////////////////////


Rem *///////////////////////////////////////////////////////////////////
Rem * "KeyPosition" Property Constants
Rem * "WaitForKeylockChange" Method: "KeyPosition" Parameter
Rem * "StatusUpdateEvent" Event: "Data" Parameter
Rem *///////////////////////////////////////////////////////////////////

Public Const LOCK_KP_ANY& = 0                          ' WaitForKeylockChange Only
Public Const LOCK_KP_LOCK& = 1
Public Const LOCK_KP_NORM& = 2
Public Const LOCK_KP_SUPR& = 3



Rem *///////////////////////////////////////////////////////////////////
Rem *
Rem * From OposMicr.h
Rem *
Rem *   MICR header file for OPOS Applications.
Rem *
Rem * Modification history
Rem * ------------------------------------------------------------------
Rem * 95-12-08 OPOS Release 1.0                                     CRM
Rem *
Rem *///////////////////////////////////////////////////////////////////


Rem *///////////////////////////////////////////////////////////////////
Rem * "CheckType" Property Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const MICR_CT_PERSONAL& = 1
Public Const MICR_CT_BUSINESS& = 2
Public Const MICR_CT_UNKNOWN& = 99


Rem *///////////////////////////////////////////////////////////////////
Rem * "CountryCode" Property Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const MICR_CC_USA& = 1
Public Const MICR_CC_CANADA& = 2
Public Const MICR_CC_MEXICO& = 3
Public Const MICR_CC_UNKNOWN& = 99


Rem *///////////////////////////////////////////////////////////////////
Rem * "ResultCodeExtended" Property Constants for MICR
Rem *///////////////////////////////////////////////////////////////////

Public Const OPOS_EMICR_NOCHECK& = 1 + OPOSERREXT       ' EndInsertion
Public Const OPOS_EMICR_CHECK& = 2 + OPOSERREXT         ' EndRemoval



Rem *///////////////////////////////////////////////////////////////////
Rem *
Rem * From OposMsr.h
Rem *
Rem *   Magnetic Stripe Reader header file for OPOS Applications.
Rem *
Rem * Modification history
Rem * ------------------------------------------------------------------
Rem * 95-12-08 OPOS Release 1.0                                     CRM
Rem * 97-06-04 OPOS Release 1.2                                     CRM
Rem *   Add ErrorReportingType values.
Rem *
Rem *///////////////////////////////////////////////////////////////////


Rem *///////////////////////////////////////////////////////////////////
Rem * "TracksToRead" Property Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const MSR_TR_1& = 1
Public Const MSR_TR_2& = 2
Public Const MSR_TR_3& = 4

Public Const MSR_TR_1_2& = MSR_TR_1 Or MSR_TR_2
Public Const MSR_TR_1_3& = MSR_TR_1 Or MSR_TR_3
Public Const MSR_TR_2_3& = MSR_TR_2 Or MSR_TR_3

Public Const MSR_TR_1_2_3& = MSR_TR_1 Or MSR_TR_2 Or MSR_TR_3


Rem *///////////////////////////////////////////////////////////////////
Rem * "ErrorReportingType" Property Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const MSR_ERT_CARD& = 0
Public Const MSR_ERT_TRACK& = 1


Rem *///////////////////////////////////////////////////////////////////
Rem * "ErrorEvent" Event: "ResultCodeExtended" Parameter Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const OPOS_EMSR_START& = 1 + OPOSERREXT
Public Const OPOS_EMSR_END& = 2 + OPOSERREXT
Public Const OPOS_EMSR_PARITY& = 3 + OPOSERREXT
Public Const OPOS_EMSR_LRC& = 4 + OPOSERREXT



Rem *///////////////////////////////////////////////////////////////////
Rem *
Rem * From OposPpad.h
Rem *
Rem *   PIN Pad header file for OPOS Applications.
Rem *
Rem * Modification history
Rem * ------------------------------------------------------------------
Rem * 98-03-06 OPOS Release 1.3                                     JDB
Rem *
Rem *///////////////////////////////////////////////////////////////////


Rem *///////////////////////////////////////////////////////////////////
Rem * "CapDisplay" Property Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const PPAD_DISP_UNRESTRICTED& = 1
Public Const PPAD_DISP_PINRESTRICTED& = 2
Public Const PPAD_DISP_RESTRICTEDLIST& = 3
Public Const PPAD_DISP_RESTRICTEDORDER& = 4


Rem *///////////////////////////////////////////////////////////////////
Rem * "AvailablePromptsList" and "Prompt" Property Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const PPAD_MSG_ENTERPIN& = 1
Public Const PPAD_MSG_PLEASEWAIT& = 2
Public Const PPAD_MSG_ENTERVALIDPIN& = 3
Public Const PPAD_MSG_RETRIESEXCEEDED& = 4
Public Const PPAD_MSG_APPROVED& = 5
Public Const PPAD_MSG_DECLINED& = 6
Public Const PPAD_MSG_CANCELED& = 7
Public Const PPAD_MSG_AMOUNTOK& = 8
Public Const PPAD_MSG_NOTREADY& = 9
Public Const PPAD_MSG_IDLE& = 10
Public Const PPAD_MSG_SLIDE_CARD& = 11
Public Const PPAD_MSG_INSERTCARD& = 12
Public Const PPAD_MSG_SELECTCARDTYPE& = 13


Rem *///////////////////////////////////////////////////////////////////
Rem * "CapLanguage" Property Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const PPAD_LANG_NONE& = 1
Public Const PPAD_LANG_ONE& = 2
Public Const PPAD_LANG_PINRESTRICTED& = 3
Public Const PPAD_LANG_UNRESTRICTED& = 4

Rem *///////////////////////////////////////////////////////////////////
Rem * "TransactionType" Property Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const PPAD_TRANS_DEBIT& = 1
Public Const PPAD_TRANS_CREDIT& = 2
Public Const PPAD_TRANS_INQ& = 3
Public Const PPAD_TRANS_RECONCILE& = 4
Public Const PPAD_TRANS_ADMIN& = 5


Rem *///////////////////////////////////////////////////////////////////
Rem * "EndEFTTransaction" Method Completion Code Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const PPAD_EFT_NORMAL& = 1
Public Const PPAD_EFT_ABNORMAL& = 2


Rem *///////////////////////////////////////////////////////////////////
Rem * "DataEvent" Event Status Constants
Rem *///////////////////////////////////////////////////////////////////
Public Const PPAD_SUCCESS& = 1
Public Const PPAD_CANCEL& = 2



Rem *///////////////////////////////////////////////////////////////////
Rem *
Rem * From OposPtr.h
Rem *
Rem *   POS Printer header file for OPOS Applications.
Rem *
Rem * Modification history
Rem * ------------------------------------------------------------------
Rem * 95-12-08 OPOS Release 1.0                                     CRM
Rem * 96-04-22 OPOS Release 1.1                                     CRM
Rem *   Add CapCharacterSet values.
Rem *   Add ErrorLevel values.
Rem *   Add TransactionPrint Control values.
Rem * 97-06-04 OPOS Release 1.2                                     CRM
Rem *   Remove PTR_RP_NORMAL_ASYNC.
Rem *   Add more barcode symbologies.
Rem * 98-03-06 OPOS Release 1.3                                     CRM
Rem *   Add more PrintTwoNormal constants.
Rem *///////////////////////////////////////////////////////////////////


Rem *///////////////////////////////////////////////////////////////////
Rem * Printer Station Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const PTR_S_JOURNAL& = 1
Public Const PTR_S_RECEIPT& = 2
Public Const PTR_S_SLIP& = 4

Public Const PTR_S_JOURNAL_RECEIPT& = PTR_S_JOURNAL Or PTR_S_RECEIPT
Public Const PTR_S_JOURNAL_SLIP& = PTR_S_JOURNAL Or PTR_S_SLIP
Public Const PTR_S_RECEIPT_SLIP& = PTR_S_RECEIPT Or PTR_S_SLIP

Public Const PTR_TWO_RECEIPT_JOURNAL& = &H8000 + PTR_S_JOURNAL_RECEIPT
Public Const PTR_TWO_SLIP_JOURNAL& = &H8000 + PTR_S_JOURNAL_SLIP
Public Const PTR_TWO_SLIP_RECEIPT& = &H8000 + PTR_S_RECEIPT_SLIP


Rem *///////////////////////////////////////////////////////////////////
Rem * "CapCharacterSet" Property Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const PTR_CCS_ALPHA& = 1
Public Const PTR_CCS_ASCII& = 998
Public Const PTR_CCS_KANA& = 10
Public Const PTR_CCS_KANJI& = 11


Rem *///////////////////////////////////////////////////////////////////
Rem * "CharacterSet" Property Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const PTR_CS_ASCII& = 998
Public Const PTR_CS_WINDOWS& = 999


Rem *///////////////////////////////////////////////////////////////////
Rem * "ErrorLevel" Property Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const PTR_EL_NONE& = 1
Public Const PTR_EL_RECOVERABLE& = 2
Public Const PTR_EL_FATAL& = 3


Rem *///////////////////////////////////////////////////////////////////
Rem * "MapMode" Property Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const PTR_MM_DOTS& = 1
Public Const PTR_MM_TWIPS& = 2
Public Const PTR_MM_ENGLISH& = 3
Public Const PTR_MM_METRIC& = 4


Rem *///////////////////////////////////////////////////////////////////
Rem * "CutPaper" Method Constant
Rem *///////////////////////////////////////////////////////////////////

Public Const PTR_CP_FULLCUT& = 100


Rem *///////////////////////////////////////////////////////////////////
Rem * "PrintBarCode" Method Constants:
Rem *///////////////////////////////////////////////////////////////////

Rem *   "Alignment" Parameter
Rem *     Either the distance from the left-most print column to the start
Rem *     of the bar code, or one of the following:

Public Const PTR_BC_LEFT& = -1
Public Const PTR_BC_CENTER& = -2
Public Const PTR_BC_RIGHT& = -3

Rem *   "TextPosition" Parameter

Public Const PTR_BC_TEXT_NONE& = -11
Public Const PTR_BC_TEXT_ABOVE& = -12
Public Const PTR_BC_TEXT_BELOW& = -13

Rem *   "Symbology" Parameter:

Rem *     One dimensional symbologies
Public Const PTR_BCS_UPCA& = 101                       ' Digits
Public Const PTR_BCS_UPCE& = 102                       ' Digits
Public Const PTR_BCS_JAN8& = 103                       ' = EAN 8
Public Const PTR_BCS_EAN8& = 103                       ' = JAN 8 (added in 1.2)
Public Const PTR_BCS_JAN13& = 104                      ' = EAN 13
Public Const PTR_BCS_EAN13& = 104                      ' = JAN 13 (added in 1.2)
Public Const PTR_BCS_TF& = 105                         ' (Discrete 2 of 5) Digits
Public Const PTR_BCS_ITF& = 106                        ' (Interleaved 2 of 5) Digits
Public Const PTR_BCS_Codabar& = 107                    ' Digits, -, $, :, /, ., +;
                                                 '   4 start/stop characters
                                                 '   (a, b, c, d)
Public Const PTR_BCS_Code39& = 108                     ' Alpha, Digits, Space, -, .,
                                                 '   $, /, +, %; start/stop (*)
                                                 ' Also has Full ASCII feature
Public Const PTR_BCS_Code93& = 109                     ' Same characters as Code 39
Public Const PTR_BCS_Code128& = 110                    ' 128 data characters
Rem *        (The following were added in Release 1.2)
Public Const PTR_BCS_UPCA_S& = 111                     ' UPC-A with supplemental
                                                 '   barcode
Public Const PTR_BCS_UPCE_S& = 112                     ' UPC-E with supplemental
                                                 '   barcode
Public Const PTR_BCS_UPCD1& = 113                      ' UPC-D1
Public Const PTR_BCS_UPCD2& = 114                      ' UPC-D2
Public Const PTR_BCS_UPCD3& = 115                      ' UPC-D3
Public Const PTR_BCS_UPCD4& = 116                      ' UPC-D4
Public Const PTR_BCS_UPCD5& = 117                      ' UPC-D5
Public Const PTR_BCS_EAN8_S& = 118                     ' EAN 8 with supplemental
                                                 '   barcode
Public Const PTR_BCS_EAN13_S& = 119                    ' EAN 13 with supplemental
                                                 '   barcode
Public Const PTR_BCS_EAN128& = 120                     ' EAN 128
Public Const PTR_BCS_OCRA& = 121                       ' OCR "A"
Public Const PTR_BCS_OCRB& = 122                       ' OCR "B"


Rem *     Two dimensional symbologies
Public Const PTR_BCS_PDF417& = 201
Public Const PTR_BCS_MAXICODE& = 202

Rem *     Start of Printer-Specific bar code symbologies
Public Const PTR_BCS_OTHER& = 501


Rem *///////////////////////////////////////////////////////////////////
Rem * "PrintBitmap" Method Constants:
Rem *///////////////////////////////////////////////////////////////////

Rem *   "Width" Parameter
Rem *     Either bitmap width or:

Public Const PTR_BM_ASIS& = -11                        ' One pixel per printer dot

Rem *   "Alignment" Parameter
Rem *     Either the distance from the left-most print column to the start
Rem *     of the bitmap, or one of the following:

Public Const PTR_BM_LEFT& = -1
Public Const PTR_BM_CENTER& = -2
Public Const PTR_BM_RIGHT& = -3


Rem *///////////////////////////////////////////////////////////////////
Rem * "RotatePrint" Method: "Rotation" Parameter Constants
Rem * "RotateSpecial" Property Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const PTR_RP_NORMAL& = &H1
Public Const PTR_RP_RIGHT90& = &H101
Public Const PTR_RP_LEFT90& = &H102
Public Const PTR_RP_ROTATE180& = &H103


Rem *///////////////////////////////////////////////////////////////////
Rem * "SetLogo" Method: "Location" Parameter Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const PTR_L_TOP& = 1
Public Const PTR_L_BOTTOM& = 2


Rem *///////////////////////////////////////////////////////////////////
Rem * "TransactionPrint" Method: "Control" Parameter Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const PTR_TP_TRANSACTION& = 11
Public Const PTR_TP_NORMAL& = 12


Rem *///////////////////////////////////////////////////////////////////
Rem * "StatusUpdateEvent" Event: "Data" Parameter Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const PTR_SUE_COVER_OPEN& = 11
Public Const PTR_SUE_COVER_OK& = 12

Public Const PTR_SUE_JRN_EMPTY& = 21
Public Const PTR_SUE_JRN_NEAREMPTY& = 22
Public Const PTR_SUE_JRN_PAPEROK& = 23

Public Const PTR_SUE_REC_EMPTY& = 24
Public Const PTR_SUE_REC_NEAREMPTY& = 25
Public Const PTR_SUE_REC_PAPEROK& = 26

Public Const PTR_SUE_SLP_EMPTY& = 27
Public Const PTR_SUE_SLP_NEAREMPTY& = 28
Public Const PTR_SUE_SLP_PAPEROK& = 29

Public Const PTR_SUE_IDLE& = 1001


Rem *///////////////////////////////////////////////////////////////////
Rem * "ResultCodeExtended" Property Constants for Printer
Rem *///////////////////////////////////////////////////////////////////

Public Const OPOS_EPTR_COVER_OPEN& = 1 + OPOSERREXT     ' (Several)
Public Const OPOS_EPTR_JRN_EMPTY& = 2 + OPOSERREXT      ' (Several)
Public Const OPOS_EPTR_REC_EMPTY& = 3 + OPOSERREXT      ' (Several)
Public Const OPOS_EPTR_SLP_EMPTY& = 4 + OPOSERREXT      ' (Several)
Public Const OPOS_EPTR_SLP_FORM& = 5 + OPOSERREXT       ' EndRemoval
Public Const OPOS_EPTR_TOOBIG& = 6 + OPOSERREXT         ' PrintBitmap
Public Const OPOS_EPTR_BADFORMAT& = 7 + OPOSERREXT      ' PrintBitmap



Rem *///////////////////////////////////////////////////////////////////
Rem *
Rem * From OposRod.h
Rem *
Rem *   Remote Order Display header file for OPOS Applications.
Rem *
Rem * Modification history
Rem * ------------------------------------------------------------------
Rem * 98-03-06 OPOS Release 1.3                                     BB
Rem *
Rem *///////////////////////////////////////////////////////////////////


Rem *///////////////////////////////////////////////////////////////////
Rem * "CurrentUnitID" and "UnitsOnline" Properties
Rem *  and "Units" Parameter Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const ROD_UID_1& = &H1
Public Const ROD_UID_2& = &H2
Public Const ROD_UID_3& = &H4
Public Const ROD_UID_4& = &H8
Public Const ROD_UID_5& = &H10
Public Const ROD_UID_6& = &H20
Public Const ROD_UID_7& = &H40
Public Const ROD_UID_8& = &H80
Public Const ROD_UID_9& = &H100
Public Const ROD_UID_10& = &H200
Public Const ROD_UID_11& = &H400
Public Const ROD_UID_12& = &H800
Public Const ROD_UID_13& = &H1000
Public Const ROD_UID_14& = &H2000
Public Const ROD_UID_15& = &H4000
Public Const ROD_UID_16& = &H8000
Public Const ROD_UID_17& = &H10000
Public Const ROD_UID_18& = &H20000
Public Const ROD_UID_19& = &H40000
Public Const ROD_UID_20& = &H80000
Public Const ROD_UID_21& = &H100000
Public Const ROD_UID_22& = &H200000
Public Const ROD_UID_23& = &H400000
Public Const ROD_UID_24& = &H800000
Public Const ROD_UID_25& = &H1000000
Public Const ROD_UID_26& = &H2000000
Public Const ROD_UID_27& = &H4000000
Public Const ROD_UID_28& = &H8000000
Public Const ROD_UID_29& = &H10000000
Public Const ROD_UID_30& = &H20000000
Public Const ROD_UID_31& = &H40000000
Public Const ROD_UID_32& = &H80000000


Rem *///////////////////////////////////////////////////////////////////
Rem * Broadcast Methods: "Attribute" Parameter Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const ROD_ATTR_BLINK& = &H80

Public Const ROD_ATTR_BG_BLACK& = &H0
Public Const ROD_ATTR_BG_BLUE& = &H10
Public Const ROD_ATTR_BG_GREEN& = &H20
Public Const ROD_ATTR_BG_CYAN& = &H30
Public Const ROD_ATTR_BG_RED& = &H40
Public Const ROD_ATTR_BG_MAGENTA& = &H50
Public Const ROD_ATTR_BG_BROWN& = &H60
Public Const ROD_ATTR_BG_GRAY& = &H70

Public Const ROD_ATTR_INTENSITY& = &H8

Public Const ROD_ATTR_FG_BLACK& = &H0
Public Const ROD_ATTR_FG_BLUE& = &H1
Public Const ROD_ATTR_FG_GREEN& = &H2
Public Const ROD_ATTR_FG_CYAN& = &H3
Public Const ROD_ATTR_FG_RED& = &H4
Public Const ROD_ATTR_FG_MAGENTA& = &H5
Public Const ROD_ATTR_FG_BROWN& = &H6
Public Const ROD_ATTR_FG_GRAY& = &H7


Rem *///////////////////////////////////////////////////////////////////
Rem * "DrawBox" Method: "BorderType" Parameter Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const ROD_BDR_SINGLE& = 1
Public Const ROD_BDR_DOUBLE& = 2
Public Const ROD_BDR_SOLID& = 3


Rem *///////////////////////////////////////////////////////////////////
Rem * "ControlClock" Method: "Function" Parameter Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const ROD_CLK_START& = 1
Public Const ROD_CLK_PAUSE& = 2
Public Const ROD_CLK_RESUME& = 3
Public Const ROD_CLK_MOVE& = 4
Public Const ROD_CLK_STOP& = 5


Rem *///////////////////////////////////////////////////////////////////
Rem * "ControlCursor" Method: "Function" Parameter Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const ROD_CRS_LINE& = 1
Public Const ROD_CRS_LINE_BLINK& = 2
Public Const ROD_CRS_BLOCK& = 3
Public Const ROD_CRS_BLOCK_BLINK& = 4
Public Const ROD_CRS_OFF& = 5


Rem *///////////////////////////////////////////////////////////////////
Rem * "SelectCharacterSet" Method: "CharacterSet" Parameter Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const ROD_CS_ASCII& = 998
Public Const ROD_CS_WINDOWS& = 999


Rem *///////////////////////////////////////////////////////////////////
Rem * "TransactionDisplay" Method: "Function" Parameter Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const ROD_TD_TRANSACTION& = 11
Public Const ROD_TD_NORMAL& = 12


Rem *///////////////////////////////////////////////////////////////////
Rem * "UpdateVideoRegionAttribute" Method: "Function" Parameter Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const ROD_UA_SET& = 1
Public Const ROD_UA_INTENSITY_ON& = 2
Public Const ROD_UA_INTENSITY_OFF& = 3
Public Const ROD_UA_REVERSE_ON& = 4
Public Const ROD_UA_REVERSE_OFF& = 5
Public Const ROD_UA_BLINK_ON& = 6
Public Const ROD_UA_BLINK_OFF& = 7


Rem *///////////////////////////////////////////////////////////////////
Rem * "EventTypes" Property and "DataEvent" Event: "Status" Parameter Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const ROD_DE_TOUCH_UP& = &H1
Public Const ROD_DE_TOUCH_DOWN& = &H2
Public Const ROD_DE_TOUCH_MOVE& = &H4


Rem *///////////////////////////////////////////////////////////////////
Rem * "ResultCodeExtended" Property Constants for Remote Order Display
Rem *///////////////////////////////////////////////////////////////////

Public Const OPOS_EROD_BADCLK& = 1 + OPOSERREXT         ' ControlClock
Public Const OPOS_EROD_NOCLOCKS& = 2 + OPOSERREXT       ' ControlClock
Public Const OPOS_EROD_NOREGION& = 3 + OPOSERREXT       ' RestoreVideo
                                                        '   Region
Public Const OPOS_EROD_NOBUFFERS& = 4 + OPOSERREXT      ' SaveVideoRegion
Public Const OPOS_EROD_NOROOM& = 5 + OPOSERREXT         ' SaveVideoRegion



Rem *///////////////////////////////////////////////////////////////////
Rem *
Rem * From OposScal.h
Rem *
Rem *   Scale header file for OPOS Applications.
Rem *
Rem * Modification history
Rem * ------------------------------------------------------------------
Rem * 95-12-08 OPOS Release 1.0                                     CRM
Rem *
Rem *///////////////////////////////////////////////////////////////////


Rem *///////////////////////////////////////////////////////////////////
Rem * "WeightUnit" Property Constants
Rem *///////////////////////////////////////////////////////////////////

Public Const SCAL_WU_GRAM& = 1
Public Const SCAL_WU_KILOGRAM& = 2
Public Const SCAL_WU_OUNCE& = 3
Public Const SCAL_WU_POUND& = 4


Rem *///////////////////////////////////////////////////////////////////
Rem * "ResultCodeExtended" Property Constants for Scale
Rem *///////////////////////////////////////////////////////////////////

Public Const OPOS_ESCAL_OVERWEIGHT& = 1 + OPOSERREXT    ' ReadWeight



Rem *///////////////////////////////////////////////////////////////////
Rem *
Rem * From OposScan.h
Rem *
Rem *   Scanner header file for OPOS Applications.
Rem *
Rem * Modification history
Rem * ------------------------------------------------------------------
Rem * 95-12-08 OPOS Release 1.0                                     CRM
Rem * 97-06-04 OPOS Release 1.2                                     CRM
Rem *   Add "ScanDataType" values.
Rem *
Rem *///////////////////////////////////////////////////////////////////


Rem *///////////////////////////////////////////////////////////////////
Rem * "ScanDataType" Property Constants
Rem *///////////////////////////////////////////////////////////////////

Rem * One dimensional symbologies
Public Const SCAN_SDT_UPCA& = 101                      ' Digits
Public Const SCAN_SDT_UPCE& = 102                      ' Digits
Public Const SCAN_SDT_JAN8& = 103                      ' = EAN 8
Public Const SCAN_SDT_EAN8& = 103                      ' = JAN 8 (added in 1.2)
Public Const SCAN_SDT_JAN13& = 104                     ' = EAN 13
Public Const SCAN_SDT_EAN13& = 104                     ' = JAN 13 (added in 1.2)
Public Const SCAN_SDT_TF& = 105                        ' (Discrete 2 of 5) Digits
Public Const SCAN_SDT_ITF& = 106                       ' (Interleaved 2 of 5) Digits
Public Const SCAN_SDT_Codabar& = 107                   ' Digits, -, $, :, /, ., +;
                                                 '   4 start/stop characters
                                                 '   (a, b, c, d)
Public Const SCAN_SDT_Code39& = 108                    ' Alpha, Digits, Space, -, .,
                                                 '   $, /, +, %; start/stop (*)
                                                 ' Also has Full ASCII feature
Public Const SCAN_SDT_Code93& = 109                    ' Same characters as Code 39
Public Const SCAN_SDT_Code128& = 110                   ' 128 data characters

Public Const SCAN_SDT_UPCA_S& = 111                    ' UPC-A with supplemental
                                                 '   barcode
Public Const SCAN_SDT_UPCE_S& = 112                    ' UPC-E with supplemental
                                                 '   barcode
Public Const SCAN_SDT_UPCD1& = 113                     ' UPC-D1
Public Const SCAN_SDT_UPCD2& = 114                     ' UPC-D2
Public Const SCAN_SDT_UPCD3& = 115                     ' UPC-D3
Public Const SCAN_SDT_UPCD4& = 116                     ' UPC-D4
Public Const SCAN_SDT_UPCD5& = 117                     ' UPC-D5
Public Const SCAN_SDT_EAN8_S& = 118                    ' EAN 8 with supplemental
                                                 '   barcode
Public Const SCAN_SDT_EAN13_S& = 119                   ' EAN 13 with supplemental
                                                 '   barcode
Public Const SCAN_SDT_EAN128& = 120                    ' EAN 128
Public Const SCAN_SDT_OCRA& = 121                      ' OCR "A"
Public Const SCAN_SDT_OCRB& = 122                      ' OCR "B"

Rem * Two dimensional symbologies
Public Const SCAN_SDT_PDF417& = 201
Public Const SCAN_SDT_MAXICODE& = 202

Rem * Special cases
Public Const SCAN_SDT_OTHER& = 501                     ' Start of Scanner-Specific bar
                                                 '   code symbologies
Public Const SCAN_SDT_UNKNOWN& = 0                     ' Cannot determine the barcode
                                                 '   symbology.



Rem *///////////////////////////////////////////////////////////////////
Rem *
Rem * From OposSig.h
Rem *
Rem *   Signature Capture header file for OPOS Applications.
Rem *
Rem * Modification history
Rem * ------------------------------------------------------------------
Rem * 95-12-08 OPOS Release 1.0                                     CRM
Rem *
Rem *///////////////////////////////////////////////////////////////////


Rem * No definitions required for this version.



Rem *///////////////////////////////////////////////////////////////////
Rem *
Rem * From OposTone.h
Rem *
Rem *   Tone Indicator header file for OPOS Applications.
Rem *
Rem * Modification history
Rem * ------------------------------------------------------------------
Rem * 97-06-04 OPOS Release 1.2                                     CRM
Rem *
Rem *///////////////////////////////////////////////////////////////////


Rem * No definitions required for this version.



Rem *///////////////////////////////////////////////////////////////////
Rem *
Rem * From OposTot.h
Rem *
Rem *   Hard Totals header file for OPOS Applications.
Rem *
Rem * Modification history
Rem * ------------------------------------------------------------------
Rem * 95-12-08 OPOS Release 1.0                                     CRM
Rem *
Rem *///////////////////////////////////////////////////////////////////


Rem *///////////////////////////////////////////////////////////////////
Rem * "ResultCodeExtended" Property Constants for Hard Totals
Rem *///////////////////////////////////////////////////////////////////

Public Const OPOS_ETOT_NOROOM& = 1 + OPOSERREXT         ' Create, Write
Public Const OPOS_ETOT_VALIDATION& = 2 + OPOSERREXT     ' Read, Write



Rem *End of OPOSALL.BAS*

