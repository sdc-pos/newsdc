Attribute VB_Name = "SYSBTRV"
Option Explicit
DefInt A-Z

'[ Btrieve ]


Type POSBLK
 '    PBElements As String * 128
    PBElements(0 To 127) As Byte
End Type
'Declare Function BTRVINIT Lib "VBBTRV32.DLL" (INIT As Any) As Integer
'Declare Function BTRVSTOP Lib "VBBTRV32.DLL" () As Integer
'Declare Function BTRV Lib "VBBTRV32.DLL" (ByVal OP%, PB As POSBLK, DB As Any, DL As Long, KB As Any, ByVal KL%, ByVal KN%) As Integer
                                'Pervasive 2000�p
Declare Function BTRV Lib "wbtrv32.dll" Alias "BTRCALL" (ByVal OP, Pb As POSBLK, Db As Any, DL As Integer, ByRef Kb As Any, ByVal Kl, ByVal Kn) As Integer


'<< Btrieve Operation >>
Public Const BtOpOpen% = 0
Public Const BtOpClose% = 1
Public Const BtOpInsert% = 2
Public Const BtOpUpdate% = 3
Public Const BtOpDelete% = 4
Public Const BtOpGetEqual% = 5
Public Const BtOpGetNext% = 6
Public Const BtOpGetPrev% = 7
Public Const BtOpGetGreater% = 8
Public Const BtOpGetGreaterEqual% = 9
Public Const BtOpGetLess% = 10
Public Const BtOpGetLessEqual% = 11
Public Const BtOpGetFirst% = 12
Public Const BtOpGetLast% = 13
Public Const BtOpCreate% = 14
Public Const BtOpStart% = 15
Public Const BtOpExtend% = 16
Public Const BtOpSetDir% = 17
Public Const BtOpGetDir% = 18
Public Const BtOpBeginTransaction% = 19
Public Const BtOpEndTransaction% = 20
Public Const BtOpAbortTransaction% = 21
Public Const BtOpGetPosition% = 22
Public Const BtOpGetDirect% = 23
Public Const BtOpStepDirect% = 24
Public Const BtOpStop% = 25
Public Const BtOpVersion% = 26
Public Const BtOpUnlock% = 27
Public Const BtOpReset% = 28
Public Const BtOpSetOwner% = 29
Public Const BtOpClearOwner% = 30
Public Const BtOpCreatSupIndex% = 31
Public Const BtOpDropSupIndex% = 32

Public Const BtOpBeginConcurrentTransaction% = 1019

'<< Btrieve Error Number >>
Public Const BtNoErr% = 0                '97.01.09
Public Const BtErrOperation% = 1
Public Const BtErrIOError% = 2
Public Const BtErrNoOpen% = 3
Public Const BtErrKeyNotFound% = 4
Public Const BtErrDuplicates% = 5
Public Const BtErrIvldKey% = 6
Public Const BtErrDiffKey% = 7
Public Const BtErrIvldPos% = 8
Public Const BtErrEOF% = 9
Public Const BtErrModif% = 10
Public Const BtErrIvldFileName% = 11
Public Const BtErrFileNotFound% = 12
Public Const BtErrExtension% = 13
Public Const BtErrPreOpen% = 14
Public Const BtErrPreImage% = 15
Public Const BtErrExpansion% = 16
Public Const BtErrClose% = 17
Public Const BtErrDiskFull% = 18
Public Const BtErrUnrecover% = 19
Public Const BtErrRecManage% = 20
Public Const BtErrKeyBuff% = 21
Public Const BtErrRecBuff% = 22
Public Const BtErrPos_Block% = 23
Public Const BtErrPageSize% = 24
Public Const BtErrCreateIO% = 25
Public Const BtErrNumberOfKeys% = 26
Public Const BtErrKeyPosition% = 27

'Version 5
Public Const BtErrAutoIncrement% = 55
Public Const BtErrCompressionBufferTooShort% = 58
Public Const BtErrAlreadyExists% = 59
Public Const BtErrPermission% = 94

'Version 6
Public Const BtErrINVALID_RECORD_LENGTH% = 28
Public Const BtErrINVALID_KEYLENGTH% = 29
Public Const BtErrNOT_A_BTRIEVE_FILE% = 30
Public Const BtErrFILE_ALREADY_EXTENDED% = 31
Public Const BtErrEXTEND_IO_ERROR% = 32
Public Const BtErrBTR_CANNOT_UNLOAD% = 33
Public Const BtErrINVALID_EXTENSION_NAME% = 34
Public Const BtErrDIRECTORY_ERROR% = 35
Public Const BtErrTRANSACTION_ERROR% = 36
Public Const BtErrTRANSACTION_IS_ACTIVE% = 37
Public Const BtErrTRANSACTION_FILE_IO_ERROR% = 38
Public Const BtErrEND_TRANSACTION_ERROR% = 39
Public Const BtErrTRANSACTION_MAX_FILES% = 40
Public Const BtErrOPERATION_NOT_ALLOWED% = 41
Public Const BtErrINCOMPLETE_ACCEL_ACCESS% = 42
Public Const BtErrINVALID_RECORD_ADDRESS% = 43
Public Const BtErrNULL_KEYPATH% = 44
Public Const BtErrINCONSISTENT_KEY_FLAGS% = 45
Public Const BtErrACCESS_TO_FILE_DENIED% = 46
Public Const BtErrMAXIMUM_OPEN_FILES% = 47
Public Const BtErrINVALID_ALT_SEQUENCE_DEF% = 48
Public Const BtErrKEY_TYPE_ERROR% = 49
Public Const BtErrOWNER_ALREADY_SET% = 50
Public Const BtErrINVALID_OWNER% = 51
Public Const BtErrERROR_WRITING_CACHE% = 52
Public Const BtErrINVALID_INTERFACE% = 53
Public Const BtErrVARIABLE_PAGE_ERROR% = 54
Public Const BtErrINCOMPLETE_INDEX% = 56
Public Const BtErrEXPANED_MEM_ERROR% = 57
Public Const BtErrREJECT_COUNT_REACHED% = 60
Public Const BtErrSMALL_EX_GET_BUFFER_ERROR% = 61
Public Const BtErrINVALID_GET_EXPRESSION% = 62
Public Const BtErrINVALID_EXT_INSERT_BUFF% = 63
Public Const BtErrOPTIMIZE_LIMIT_REACHED% = 64
Public Const BtErrINVALID_EXTRACTOR% = 65
Public Const BtErrRI_TOO_MANY_DATABASES% = 66
Public Const BtErrRIDDF_CANNOT_OPEN% = 67
Public Const BtErrRI_CASCADE_TOO_DEEP% = 68
Public Const BtErrRI_CASCADE_ERROR% = 69
Public Const BtErrRI_VIOLATION% = 71
Public Const BtErrRI_REFERENCED_FILE_CANNOT_OPEN% = 72
Public Const BtErrRI_OUT_OF_SYNC% = 73
Public Const BtErrEND_CHANGED_TO_ABORT% = 74
Public Const BtErrRI_CONFLICT% = 76
Public Const BtErrCANT_LOOP_IN_SERVER% = 77
Public Const BtErrDEAD_LOCK% = 78
Public Const BtErrPROGRAMMING_ERROR% = 79
Public Const BtErrCONFLICT% = 80
Public Const BtErrLOCKERROR% = 81
Public Const BtErrLOST_POSITION% = 82
Public Const BtErrREAD_OUTSIDE_TRANSACTION% = 83
Public Const BtErrRECORD_INUSE% = 84
Public Const BtErrFILE_INUSE% = 85
Public Const BtErrFILE_TABLE_FULL% = 86
Public Const BtErrNOHANDLES_AVAILABLE% = 87
Public Const BtErrINCOMPATIBLE_MODE_ERROR% = 88

Public Const BtErrDEVICE_TABLE_FULL% = 90
Public Const BtErrSERVER_ERROR% = 91
Public Const BtErrTRANSACTION_TABLE_FULL% = 92
Public Const BtErrINCOMPATIBLE_LOCK_TYPE% = 93
Public Const BtErrSESSION_NO_LONGER_VALID% = 95
Public Const BtErrCOMMUNICATIONS_ERROR% = 96
Public Const BtErrDATA_MESSAGE_TOO_SMALL% = 97
Public Const BtErrINTERNAL_TRANSACTION_ERROR% = 98
Public Const BtErrREQUESTER_CANT_ACCESS_RUNTIME% = 99
Public Const BtErrNO_CACHE_BUFFERS_AVAIL% = 100
Public Const BtErrNO_OS_MEMORY_AVAIL% = 101
Public Const BtErrNO_STACK_AVAIL% = 102
Public Const BtErrCHUNK_OFFSET_TOO_LONG% = 103
Public Const BtErrLOCALE_ERROR% = 104
Public Const BtErrCANNOT_CREATE_WITH_BAT% = 105
Public Const BtErrCHUNK_CANNOT_GET_NEXT% = 106
Public Const BtErrCHUNK_INCOMPATIBLE_FILE% = 107

Public Const BtErrTRANSACTION_TOO_COMPLEX% = 109
Public Const BtErrNO_SYSTEM_LOCKS_AVAILABLE% = 130
Public Const BtErrMORE_THAN_5_CONCURRENT_USERS% = 133


'File Flag
Public Const BtFfChg% = 1
Public Const BtFfSpc% = 2
Public Const BtFfPre% = 4
Public Const BtFfKeyOnly% = 16       '97.01.09
'Key Flag
Public Const BtKfDup% = 1            ' �d���L�[
Public Const BtKfChg% = 2            ' �ύX�\�L�[
Public Const BtKfBin% = 4            ' �o�C�i���[�L�[
Public Const BtKfNul% = 8            ' �k���E�L�[
Public Const BtKfSeg% = 16           ' �Z�O�����g�E�L�[
Public Const BtKfAlt% = 32           ' �I���^�l�[�g�E�L�[
Public Const BtKfDec% = 64           ' �~���L�[
Public Const BtKfSup% = 128          ' �T�v�������g�E�L�[
Public Const BtKfExt% = 256          ' �g���L�[

'Key Type
Public Const BtKtString% = 0
Public Const BtKtInteger% = 1
Public Const BtKtFloat% = 2
Public Const BtKtDate% = 3
Public Const BtKtTime% = 4
Public Const BtKtDecimal% = 5
Public Const BtKtMoney% = 6
Public Const BtKtLogical% = 7
Public Const BtKtNumeric% = 8
Public Const BtKtBFloat% = 9
Public Const BtKtLString% = 10
Public Const BtKtZString% = 11
Public Const BtKtUSInteger% = 14

'Open Mode                      97.01.09
Public Const BtOpenNomal% = 0
Public Const BtOpenAcess% = -1
Public Const BtOpenRead% = -2
Public Const BtOpenExec% = -4

'RECODE LOCK                    98.01.19
Public Const BtSWait% = 100      '�V���O���E�G�C�g���b�N
Public Const BtSNoWait% = 200    '�V���O���m�n�E�G�C�g���b�N
Public Const BtMWait% = 300      '�}���`�E�G�C�g���b�N
Public Const BtMNoWait% = 400    '�}���`�m�n�E�G�C�g���b�N

'NCC
Public Const BtNCC% = -1        'NCC UPDATE
'Btrive File Data Structure
Type BtFileSpeck
    recoleng As Integer
    PageSize As Integer
    idexnumb As Integer
    confnumb As String * 4
    fileflag As Integer
    dupPointers As String * 1
    reserve As String * 1
    allocation As Integer
End Type

Type BtKeySpeck
     keypos As Integer
     keyleng As Integer
     keyflag As Integer
     confnumb As String * 4
     keytype As String * 1
     nulchar As String * 1
     reserve As String * 2
     menualKeyNumber As String * 1
     acsNumber As String * 1
End Type


Sub Bt_Error(sts As Integer, Opretion As Integer, file As String)

    Dim mess As String

    mess = GetMsg_Japanese(sts)
    
    mess = file & " ERROR Operation = " & Opretion & " " & "sts= " & sts & " : " & mess
'    MsgBox mess, vbOKOnly + vbExclamation, "BTRV ERROR"
    MsgBox mess, vbOKOnly + vbExclamation

End Sub

Private Function GetMsg_English(ByVal sts As Integer) As String

    Select Case sts
        Case BtErrOperation
            GetMsg_English = "Operation Error"
        Case BtErrIOError
            GetMsg_English = "I/O ERROR"
        Case BtErrNoOpen
            GetMsg_English = "File no Open"
        Case BtErrKeyNotFound
            GetMsg_English = "Key not Found"
        Case BtErrDuplicates
            GetMsg_English = "Duplicates Error"
        Case BtErrIvldKey
            GetMsg_English = "Invalid Key Number"
        Case BtErrDiffKey
            GetMsg_English = "Different Key Number"
        Case BtErrIvldPos
            GetMsg_English = "Invalid Positioning"
        Case BtErrEOF
            GetMsg_English = "End Of File"
        Case BtErrModif
            GetMsg_English = "Modifiable Error"
        Case BtErrIvldFileName
            GetMsg_English = "Invalid File Name"
        Case BtErrFileNotFound
            GetMsg_English = "File not Found"
        Case BtErrExtension
            GetMsg_English = "Extension Error"
        Case BtErrPreOpen
            GetMsg_English = "Pre-Open Error"
        Case BtErrPreImage
            GetMsg_English = "Pre-Image Error"
        Case BtErrExpansion
            GetMsg_English = "Expansion Error"
        Case BtErrClose
            GetMsg_English = "Close Error"
        Case BtErrDiskFull
            GetMsg_English = "Disk Full"
        Case BtErrUnrecover
            GetMsg_English = "Unrecoverable Error"
        Case BtErrRecManage
            GetMsg_English = "Record Manager Inactive"
        Case BtErrKeyBuff
            GetMsg_English = "Key Buffer Error"
        Case BtErrRecBuff
            GetMsg_English = "Record Buffer Error"
        Case BtErrPos_Block
            GetMsg_English = "Position Block Error"
        Case BtErrPageSize
            GetMsg_English = "Page Size Error"
        Case BtErrCreateIO
            GetMsg_English = "Create I/O Error"
        Case BtErrNumberOfKeys
            GetMsg_English = "Number of Keys"
        Case BtErrKeyPosition
            GetMsg_English = "Key Position"

        'Ver 5
        Case BtErrAutoIncrement
            GetMsg_English = "AutoIncrement Error"
        Case BtErrCompressionBufferTooShort
            GetMsg_English = "Compression Buffer Too Short"
        Case BtErrAlreadyExists
            GetMsg_English = "File Alredy Exists"

        Case Else
            GetMsg_English = "Unknown"
    End Select

End Function

Private Function GetMsg_Japanese(ByVal sts As Integer) As String

    Select Case sts
        Case BtErrOperation
            GetMsg_Japanese = "���o�^�I�y���[�V�����ł��B"
        Case BtErrIOError
            GetMsg_Japanese = "���o�ُ͈�ł��B"
        Case BtErrNoOpen
            GetMsg_Japanese = "�t�@�C�����J����Ă��܂���B"
        Case BtErrKeyNotFound
            GetMsg_Japanese = "�L�[��������܂���B"
        Case BtErrDuplicates
            GetMsg_Japanese = "�d���f�[�^��ǉ����悤�Ƃ��܂����B"
        Case BtErrIvldKey
            GetMsg_Japanese = "�ݒ�O�̃L�[�E�i���o�[���g�p����܂����B"
        Case BtErrDiffKey
            GetMsg_Japanese = "��v���Ă��Ȃ��L�[�E�i���o�[���g�p����܂����B"
        Case BtErrIvldPos
            GetMsg_Japanese = "�|�W�V���j���O�����s����Ă��܂���B"
        Case BtErrEOF
            GetMsg_Japanese = "�ŏI�f�[�^���z���ēǂݍ������Ƃ��܂����B"
        Case BtErrModif
            GetMsg_Japanese = "�ύX�s�\�ȃL�[�E�t�B�[���h��ύX���悤�Ƃ��܂����B"
        Case BtErrIvldFileName
            GetMsg_Japanese = "�t�@�C�������s���ł��B"
        Case BtErrFileNotFound
            GetMsg_Japanese = "�t�@�C����������܂���B"
        Case BtErrExtension
            GetMsg_Japanese = "�����g���t�@�C����������܂���B"
        Case BtErrPreOpen
            GetMsg_Japanese = "�v���E�C���[�W�E�t�@�C�����J���܂���B"
        Case BtErrPreImage
            GetMsg_Japanese = "�v���E�C���[�W�E�t�@�C�����ُ�ł��B"
        Case BtErrExpansion
            GetMsg_Japanese = "�g���@�\�ُ�ł��B"
        Case BtErrClose
            GetMsg_Japanese = "�t�@�C���E�N���[�Y�ُ�ł��B"
        Case BtErrDiskFull
            GetMsg_Japanese = "�f�B�X�N�������ς��ł��B"
        Case BtErrUnrecover
            GetMsg_Japanese = "�C���s�\�Ȉُ�ł��B"
        Case BtErrRecManage
            GetMsg_Japanese = "���R�[�h�E�}�l�[�W�����N������Ă��܂���B"
        Case BtErrKeyBuff
            GetMsg_Japanese = "�L�[�E�o�b�t�@���Z�������܂��B"
        Case BtErrRecBuff
            GetMsg_Japanese = "���R�[�h�E�o�b�t�@���Z�������܂��B"
        Case BtErrPos_Block
            GetMsg_Japanese = "�|�W�V�����E�u���b�N�̃T�C�Y��128�o�C�g�ł͂���܂���B"
        Case BtErrPageSize
            GetMsg_Japanese = "�y�[�W�E�T�C�Y���ُ�ł��B"
        Case BtErrCreateIO
            GetMsg_Japanese = "�t�@�C�����쐬�ł��܂���B"
        Case BtErrNumberOfKeys
            GetMsg_Japanese = "�L�[�����ُ�ł��B"
        Case BtErrKeyPosition
            GetMsg_Japanese = "�L�[�E�|�W�V���������R�[�h�����z���Ă��܂��B"

        'Ver 5
        Case BtErrAutoIncrement
            GetMsg_Japanese = "�I�[�g�C���N�������g�E�L�[��ݒ�s�\�ȃL�[�ɐݒ肵�悤�Ƃ��܂����B"
        Case BtErrCompressionBufferTooShort
            GetMsg_Japanese = "�f�[�^�̈��k�E�����̂��߂̃o�b�t�@���s�����Ă��܂��B"
        Case BtErrAlreadyExists
            GetMsg_Japanese = "�����̃t�@�C�������݂��Ă��܂��B"

        'Ver 6
        Case BtErrINVALID_RECORD_LENGTH ' 28
            GetMsg_Japanese = "���R�[�h�����s���ł�"
        Case BtErrINVALID_KEYLENGTH ' 29
            GetMsg_Japanese = "�L�[�����s���ł�"
        Case BtErrNOT_A_BTRIEVE_FILE ' 30
            GetMsg_Japanese = "�w�肳�ꂽ�t�@�C����Btrieve�݊��t�@�C���ł͂���܂���"
        Case BtErrFILE_ALREADY_EXTENDED ' 31
            GetMsg_Japanese = "�t�@�C���͊��Ɋg������Ă��܂�"
        Case BtErrEXTEND_IO_ERROR ' 32
            GetMsg_Japanese = "�t�@�C�����g���ł��܂���"
        Case BtErrBTR_CANNOT_UNLOAD ' 33
            GetMsg_Japanese = "MKDE���A�����[�h�ł��܂���"
        Case BtErrINVALID_EXTENSION_NAME ' 34
            GetMsg_Japanese = "�g�������Ɏw�肳�ꂽ���O���s���ł�"
        Case BtErrDIRECTORY_ERROR ' 35
            GetMsg_Japanese = "�f�B���N�g���G���[���������܂���"
        Case BtErrTRANSACTION_ERROR ' 36
            GetMsg_Japanese = "�g�����U�N�V�����G���[���������܂���"
        Case BtErrTRANSACTION_IS_ACTIVE ' 37
            GetMsg_Japanese = "�ʂ̃g�����U�N�V���������s���ł�"
        Case BtErrTRANSACTION_FILE_IO_ERROR ' 38
            GetMsg_Japanese = "�g�����U�N�V��������t�@�C��I/O�G���[���������܂���"
        Case BtErrEND_TRANSACTION_ERROR ' 39
            GetMsg_Japanese = "End(,Abort) Transaction �̑O�ɂ� Begin Transaction�̎��s���K�v"
        Case BtErrTRANSACTION_MAX_FILES ' 40
            GetMsg_Japanese = "�t�@�C���A�N�Z�X�v�������e�����t�@�C���̍ő吔�𒴂���"
        Case BtErrOPERATION_NOT_ALLOWED ' 41
            GetMsg_Japanese = "���s���悤�Ƃ�������͋�����Ă��܂���"
        Case BtErrINCOMPLETE_ACCEL_ACCESS ' 42
            GetMsg_Japanese = "�ȑO�ɃA�N�Z�����C�e�B�h���[�h�ŃI�[�v�����ꂽ�t�@�C�����N���[�Y����Ă��܂���ł���"
        Case BtErrINVALID_RECORD_ADDRESS '43
            GetMsg_Japanese = "�w�肳�ꂽ���R�[�h�A�N�Z�X�͕s���ł�"
        Case BtErrNULL_KEYPATH ' 44
            GetMsg_Japanese = "�w�肳�ꂽ�L�[�̃C���f�b�N�X�p�X�͕s���ł�"
        Case BtErrINCONSISTENT_KEY_FLAGS ' 45
            GetMsg_Japanese = "�w�肳�ꂽ�L�[�����͕s���ł�"
        Case BtErrACCESS_TO_FILE_DENIED ' 46
            GetMsg_Japanese = "�v�����ꂽ�t�@�C���ւ̃A�N�Z�X�����ۂ���܂���"
        Case BtErrMAXIMUM_OPEN_FILES ' 47
            GetMsg_Japanese = "�I�[�v������Ă���t�@�C�������������ő吔�𒴂��܂���"
        Case BtErrINVALID_ALT_SEQUENCE_DEF ' 48
            GetMsg_Japanese = "�I���^�l�[�g�R���[�e�B���O�V�[�P���X��`���s���ł�"
        Case BtErrKEY_TYPE_ERROR ' 49
            GetMsg_Japanese = "�g���L�[�^�C�v���s���ł�"
        Case BtErrOWNER_ALREADY_SET ' 50
            GetMsg_Japanese = "�t�@�C���̃I�[�i�[�l�[�������ɐݒ肳��Ă��܂�"
        Case BtErrINVALID_OWNER ' 51
            GetMsg_Japanese = "�I�[�i�[�l�[�����s���ł�"
        Case BtErrERROR_WRITING_CACHE ' 52
            GetMsg_Japanese = "����C���^�[�t�F�C�X�̃o�[�W�������s���ł�"
        Case BtErrINVALID_INTERFACE ' 53
            GetMsg_Japanese = "���R�[�h�̉ϒ��������j�����Ă��܂�"
        Case BtErrVARIABLE_PAGE_ERROR ' 54
            GetMsg_Japanese = "�A�v���P�[�V�������I�[�g�C���N�������g�L�[�ɕs���ȑ�����ݒ肵�܂���"
        Case BtErrINCOMPLETE_INDEX ' 56
            GetMsg_Japanese = "�C���f�b�N�X���s���S�ł�"
        Case BtErrEXPANED_MEM_ERROR ' 57
            GetMsg_Japanese = "Expaned Memory Error!!"
        Case BtErrREJECT_COUNT_REACHED ' 60
            GetMsg_Japanese = "�w�肳�ꂽ���W�F�N�g�J�E���g�ɒB���܂���"
        Case BtErrSMALL_EX_GET_BUFFER_ERROR ' 61
            GetMsg_Japanese = "��Ɨ̈悪���������܂�"
        Case BtErrINVALID_GET_EXPRESSION ' 62
            GetMsg_Japanese = "�f�B�X�N���v�^���Ԉ���Ă��܂�"
        Case BtErrINVALID_EXT_INSERT_BUFF ' 63
            GetMsg_Japanese = "Insert Extended �I�y���[�V�����Ŏw�肳�ꂽ�f�[�^�o�b�t�@���s���ł�"
        Case BtErrOPTIMIZE_LIMIT_REACHED ' 64
            GetMsg_Japanese = "�t�B���^�����ɒB���܂���"
        Case BtErrINVALID_EXTRACTOR ' 65
            GetMsg_Japanese = "�t�B�[���h�I�t�Z�b�g���s���ł�"
        Case BtErrRI_TOO_MANY_DATABASES ' 66
            GetMsg_Japanese = "�I�[�v���ł���f�[�^�x�[�X�̍ő吔�𒴂��܂���"
        Case BtErrRIDDF_CANNOT_OPEN ' 67
            GetMsg_Japanese = "SQL �f�[�^�������I�[�v���ł��܂���"
        Case BtErrRI_CASCADE_TOO_DEEP ' 68
            GetMsg_Japanese = "RI Delete Cascade �I�y���[�V���������s�ł��܂���"
        Case BtErrRI_CASCADE_ERROR ' 69
            GetMsg_Japanese = "Delete �I�y���[�V�������A�j�󂳂�Ă���t�@�C���̃��R�[�h�ɑ΂��čs���܂���"
        Case BtErrRI_VIOLATION ' 71
            GetMsg_Japanese = "�Q�Ɛ������̒�`�Ɍ�肪�L��܂�"
        Case BtErrRI_REFERENCED_FILE_CANNOT_OPEN ' 72
            GetMsg_Japanese = "�Q�Ɛ������̃t�@�C�����I�[�v���ł��܂���"
        Case BtErrRI_OUT_OF_SYNC ' 73
            GetMsg_Japanese = "�Q�Ɛ������̒�`���H������Ă��܂�"
        Case BtErrEND_CHANGED_TO_ABORT ' 74
            GetMsg_Japanese = "�g�����U�N�V�����𒆎~���܂���"
        Case BtErrRI_CONFLICT ' 76
            GetMsg_Japanese = "�Q�Ƃ��Ă���t�@�C���ɖ������L��܂�"
        Case BtErrCANT_LOOP_IN_SERVER ' 77
            GetMsg_Japanese = "�E�F�C�g�G���[���������܂���"
        Case BtErrDEAD_LOCK ' 78
            GetMsg_Japanese = "�f�b�h���b�N�����o���܂���"
        Case BtErrPROGRAMMING_ERROR ' 79
            GetMsg_Japanese = "�v���O���~���O�G���[�������܂���"
        Case BtErrCONFLICT ' 80
            GetMsg_Japanese = "���R�[�h���x���̖����������܂���"
        Case BtErrLOCKERROR ' 81
            GetMsg_Japanese = "���b�N�G���[���������܂���"
        Case BtErrLOST_POSITION ' 82
            GetMsg_Japanese = "�|�W�V�����������܂���"
        Case BtErrREAD_OUTSIDE_TRANSACTION ' 83
            GetMsg_Japanese = "�g�����U�N�V�����O�œǂݍ��񂾃��R�[�h��ύX���悤�Ƃ��܂���"
        Case BtErrRECORD_INUSE ' 84
            GetMsg_Japanese = "���R�[�h�܂��̓y�[�W�����b�N����Ă��܂�"
        Case BtErrFILE_INUSE ' 85
            GetMsg_Japanese = "�t�@�C�������b�N����Ă��܂�"
        Case BtErrFILE_TABLE_FULL ' 86
            GetMsg_Japanese = "�t�@�C���e�[�u������t�ł�"
        Case BtErrNOHANDLES_AVAILABLE ' 87
            GetMsg_Japanese = "�n���h���e�[�u������t�ł�"
        Case BtErrINCOMPATIBLE_MODE_ERROR ' 88
            GetMsg_Japanese = "�s��v���[�h�G���[���������܂���"
        Case BtErrDEVICE_TABLE_FULL ' 90
            GetMsg_Japanese = "���_�C���N�g�f�o�C�X�e�[�u������t�ł�"
        Case BtErrSERVER_ERROR ' 91
            GetMsg_Japanese = "�T�[�o�G���[�ł�"
        Case BtErrTRANSACTION_TABLE_FULL ' 92
            GetMsg_Japanese = "�g�����U�N�V�����e�[�u������t�ł�"
        Case BtErrINCOMPATIBLE_LOCK_TYPE ' 93
            GetMsg_Japanese = "���R�[�h���b�N�̎�ނ���v���Ă��܂���"
        Case 94                          ' 94
            GetMsg_Japanese = "�p�[�~�b�V�����G���[���������܂���"
        Case BtErrSESSION_NO_LONGER_VALID ' 95
            GetMsg_Japanese = "�Z�b�V�����͊��ɖ����ɂȂ��Ă��܂�"
        Case BtErrCOMMUNICATIONS_ERROR ' 96
            GetMsg_Japanese = "�ʐM���ɃG���[���������܂���"
        Case BtErrDATA_MESSAGE_TOO_SMALL ' 97
            GetMsg_Japanese = "�f�[�^�o�b�t�@�����������܂�"
        Case BtErrINTERNAL_TRANSACTION_ERROR ' 98
            GetMsg_Japanese = "�����g�����U�N�V�����G���[�����o����܂���"
        
        Case BtErrREQUESTER_CANT_ACCESS_RUNTIME ' 99
            GetMsg_Japanese = ""
        Case BtErrNO_CACHE_BUFFERS_AVAIL ' 100
            GetMsg_Japanese = ""
        Case BtErrNO_OS_MEMORY_AVAIL ' 101
            GetMsg_Japanese = ""
        Case BtErrNO_STACK_AVAIL ' 102
            GetMsg_Japanese = ""
        Case BtErrCHUNK_OFFSET_TOO_LONG ' 103
            GetMsg_Japanese = ""
        Case BtErrLOCALE_ERROR ' 104
            GetMsg_Japanese = ""
        Case BtErrCANNOT_CREATE_WITH_BAT ' 105
            GetMsg_Japanese = ""
        Case BtErrCHUNK_CANNOT_GET_NEXT ' 106
            GetMsg_Japanese = ""
        Case BtErrCHUNK_INCOMPATIBLE_FILE ' 107
            GetMsg_Japanese = ""
        Case BtErrTRANSACTION_TOO_COMPLEX ' 109
            GetMsg_Japanese = ""
        Case BtErrNO_SYSTEM_LOCKS_AVAILABLE ' 130
            GetMsg_Japanese = ""
        Case BtErrMORE_THAN_5_CONCURRENT_USERS ' 133
            GetMsg_Japanese = ""
        
        Case Else
            GetMsg_Japanese = "�t�@�C���ُ킪�������܂����I�I"
    End Select

End Function

