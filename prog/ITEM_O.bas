Attribute VB_Name = "ITEM_O"
Option Explicit
'********************************************************************
'*
'*              åã@©ÏpiÚ}X^  t@Cè`
'*
'*          CREATE 2016.05.24
'********************************************************************
't@Chc
Public Const ITEM_O_ID$ = "ITEM_O"

'y[WTCY
Public Const ITEM_O_PG_SIZ% = 4096

'|WVEubN
Public ITEM_O_POS               As POSBLK
'********************************************************************
'*
'*                           \¢Ìè`
'*
'********************************************************************
'*************************** Ú¼è` *****************************



'R[hè`
Type ITEM_O_REC_Tag
    JGYOBU(0 To 0)              As Byte     'Ææª
    NAIGAI(0 To 0)              As Byte     'àO
    HIN_GAI(0 To 19)            As Byte     'iÔ(O)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.11.07
    SEQ_NO(0 To 2)              As Byte     'SEQ_NO
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.11.07
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.09.28
    KO_JGYOBU(0 To 0)           As Byte     'q@Ææª
    KO_NAIGAI(0 To 0)           As Byte     'q@àO
    KO_HIN_GAI(0 To 19)         As Byte     'q@iÔ(O)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.09.28
    
    
    NAKANISHI_TANI(0 To 3)      As Byte     '¼H¿@PÊ
    NAKANISHI_KIN(0 To 10)      As Byte     '¼H¿@àz

    SHOHIN_TANI(0 To 3)         As Byte     '¤i»H¿@PÊ
    SHOHIN_KIN(0 To 10)         As Byte     '¤i»H¿@àz

    PF_KAKOU_TANI(0 To 3)       As Byte     'PFÁH@PÊ
    PF_KAKOU_KIN(0 To 10)       As Byte     'PFÁH@àz

    PE_KAKOU_TANI(0 To 3)       As Byte     'PEÁH@PÊ
    PE_KAKOU_KIN(0 To 10)       As Byte     'PEÁH@àz

    PE_SHIZAI_TANI(0 To 3)      As Byte     'PFÞ@PÊ
    PE_SHIZAI_KIN(0 To 10)      As Byte     'PFÞ@àz

    HINBAN_LABEL_TANI(0 To 3)   As Byte     'iÔ\¦×ÍÞÙ@PÊ
    HINBAN_LABEL_KIN(0 To 10)   As Byte     'iÔ\¦×ÍÞÙ@àz

    KOUJI_SETSU_TANI(0 To 3)    As Byte     'ÝuHà¾@PÊ
    KOUJI_SETSU_KIN(0 To 10)    As Byte     'ÝuHà¾@àz

    KONPOU_TANI(0 To 3)         As Byte     '«ïÞ@PÊ
    KONPOU_KIN(0 To 10)         As Byte     '«ïÞ@àz

    FUKU_SHIZAI_TANI(0 To 3)    As Byte     'Þ@PÊ
    FUKU_SHIZAI_KIN(0 To 10)    As Byte     'Þ@àz

    KONPOU_ASSY_TANI(0 To 3)    As Byte     '«ïASSY@PÊ
    KONPOU_ASSY_KIN(0 To 10)    As Byte     '«ïASSY@àz

    KANRI_TANI(0 To 3)          As Byte     'Çï@PÊ
    KANRI_KIN(0 To 10)          As Byte     'Çï@àz
    
    GOUKEI_KIN(0 To 10)         As Byte     'vàz
    
        
    
    INPUT_TANTO_CODE(0 To 4)    As Byte     'üÍSÒº°ÄÞ
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.09.28
    BUZAI_TANTO_NAME(0 To 19)   As Byte     'ÞSÒ¼
    T_HIN_NAME(0 To 39)         As Byte     'ñoi¼
    TANI(0 To 3)                As Byte     'PÊ
    T_TANKA(0 To 10)            As Byte     'ñoP¿
    T_KINGAKU(0 To 10)          As Byte     'ñoàz
    NAKANISHI_T_KIN(0 To 10)    As Byte     '¼H¿@ñoàz
    SHOHIN_T_KIN(0 To 10)       As Byte     '¤i»H¿@ñoàz
    PF_KAKOU_T_KIN(0 To 10)     As Byte     'PFÁH@ñoàz
    PE_KAKOU_T_KIN(0 To 10)     As Byte     'PEÁH@ñoàz
    PE_SHIZAI_T_KIN(0 To 10)    As Byte     'PFÞ@ñoàz
    HINBAN_LABEL_T_KIN(0 To 10) As Byte     'iÔ\¦×ÍÞÙ@ñoàz
    KOUJI_SETSU_T_KIN(0 To 10)  As Byte     'ÝuHà¾@ñoàz
    KONPOU_T_KIN(0 To 10)       As Byte     '«ïÞ@ñoàz
    FUKU_SHIZAI_T_KIN(0 To 10)  As Byte     'Þ@ñoàz
    KONPOU_ASSY_T_KIN(0 To 10)  As Byte     '«ïASSY@ñoàz
    KANRI_T_KIN(0 To 10)        As Byte     'Çï@ñoàz
    GOUKEI_T_KIN(0 To 10)       As Byte     'ñovàz

    NAKANISHI_F(0 To 0)         As Byte     '¼H¿ ©Ï\¦Ì×¸Þ
    SHOHIN_F(0 To 0)            As Byte     '¤i»H¿ ©Ï\¦Ì×¸Þ
    PF_KAKOU_F(0 To 0)          As Byte     'PFÁH ©Ï\¦Ì×¸Þ
    PE_KAKOU_F(0 To 0)          As Byte     'PEÁH ©Ï\¦Ì×¸Þ
    PE_SHIZAI_F(0 To 0)         As Byte     'PFÞ ©Ï\¦Ì×¸Þ
    HINBAN_LABEL_F(0 To 0)      As Byte     'iÔ\¦×ÍÞÙ ©Ï\¦Ì×¸Þ
    KOUJI_SETSU_F(0 To 0)       As Byte     'ÝuHà¾ ©Ï\¦Ì×¸Þ
    KONPOU_F(0 To 0)            As Byte     '«ïÞ ©Ï\¦Ì×¸Þ
    FUKU_SHIZAI_F(0 To 0)       As Byte     'Þ ©Ï\¦Ì×¸Þ
    KONPOU_ASSY_F(0 To 0)       As Byte     '«ïASSY ©Ï\¦Ì×¸Þ
    KANRI_F(0 To 0)             As Byte     'Çï ©Ï\¦Ì×¸Þ

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.09.28
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.11.07
    KO_QTY(0 To 5)              As Byte     'õ
    NAKANISHI_QTY(0 To 5)       As Byte     '¼H¿@Ê
    SHOHIN_QTY(0 To 5)          As Byte     '¤i»H¿@Ê
    PF_KAKOU_QTY(0 To 5)        As Byte     'PFÁH@Ê
    PE_KAKOU_QTY(0 To 5)        As Byte     'PEÁH@Ê
    PE_SHIZAI_QTY(0 To 5)       As Byte     'PFÞ@Ê
    HINBAN_LABEL_QTY(0 To 5)    As Byte     'iÔ\¦×ÍÞÙ@Ê
    KOUJI_SETSU_QTY(0 To 5)     As Byte     'ÝuHà¾@Ê
    KONPOU_QTY(0 To 5)          As Byte     '«ïÞ@Ê
    FUKU_SHIZAI_QTY(0 To 5)     As Byte     'Þ@Ê
    KONPOU_ASSY_QTY(0 To 5)     As Byte     '«ïASSY@Ê
    KANRI_QTY(0 To 5)           As Byte     'Çï@Ê
    
    
    NAKANISHI_T_TAN(0 To 10)    As Byte     '¼H¿@ñoP¿
    SHOHIN_T_TAN(0 To 10)       As Byte     '¤i»H¿@ñoP¿
    PF_KAKOU_T_TAN(0 To 10)     As Byte     'PFÁH@ñoP¿
    PE_KAKOU_T_TAN(0 To 10)     As Byte     'PEÁH@ñoP¿
    PE_SHIZAI_T_TAN(0 To 10)    As Byte     'PFÞ@ñoP¿
    HINBAN_LABEL_T_TAN(0 To 10) As Byte     'iÔ\¦×ÍÞÙ@ñoP¿
    KOUJI_SETSU_T_TAN(0 To 10)  As Byte     'ÝuHà¾@ñoP¿
    KONPOU_T_TAN(0 To 10)       As Byte     '«ïÞ@ñoP¿
    FUKU_SHIZAI_T_TAN(0 To 10)  As Byte     'Þ@ñoP¿
    KONPOU_ASSY_T_TAN(0 To 10)  As Byte     '«ïASSY@ñoP¿
    KANRI_T_TAN(0 To 10)        As Byte     'Çï@ñoP¿
    
    KO_SYUBETSU(0 To 1)         As Byte         'q@íÊ
    
    FILLER(0 To 323)            As Byte
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.11.07
    
    INS_TANTO(0 To 9)           As Byte     'ÇÁ@SÒ
    Ins_DateTime(0 To 13)       As Byte     'ÇÁ@ú
    UPD_TANTO(0 To 9)           As Byte     'XV@SÒ
    UPD_DATETIME(0 To 13)       As Byte     'XV@ú

End Type
'f[^Eobt@
Public ITEM_O_REC               As ITEM_O_REC_Tag

'L[è`

Type KEY0_ITEM_O                'jdxO
    JGYOBU(0 To 0)              As Byte     'Ææª
    NAIGAI(0 To 0)              As Byte     'àO
    HIN_GAI(0 To 19)            As Byte     'iÔ(O)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.11.07
    SEQ_NO(0 To 2)              As Byte     'SEQ_NO
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.11.07

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.09.28-->2017.11.07
'    KO_JGYOBU(0 To 0)           As Byte     'q@Ææª
'    KO_NAIGAI(0 To 0)           As Byte     'q@àO
'    KO_HIN_GAI(0 To 19)         As Byte     'q@iÔ(O)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.09.28-->2017.11.07

End Type


Type KEY1_ITEM_O                'jdxO
    JGYOBU(0 To 0)              As Byte     'Ææª
    NAIGAI(0 To 0)              As Byte     'àO
    HIN_GAI(0 To 19)            As Byte     'iÔ(O)

    KO_JGYOBU(0 To 0)           As Byte     'q@Ææª
    KO_NAIGAI(0 To 0)           As Byte     'q@àO
    KO_HIN_GAI(0 To 19)         As Byte     'q@iÔ(O)

    SEQ_NO(0 To 2)              As Byte     'SEQ_NO

End Type




'L[Ef[^
Public K0_ITEM_O                As KEY0_ITEM_O
Public K1_ITEM_O                As KEY1_ITEM_O

Type ITEM_O_FSpeck
    fs      As BtFileSpeck                  ' Ì§²Ù ½Íß¯¸\¢Ì
    ks0     As BtKeySpeck                   ' ·° ½Íß¯¸\¢Ì
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
    ks3     As BtKeySpeck                   ' ·° ½Íß¯¸\¢Ì    2017.09.28
    
    ks4     As BtKeySpeck                   ' ·° ½Íß¯¸\¢Ì    2017.09.28
    ks5     As BtKeySpeck                   ' ·° ½Íß¯¸\¢Ì    2017.09.28
    ks6     As BtKeySpeck                   ' ·° ½Íß¯¸\¢Ì    2017.09.28
    ks7     As BtKeySpeck                   ' ·° ½Íß¯¸\¢Ì    2017.09.28
    ks8     As BtKeySpeck                   ' ·° ½Íß¯¸\¢Ì    2017.09.28
    ks9     As BtKeySpeck                   ' ·° ½Íß¯¸\¢Ì    2017.09.28
    ks10     As BtKeySpeck                   ' ·° ½Íß¯¸\¢Ì    2017.09.28

End Type

Private ITEM_O_Speck            As ITEM_O_FSpeck

Private Function ITEM_O_Create() As Integer
'********************************************************************
'*
'*              åã@©ÏpiÚ}X^  CREATE
'*
'*      ø  :Èµ
'*      ßèl:false ³í
'*             true  Ùí
'*
'********************************************************************

Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    ITEM_O_Create = True
                                        'åã@©ÏpiÚ}X^ tpXæÝ
    sts = GetIni("FILE", ITEM_O_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [ITEM_O]ÇÝÝG[ ")
        Exit Function
    End If

    FullPath = RTrim(c)

    ITEM_O_Speck.fs.recoleng = Len(ITEM_O_REC)          ' R[h·
    ITEM_O_Speck.fs.PageSize = ITEM_O_PG_SIZ            ' y[WTCY
    ITEM_O_Speck.fs.idexnumb = 1                        ' CfbNX
    ITEM_O_Speck.fs.fileflag = 0                        ' t@CtO
    ITEM_O_Speck.fs.reserve = &H0                       ' \ñÏÝ
'-----------------------------------------------
                                                ' L[O
    ITEM_O_Speck.ks0.keypos = 1                         ' L[|WV
    ITEM_O_Speck.ks0.keyleng = 1                        ' L[·
    ITEM_O_Speck.ks0.keyflag = BtKfExt + BtKfSeg        ' L[tO
    ITEM_O_Speck.ks0.keytype = Chr(BtKtString)          ' L[^Cv
    ITEM_O_Speck.ks0.reserve = &H0                      ' \ñÏÝ

    ITEM_O_Speck.ks1.keypos = 2                         ' L[|WV
    ITEM_O_Speck.ks1.keyleng = 1                        ' L[·
    ITEM_O_Speck.ks1.keyflag = BtKfExt + BtKfSeg        ' L[tO
    ITEM_O_Speck.ks1.keytype = Chr(BtKtString)          ' L[^Cv
    ITEM_O_Speck.ks1.reserve = &H0                      ' \ñÏÝ

    ITEM_O_Speck.ks2.keypos = 3                         ' L[|WV
    ITEM_O_Speck.ks2.keyleng = 20                       ' L[·
    ITEM_O_Speck.ks2.keyflag = BtKfExt + BtKfSeg        ' L[tO
    ITEM_O_Speck.ks2.keytype = Chr(BtKtString)          ' L[^Cv
    ITEM_O_Speck.ks2.reserve = &H0                      ' \ñÏÝ


    ITEM_O_Speck.ks3.keypos = 23                        ' L[|WV
    ITEM_O_Speck.ks3.keyleng = 3                        ' L[·
    ITEM_O_Speck.ks3.keyflag = BtKfExt                  ' L[tO
    ITEM_O_Speck.ks3.keytype = Chr(BtKtString)          ' L[^Cv
    ITEM_O_Speck.ks3.reserve = &H0                      ' \ñÏÝ



'-----------------------------------------------
    sts = BTRV(BtOpCreate, ITEM_O_POS, ITEM_O_Speck, Len(ITEM_O_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "åã@©ÏpiÚ}X^")
        Exit Function
    End If

    ITEM_O_Create = False

End Function

Public Function ITEM_O_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              åã@©ÏpiÚ}X^  nodm
'*
'*      ø  :Open Mode(BtrieveQÆ)
'*      ßèl:false ³í
'*             true  Ùí
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    ITEM_O_Open = True
                                            'åã@©ÏpiÚ}X^ tpXæÝ
    sts = GetIni("FILE", ITEM_O_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [ITEM_O]ÇÝÝG[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, ITEM_O_POS, ITEM_O_REC, Len(ITEM_O_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = ITEM_O_Create()    'åã@©ÏpiÚ}X^ì¬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ITEM_O_POS, ITEM_O_REC, Len(ITEM_O_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "åã@©ÏpiÚ}X^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "åã@©ÏpiÚ}X^")
                Exit Function
        End Select
    Loop

    ITEM_O_Open = False

End Function

Public Sub Rclr_ITEM_O_REC()

'********************************************************************
'*
'*              åã@©ÏpiÚ}X^  R[hú»
'*
'********************************************************************

    Call UniCode_Conv(ITEM_O_REC.JGYOBU, "")            'Ææª
    Call UniCode_Conv(ITEM_O_REC.NAIGAI, "")            'àO
    Call UniCode_Conv(ITEM_O_REC.HIN_GAI, "")           'iÔiOj


    Call UniCode_Conv(ITEM_O_REC.KO_JGYOBU, "")         'Ææª     2017.09.28
    Call UniCode_Conv(ITEM_O_REC.KO_NAIGAI, "")         'àO         2017.09.28
    Call UniCode_Conv(ITEM_O_REC.KO_HIN_GAI, "")        'iÔiOj   2017.09.28


    Call UniCode_Conv(ITEM_O_REC.NAKANISHI_TANI, "")    '¼H¿@PÊ
    Call UniCode_Conv(ITEM_O_REC.NAKANISHI_KIN, "")     '¼H¿@àz

    Call UniCode_Conv(ITEM_O_REC.SHOHIN_TANI, "")       '¤i»H¿@PÊ
    Call UniCode_Conv(ITEM_O_REC.SHOHIN_KIN, "")        '¤i»H¿@àz

    Call UniCode_Conv(ITEM_O_REC.PF_KAKOU_TANI, "")     'PFÁH@PÊ
    Call UniCode_Conv(ITEM_O_REC.PF_KAKOU_KIN, "")      'PFÁH@àz

    Call UniCode_Conv(ITEM_O_REC.PE_KAKOU_TANI, "")     'PEÁH@PÊ
    Call UniCode_Conv(ITEM_O_REC.PE_KAKOU_KIN, "")      'PEÁH@àz

    Call UniCode_Conv(ITEM_O_REC.PE_SHIZAI_TANI, "")    'PFÞ@PÊ
    Call UniCode_Conv(ITEM_O_REC.PE_SHIZAI_KIN, "")     'PFÞ@àz

    Call UniCode_Conv(ITEM_O_REC.HINBAN_LABEL_TANI, "") 'iÔ\¦×ÍÞÙ@PÊ
    Call UniCode_Conv(ITEM_O_REC.HINBAN_LABEL_KIN, "") 'iÔ\¦×ÍÞÙ@àz

    Call UniCode_Conv(ITEM_O_REC.KOUJI_SETSU_TANI, "")  'ÝuHà¾@PÊ
    Call UniCode_Conv(ITEM_O_REC.KOUJI_SETSU_KIN, "")  'ÝuHà¾@àz

    Call UniCode_Conv(ITEM_O_REC.KONPOU_TANI, "")       '«ïÞ@PÊ
    Call UniCode_Conv(ITEM_O_REC.KONPOU_KIN, "")        '«ïÞ@àz

    Call UniCode_Conv(ITEM_O_REC.FUKU_SHIZAI_TANI, "")  'Þ@PÊ
    Call UniCode_Conv(ITEM_O_REC.FUKU_SHIZAI_KIN, "")   'Þ@àz

    Call UniCode_Conv(ITEM_O_REC.KONPOU_ASSY_TANI, "")  '«ïASSY@PÊ
    Call UniCode_Conv(ITEM_O_REC.KONPOU_ASSY_KIN, "")   '«ïASSY@àz

    Call UniCode_Conv(ITEM_O_REC.KANRI_TANI, "")        'Çï@PÊ
    Call UniCode_Conv(ITEM_O_REC.KANRI_KIN, "")         'Çï@àz
    
    Call UniCode_Conv(ITEM_O_REC.GOUKEI_KIN, "")        'v@àz
    Call UniCode_Conv(ITEM_O_REC.INPUT_TANTO_CODE, "")  'üÍSÒº°ÄÞ
    
    
 '>>>>>>>>>>>>>>>>>>>>> 2017.09.27
    Call UniCode_Conv(ITEM_O_REC.BUZAI_TANTO_NAME, "")  'ÞSÒ¼
    Call UniCode_Conv(ITEM_O_REC.T_HIN_NAME, "")        'ñoi¼
    Call UniCode_Conv(ITEM_O_REC.TANI, "")              'PÊ
    Call UniCode_Conv(ITEM_O_REC.T_TANKA, "")           'ñoPÊ
    Call UniCode_Conv(ITEM_O_REC.T_KINGAKU, "")         'ñoàz
    Call UniCode_Conv(ITEM_O_REC.NAKANISHI_T_KIN, "")   '¼H¿@ñoàz
    Call UniCode_Conv(ITEM_O_REC.SHOHIN_T_KIN, "")      '¤i»H¿@ñoàz
    Call UniCode_Conv(ITEM_O_REC.PF_KAKOU_T_KIN, "")    'PFÁH@ñoàz
    Call UniCode_Conv(ITEM_O_REC.PE_KAKOU_T_KIN, "")    'PEÁH@ñoàz
    Call UniCode_Conv(ITEM_O_REC.PE_SHIZAI_T_KIN, "")   'PFÞ@ñoàz
    Call UniCode_Conv(ITEM_O_REC.HINBAN_LABEL_T_KIN, "") 'iÔ\¦×ÍÞÙ@ñoàz
    Call UniCode_Conv(ITEM_O_REC.KOUJI_SETSU_T_KIN, "") 'ÝuHà¾@ñoàz
    Call UniCode_Conv(ITEM_O_REC.KONPOU_T_KIN, "")      '«ïÞ@ñoàz
    Call UniCode_Conv(ITEM_O_REC.FUKU_SHIZAI_T_KIN, "") 'Þ@ñoàz
    Call UniCode_Conv(ITEM_O_REC.KONPOU_ASSY_T_KIN, "") '«ïASSY@ñoàz
    Call UniCode_Conv(ITEM_O_REC.KANRI_T_KIN, "")       'Çï@ñoàz
    Call UniCode_Conv(ITEM_O_REC.GOUKEI_T_KIN, "")      'ñovàz

    Call UniCode_Conv(ITEM_O_REC.NAKANISHI_F, "")       '¼H¿ ©Ï\¦Ì×¸Þ
    Call UniCode_Conv(ITEM_O_REC.SHOHIN_F, "")          '¤i»H¿ ©Ï\¦Ì×¸Þ
    Call UniCode_Conv(ITEM_O_REC.PF_KAKOU_F, "")        'PFÁH ©Ï\¦Ì×¸Þ
    Call UniCode_Conv(ITEM_O_REC.PE_KAKOU_F, "")        'PEÁH ©Ï\¦Ì×¸Þ
    Call UniCode_Conv(ITEM_O_REC.PE_SHIZAI_F, "")       'PFÞ ©Ï\¦Ì×¸Þ
    Call UniCode_Conv(ITEM_O_REC.HINBAN_LABEL_F, "")    'iÔ\¦×ÍÞÙ ©Ï\¦Ì×¸Þ
    Call UniCode_Conv(ITEM_O_REC.KOUJI_SETSU_F, "")     'ÝuHà¾ ©Ï\¦Ì×¸Þ
    Call UniCode_Conv(ITEM_O_REC.KONPOU_F, "")          '«ïÞ ©Ï\¦Ì×¸Þ
    Call UniCode_Conv(ITEM_O_REC.FUKU_SHIZAI_F, "")     'Þ ©Ï\¦Ì×¸Þ
    Call UniCode_Conv(ITEM_O_REC.KONPOU_ASSY_F, "")     '«ïASSY ©Ï\¦Ì×¸Þ
    Call UniCode_Conv(ITEM_O_REC.KANRI_F, "")           'Çï ©Ï\¦Ì×¸Þ
 '>>>>>>>>>>>>>>>>>>>>> 2017.09.27
    
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.11.07
    Call UniCode_Conv(ITEM_O_REC.KO_QTY, "")            'õ
    Call UniCode_Conv(ITEM_O_REC.NAKANISHI_QTY, "")     '¼H¿@Ê
    Call UniCode_Conv(ITEM_O_REC.SHOHIN_QTY, "")        '¤i»H¿@Ê
    Call UniCode_Conv(ITEM_O_REC.PF_KAKOU_QTY, "")      'PFÁH@Ê
    Call UniCode_Conv(ITEM_O_REC.PE_KAKOU_QTY, "")      'PEÁH@Ê
    Call UniCode_Conv(ITEM_O_REC.PE_SHIZAI_QTY, "")     'PFÞ@Ê
    Call UniCode_Conv(ITEM_O_REC.HINBAN_LABEL_QTY, "")  'iÔ\¦×ÍÞÙ@Ê
    Call UniCode_Conv(ITEM_O_REC.KOUJI_SETSU_QTY, "")   'ÝuHà¾@Ê
    Call UniCode_Conv(ITEM_O_REC.KONPOU_QTY, "")        '«ïÞ@Ê
    Call UniCode_Conv(ITEM_O_REC.FUKU_SHIZAI_QTY, "")   'Þ@Ê
    Call UniCode_Conv(ITEM_O_REC.KONPOU_ASSY_QTY, "")   '«ïASSY@Ê
    Call UniCode_Conv(ITEM_O_REC.KANRI_QTY, "")         'Çï@Ê
    
    
    Call UniCode_Conv(ITEM_O_REC.NAKANISHI_T_TAN, "")   '¼H¿@ñoP¿
    Call UniCode_Conv(ITEM_O_REC.SHOHIN_T_TAN, "")      '¤i»H¿@ñoP¿
    Call UniCode_Conv(ITEM_O_REC.PF_KAKOU_T_TAN, "")    'PFÁH@ñoP¿
    Call UniCode_Conv(ITEM_O_REC.PE_KAKOU_T_TAN, "")    'PEÁH@ñoP¿
    Call UniCode_Conv(ITEM_O_REC.PE_SHIZAI_T_TAN, "")   'PFÞ@ñoP¿
    Call UniCode_Conv(ITEM_O_REC.HINBAN_LABEL_T_TAN, "") 'iÔ\¦×ÍÞÙ@ñoP¿
    Call UniCode_Conv(ITEM_O_REC.KOUJI_SETSU_T_TAN, "") 'ÝuHà+¾@ñoP¿
    Call UniCode_Conv(ITEM_O_REC.KONPOU_T_TAN, "")      '«ïÞ@ñoP¿
    Call UniCode_Conv(ITEM_O_REC.FUKU_SHIZAI_T_TAN, "") 'Þ@ñoP¿
    Call UniCode_Conv(ITEM_O_REC.KONPOU_ASSY_T_TAN, "") '«ïASSY@ñoP¿
    Call UniCode_Conv(ITEM_O_REC.KANRI_T_TAN, "")       'Çï@ñoP¿
    
    
    
    Call UniCode_Conv(ITEM_O_REC.FILLER, "")
    
    Call UniCode_Conv(ITEM_O_REC.INS_TANTO, "")         'ÇÁ@SÒ
    Call UniCode_Conv(ITEM_O_REC.Ins_DateTime, "")      'ÇÁ@ú
    Call UniCode_Conv(ITEM_O_REC.UPD_TANTO, "")         'XV@SÒ
    Call UniCode_Conv(ITEM_O_REC.UPD_DATETIME, "")      'XV@ú



End Sub
