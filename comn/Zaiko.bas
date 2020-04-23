Attribute VB_Name = "ZAIKO"
Option Explicit
'********************************************************************
'*
'*              �݌Ƀf�[�^ �t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const ZAIKO_ID$ = "ZAIKO"

'�y�[�W�T�C�Y
Public Const ZAIKO_PG_SIZ% = 2048

'�|�W�V�����E�u���b�N
Public ZAIKO_POS As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
Type ZAIKOREC_Tag
    Soko_No(0 To 1)     As Byte     '�q�ɇ�
    Retu(0 To 1)        As Byte     '�I�ԁ@��
    Ren(0 To 1)         As Byte     '�I�ԁ@�A
    Dan(0 To 1)         As Byte     '�I�ԁ@�i
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    '2005.12.05 13-->20
    HIN_GAI(0 To 19)    As Byte     '�i�ԁi�O���j
    GOODS_ON(0 To 0)    As Byte     '���i���^�����i��
    NYUKA_DT(0 To 7)    As Byte     '���ד��t
    NYUKO_DT(0 To 7)    As Byte     '���ɓ��t
    '2005.12.05 13-->20
    HIN_NAI(0 To 19)    As Byte     '�i�ԁi�����j
    YUKO_Z_QTY(0 To 7)  As Byte     '�L���݌ɐ�
    LOCK_F(0 To 0)      As Byte     '�r���t���O
    WEL_ID(0 To 2)      As Byte     '�g�p�q�@ID
    PRG_ID(0 To 7)      As Byte     '�g�p���v���O����
    GOODS_YMD(0 To 7)   As Byte     '���i�����t
    
    '2005.12.05 ���ڒǉ�
    SHIIRE_CODE(0 To 4) As Byte     '�d���溰��
    SHIIRE_TANKA(0 To 10) As Byte   '�d���P��(9(8)V99)
    KEIJYO_YM(0 To 5)   As Byte     '�v��N��
    '2005.12.05 ���ڒǉ�
    
    
    '----------------   2010.07.08 ��
    GENSANKOKU(0 To 19)         As Byte     '���Y����
    SHIIRE_WORK_CENTER(0 To 7)  As Byte     '���ގd����ܰ�����
    ID_NO2(0 To 11)             As Byte     'ID_NO
    YOSAN_FROM(0 To 4)          As Byte     '�\�Z�P�ʁi���j
    YOSAN_TO(0 To 4)            As Byte     '�\�Z�P�ʁi��j
    '----------------   2010.07.08 ��
    
    
    FILLER(0 To 24)     As Byte     'FILLER
End Type

'�f�[�^�E�o�b�t�@
Public ZAIKOREC         As ZAIKOREC_Tag

'�L�[��`
Type KEY0_ZAIKO                    '�j�d�x�O
    Soko_No(0 To 1)     As Byte     '�q�ɇ�
    Retu(0 To 1)        As Byte     '�I�ԁ@��
    Ren(0 To 1)         As Byte     '�I�ԁ@�A
    Dan(0 To 1)         As Byte     '�I�ԁ@�i
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    HIN_GAI(0 To 19)    As Byte     '�i�ԁi�O���j
    GOODS_ON(0 To 0)    As Byte     '���i���^�����i��
    NYUKA_DT(0 To 7)    As Byte     '���ד��t
End Type

Type KEY1_ZAIKO                     '�j�d�x�P
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    HIN_GAI(0 To 19)    As Byte     '�i�ԁi�O���j
    GOODS_ON(0 To 0)    As Byte     '���i���^�����i��
    NYUKA_DT(0 To 7)    As Byte     '���ד��t
    Soko_No(0 To 1)     As Byte     '�q�ɇ�
    Retu(0 To 1)        As Byte     '�I�ԁ@��
    Ren(0 To 1)         As Byte     '�I�ԁ@�A
    Dan(0 To 1)         As Byte     '�I�ԁ@�i
End Type

Type KEY2_ZAIKO                     '�j�d�x�Q
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    HIN_GAI(0 To 19)    As Byte     '�i�ԁi�O���j
    GOODS_ON(0 To 0)    As Byte     '���i���^�����i��
    Soko_No(0 To 1)     As Byte     '�q�ɇ�
    Retu(0 To 1)        As Byte     '�I�ԁ@��
    Ren(0 To 1)         As Byte     '�I�ԁ@�A
    Dan(0 To 1)         As Byte     '�I�ԁ@�i
End Type

Type KEY3_ZAIKO                     '�j�d�x�R
    WEL_ID(0 To 2)      As Byte     '�g�p�q�@ID
    PRG_ID(0 To 7)      As Byte     '�g�p���v���O����
End Type

Type KEY4_ZAIKO                     '�j�d�x�S
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    HIN_GAI(0 To 19)    As Byte     '�i�ԁi�O���j
    Soko_No(0 To 1)     As Byte     '�q�ɇ�
    Retu(0 To 1)        As Byte     '�I�ԁ@��
    Ren(0 To 1)         As Byte     '�I�ԁ@�A
    Dan(0 To 1)         As Byte     '�I�ԁ@�i
End Type

Type KEY5_ZAIKO                     '�j�d�x�T
    Soko_No(0 To 1)     As Byte     '�q�ɇ�
    Retu(0 To 1)        As Byte     '�I�ԁ@��
    Ren(0 To 1)         As Byte     '�I�ԁ@�A
    Dan(0 To 1)         As Byte     '�I�ԁ@�i
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    HIN_GAI(0 To 19)    As Byte     '�i�ԁi�O���j
    NYUKA_DT(0 To 7)    As Byte     '���ד��t
End Type

Type KEY6_ZAIKO                     '�j�d�x�U
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    HIN_GAI(0 To 19)    As Byte     '�i�ԁi�O���j
    NYUKA_DT(0 To 7)    As Byte     '���ד��t
    Soko_No(0 To 1)     As Byte     '�q�ɇ�
    Retu(0 To 1)        As Byte     '�I�ԁ@��
    Ren(0 To 1)         As Byte     '�I�ԁ@�A
    Dan(0 To 1)         As Byte     '�I�ԁ@�i
End Type

'�L�[�E�f�[�^
Public K0_ZAIKO         As KEY0_ZAIKO
Public K1_ZAIKO         As KEY1_ZAIKO
Public K2_ZAIKO         As KEY2_ZAIKO
Public K3_ZAIKO         As KEY3_ZAIKO
Public K4_ZAIKO         As KEY4_ZAIKO
Public K5_ZAIKO         As KEY5_ZAIKO
Public K6_ZAIKO         As KEY6_ZAIKO

Type ZAIKO_FSpeck
    fs      As BtFileSpeck          ' ̧�� ��߯��\����
    ks0     As BtKeySpeck           ' �� ��߯��\����
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
    ks3     As BtKeySpeck
    ks4     As BtKeySpeck
    ks5     As BtKeySpeck
    ks6     As BtKeySpeck
    ks7     As BtKeySpeck
    ks8     As BtKeySpeck
    ks9     As BtKeySpeck
    ks10    As BtKeySpeck
    ks11    As BtKeySpeck
    ks12    As BtKeySpeck
    ks13    As BtKeySpeck
    ks14    As BtKeySpeck
    ks15    As BtKeySpeck
    ks16    As BtKeySpeck
    ks17    As BtKeySpeck
    ks18    As BtKeySpeck
    ks19    As BtKeySpeck
    ks20    As BtKeySpeck
    ks21    As BtKeySpeck
    ks22    As BtKeySpeck
    ks23    As BtKeySpeck
    ks24    As BtKeySpeck
    ks25    As BtKeySpeck
    ks26    As BtKeySpeck
    ks27    As BtKeySpeck
    ks28    As BtKeySpeck
    ks29    As BtKeySpeck
    ks30    As BtKeySpeck
    ks31    As BtKeySpeck
    ks32    As BtKeySpeck
    ks33    As BtKeySpeck
    ks34    As BtKeySpeck
    ks35    As BtKeySpeck
    ks36    As BtKeySpeck
    ks37    As BtKeySpeck
    ks38    As BtKeySpeck
    ks39    As BtKeySpeck
    ks40    As BtKeySpeck
    ks41    As BtKeySpeck
    ks42    As BtKeySpeck
    ks43    As BtKeySpeck
    ks44    As BtKeySpeck
    ks45    As BtKeySpeck
    ks46    As BtKeySpeck
    ks47    As BtKeySpeck
    ks48    As BtKeySpeck
    ks49    As BtKeySpeck
    ks50    As BtKeySpeck
End Type

Private ZAIKO_Speck As ZAIKO_FSpeck
Private Function ZAIKO_Create() As Integer
'********************************************************************
'*
'*              �݌Ƀf�[�^�@�b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    ZAIKO_Create = True
                                            '�݌Ƀf�[�^�@�t���p�X�捞��
    sts = GetIni("FILE", ZAIKO_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [ZAIKO]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    ZAIKO_Speck.fs.recoleng = Len(ZAIKOREC)         ' ���R�[�h��
    ZAIKO_Speck.fs.PageSize = ZAIKO_PG_SIZ          ' �y�[�W�T�C�Y
    ZAIKO_Speck.fs.idexnumb = 7                     ' �C���f�b�N�X��
    ZAIKO_Speck.fs.fileflag = 0                     ' �t�@�C���t���O
    ZAIKO_Speck.fs.reserve = &H0                    ' �\��ς�
'---------------------------------------------------'
                                                    ' �L�[�O
    ZAIKO_Speck.ks0.keypos = 1                      ' �L�[�|�W�V����
    ZAIKO_Speck.ks0.keyleng = 2                     ' �L�[��
    ZAIKO_Speck.ks0.keyflag = BtKfExt + BtKfSeg     ' �L�[�t���O
    ZAIKO_Speck.ks0.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    ZAIKO_Speck.ks0.reserve = &H0                   ' �\��ς�
                                                    
    ZAIKO_Speck.ks1.keypos = 3                      ' �L�[�|�W�V����
    ZAIKO_Speck.ks1.keyleng = 2                     ' �L�[��
    ZAIKO_Speck.ks1.keyflag = BtKfExt + BtKfSeg     ' �L�[�t���O
    ZAIKO_Speck.ks1.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    ZAIKO_Speck.ks1.reserve = &H0                   ' �\��ς�
                                                    
    ZAIKO_Speck.ks2.keypos = 5                      ' �L�[�|�W�V����
    ZAIKO_Speck.ks2.keyleng = 2                     ' �L�[��
    ZAIKO_Speck.ks2.keyflag = BtKfExt + BtKfSeg     ' �L�[�t���O
    ZAIKO_Speck.ks2.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    ZAIKO_Speck.ks2.reserve = &H0                   ' �\��ς�
                                                    
    ZAIKO_Speck.ks3.keypos = 7                      ' �L�[�|�W�V����
    ZAIKO_Speck.ks3.keyleng = 2                     ' �L�[��
    ZAIKO_Speck.ks3.keyflag = BtKfExt + BtKfSeg     ' �L�[�t���O
    ZAIKO_Speck.ks3.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    ZAIKO_Speck.ks3.reserve = &H0                   ' �\��ς�
                                                    
    ZAIKO_Speck.ks4.keypos = 9                      ' �L�[�|�W�V����
    ZAIKO_Speck.ks4.keyleng = 1                     ' �L�[��
    ZAIKO_Speck.ks4.keyflag = BtKfExt + BtKfSeg     ' �L�[�t���O
    ZAIKO_Speck.ks4.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    ZAIKO_Speck.ks4.reserve = &H0                   ' �\��ς�
                                                    
    ZAIKO_Speck.ks5.keypos = 10                     ' �L�[�|�W�V����
    ZAIKO_Speck.ks5.keyleng = 1                     ' �L�[��
    ZAIKO_Speck.ks5.keyflag = BtKfExt + BtKfSeg     ' �L�[�t���O
    ZAIKO_Speck.ks5.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    ZAIKO_Speck.ks5.reserve = &H0                   ' �\��ς�
                                                    
    ZAIKO_Speck.ks6.keypos = 11                     ' �L�[�|�W�V����
    ZAIKO_Speck.ks6.keyleng = 20                    ' �L�[��
    ZAIKO_Speck.ks6.keyflag = BtKfExt + BtKfSeg     ' �L�[�t���O
    ZAIKO_Speck.ks6.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    ZAIKO_Speck.ks6.reserve = &H0                   ' �\��ς�
                                                    
    ZAIKO_Speck.ks7.keypos = 31                     ' �L�[�|�W�V����
    ZAIKO_Speck.ks7.keyleng = 1                     ' �L�[��
    ZAIKO_Speck.ks7.keyflag = BtKfExt + BtKfSeg     ' �L�[�t���O
    ZAIKO_Speck.ks7.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    ZAIKO_Speck.ks7.reserve = &H0                   ' �\��ς�
                                                    ' �L�[�O
    ZAIKO_Speck.ks8.keypos = 32                     ' �L�[�|�W�V����
    ZAIKO_Speck.ks8.keyleng = 8                     ' �L�[��
    ZAIKO_Speck.ks8.keyflag = BtKfExt               ' �L�[�t���O
    ZAIKO_Speck.ks8.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    ZAIKO_Speck.ks8.reserve = &H0                   ' �\��ς�
'---------------------------------------------------'
                                                    ' �L�[�P
    ZAIKO_Speck.ks9.keypos = 9                      ' �L�[�|�W�V����
    ZAIKO_Speck.ks9.keyleng = 1                     ' �L�[��
    ZAIKO_Speck.ks9.keyflag = BtKfExt + BtKfSeg     ' �L�[�t���O
    ZAIKO_Speck.ks9.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    ZAIKO_Speck.ks9.reserve = &H0                   ' �\��ς�
                                                    
    ZAIKO_Speck.ks10.keypos = 10                    ' �L�[�|�W�V����
    ZAIKO_Speck.ks10.keyleng = 1                    ' �L�[��
    ZAIKO_Speck.ks10.keyflag = BtKfExt + BtKfSeg    ' �L�[�t���O
    ZAIKO_Speck.ks10.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks10.reserve = &H0                  ' �\��ς�
                                                    
    ZAIKO_Speck.ks11.keypos = 11                    ' �L�[�|�W�V����
    ZAIKO_Speck.ks11.keyleng = 20                   ' �L�[��
    ZAIKO_Speck.ks11.keyflag = BtKfExt + BtKfSeg    ' �L�[�t���O
    ZAIKO_Speck.ks11.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks11.reserve = &H0                  ' �\��ς�
                                                    
    ZAIKO_Speck.ks12.keypos = 31                    ' �L�[�|�W�V����
    ZAIKO_Speck.ks12.keyleng = 1                    ' �L�[��
    ZAIKO_Speck.ks12.keyflag = BtKfExt + BtKfSeg    ' �L�[�t���O
    ZAIKO_Speck.ks12.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks12.reserve = &H0                  ' �\��ς�
                                                    
    ZAIKO_Speck.ks13.keypos = 32                    ' �L�[�|�W�V����
    ZAIKO_Speck.ks13.keyleng = 8                    ' �L�[��
    ZAIKO_Speck.ks13.keyflag = BtKfExt + BtKfSeg    ' �L�[�t���O
    ZAIKO_Speck.ks13.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks13.reserve = &H0                  ' �\��ς�
                                                    
    ZAIKO_Speck.ks14.keypos = 1                     ' �L�[�|�W�V����
    ZAIKO_Speck.ks14.keyleng = 2                    ' �L�[��
    ZAIKO_Speck.ks14.keyflag = BtKfExt + BtKfSeg    ' �L�[�t���O
    ZAIKO_Speck.ks14.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks14.reserve = &H0                  ' �\��ς�
                                                    
    ZAIKO_Speck.ks15.keypos = 3                     ' �L�[�|�W�V����
    ZAIKO_Speck.ks15.keyleng = 2                    ' �L�[��
    ZAIKO_Speck.ks15.keyflag = BtKfExt + BtKfSeg    ' �L�[�t���O
    ZAIKO_Speck.ks15.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks15.reserve = &H0                  ' �\��ς�
                                                    
    ZAIKO_Speck.ks16.keypos = 5                     ' �L�[�|�W�V����
    ZAIKO_Speck.ks16.keyleng = 2                    ' �L�[��
    ZAIKO_Speck.ks16.keyflag = BtKfExt + BtKfSeg    ' �L�[�t���O
    ZAIKO_Speck.ks16.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks16.reserve = &H0                  ' �\��ς�
                                                    
    ZAIKO_Speck.ks17.keypos = 7                     ' �L�[�|�W�V����
    ZAIKO_Speck.ks17.keyleng = 2                    ' �L�[��
    ZAIKO_Speck.ks17.keyflag = BtKfExt              ' �L�[�t���O
    ZAIKO_Speck.ks17.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks17.reserve = &H0                  ' �\��ς�
'---------------------------------------------------'
                                                    ' �L�[�Q
    ZAIKO_Speck.ks18.keypos = 9                     ' �L�[�|�W�V����
    ZAIKO_Speck.ks18.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    ZAIKO_Speck.ks18.keyflag = BtKfExt + BtKfSeg + BtKfDup
    ZAIKO_Speck.ks18.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks18.reserve = &H0                  ' �\��ς�
                                                    
    ZAIKO_Speck.ks19.keypos = 10                    ' �L�[�|�W�V����
    ZAIKO_Speck.ks19.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    ZAIKO_Speck.ks19.keyflag = BtKfExt + BtKfSeg + BtKfDup
    ZAIKO_Speck.ks19.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks19.reserve = &H0                  ' �\��ς�
                                                    
    ZAIKO_Speck.ks20.keypos = 11                    ' �L�[�|�W�V����
    ZAIKO_Speck.ks20.keyleng = 20                   ' �L�[��
                                                    ' �L�[�t���O
    ZAIKO_Speck.ks20.keyflag = BtKfExt + BtKfSeg + BtKfDup
    ZAIKO_Speck.ks20.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks20.reserve = &H0                  ' �\��ς�
                                                    
    ZAIKO_Speck.ks21.keypos = 31                    ' �L�[�|�W�V����
    ZAIKO_Speck.ks21.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    ZAIKO_Speck.ks21.keyflag = BtKfExt + BtKfSeg + BtKfDup
    ZAIKO_Speck.ks21.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks21.reserve = &H0                  ' �\��ς�
                                                    
    ZAIKO_Speck.ks22.keypos = 1                     ' �L�[�|�W�V����
    ZAIKO_Speck.ks22.keyleng = 2                    ' �L�[��
                                                    ' �L�[�t���O
    ZAIKO_Speck.ks22.keyflag = BtKfExt + BtKfSeg + BtKfDup
    ZAIKO_Speck.ks22.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks22.reserve = &H0                  ' �\��ς�
                                                    
    ZAIKO_Speck.ks23.keypos = 3                     ' �L�[�|�W�V����
    ZAIKO_Speck.ks23.keyleng = 2                    ' �L�[��
                                                    ' �L�[�t���O
    ZAIKO_Speck.ks23.keyflag = BtKfExt + BtKfSeg + BtKfDup
    ZAIKO_Speck.ks23.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks23.reserve = &H0                  ' �\��ς�
                                                    
    ZAIKO_Speck.ks24.keypos = 5                     ' �L�[�|�W�V����
    ZAIKO_Speck.ks24.keyleng = 2                    ' �L�[��
                                                    ' �L�[�t���O
    ZAIKO_Speck.ks24.keyflag = BtKfExt + BtKfSeg + BtKfDup
    ZAIKO_Speck.ks24.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks24.reserve = &H0                  ' �\��ς�
                                                    
    ZAIKO_Speck.ks25.keypos = 7                     ' �L�[�|�W�V����
    ZAIKO_Speck.ks25.keyleng = 2                    ' �L�[��
    ZAIKO_Speck.ks25.keyflag = BtKfExt + BtKfDup    ' �L�[�t���O
    ZAIKO_Speck.ks25.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks25.reserve = &H0                  ' �\��ς�
'---------------------------------------------------'
                                                    ' �L�[�R
    ZAIKO_Speck.ks26.keypos = 77                    ' �L�[�|�W�V����
    ZAIKO_Speck.ks26.keyleng = 3                    ' �L�[��
                                                    ' �L�[�t���O
    ZAIKO_Speck.ks26.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg
    ZAIKO_Speck.ks26.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks26.reserve = &H0                  ' �\��ς�
                                                    
    ZAIKO_Speck.ks27.keypos = 80                    ' �L�[�|�W�V����
    ZAIKO_Speck.ks27.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    ZAIKO_Speck.ks27.keyflag = BtKfExt + BtKfDup + BtKfChg
    ZAIKO_Speck.ks27.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks27.reserve = &H0                  ' �\��ς�
'---------------------------------------------------'
                                                    ' �L�[�S
    ZAIKO_Speck.ks28.keypos = 9                     ' �L�[�|�W�V����
    ZAIKO_Speck.ks28.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    ZAIKO_Speck.ks28.keyflag = BtKfExt + BtKfSeg + BtKfDup
    ZAIKO_Speck.ks28.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks28.reserve = &H0                  ' �\��ς�

    ZAIKO_Speck.ks29.keypos = 10                    ' �L�[�|�W�V����
    ZAIKO_Speck.ks29.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    ZAIKO_Speck.ks29.keyflag = BtKfExt + BtKfSeg + BtKfDup
    ZAIKO_Speck.ks29.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks29.reserve = &H0                  ' �\��ς�

    ZAIKO_Speck.ks30.keypos = 11                    ' �L�[�|�W�V����
    ZAIKO_Speck.ks30.keyleng = 20                   ' �L�[��
                                                    ' �L�[�t���O
    ZAIKO_Speck.ks30.keyflag = BtKfExt + BtKfSeg + BtKfDup
    ZAIKO_Speck.ks30.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks30.reserve = &H0                  ' �\��ς�

    ZAIKO_Speck.ks31.keypos = 1                     ' �L�[�|�W�V����
    ZAIKO_Speck.ks31.keyleng = 2                    ' �L�[��
                                                    ' �L�[�t���O
    ZAIKO_Speck.ks31.keyflag = BtKfExt + BtKfSeg + BtKfDup
    ZAIKO_Speck.ks31.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks31.reserve = &H0                  ' �\��ς�

    ZAIKO_Speck.ks32.keypos = 3                     ' �L�[�|�W�V����
    ZAIKO_Speck.ks32.keyleng = 2                    ' �L�[��
                                                    ' �L�[�t���O
    ZAIKO_Speck.ks32.keyflag = BtKfExt + BtKfSeg + BtKfDup
    ZAIKO_Speck.ks32.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks32.reserve = &H0                  ' �\��ς�

    ZAIKO_Speck.ks33.keypos = 5                     ' �L�[�|�W�V����
    ZAIKO_Speck.ks33.keyleng = 2                    ' �L�[��
                                                    ' �L�[�t���O
    ZAIKO_Speck.ks33.keyflag = BtKfExt + BtKfSeg + BtKfDup
    ZAIKO_Speck.ks33.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks33.reserve = &H0                  ' �\��ς�

    ZAIKO_Speck.ks34.keypos = 7                     ' �L�[�|�W�V����
    ZAIKO_Speck.ks34.keyleng = 2                    ' �L�[��
                                                    ' �L�[�t���O
    ZAIKO_Speck.ks34.keyflag = BtKfExt + BtKfDup
    ZAIKO_Speck.ks34.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks34.reserve = &H0                  ' �\��ς�

'---------------------------------------------------'
                                                    ' �L�[�T
    ZAIKO_Speck.ks35.keypos = 1                     ' �L�[�|�W�V����
    ZAIKO_Speck.ks35.keyleng = 2                    ' �L�[��
                                                    ' �L�[�t���O
    ZAIKO_Speck.ks35.keyflag = BtKfExt + BtKfDup + BtKfSeg
    ZAIKO_Speck.ks35.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks35.reserve = &H0                  ' �\��ς�
                                                    
    ZAIKO_Speck.ks36.keypos = 3                     ' �L�[�|�W�V����
    ZAIKO_Speck.ks36.keyleng = 2                    ' �L�[��
                                                    ' �L�[�t���O
    ZAIKO_Speck.ks36.keyflag = BtKfExt + BtKfDup + BtKfSeg
    ZAIKO_Speck.ks36.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks36.reserve = &H0                  ' �\��ς�
                                                    
    ZAIKO_Speck.ks37.keypos = 5                     ' �L�[�|�W�V����
    ZAIKO_Speck.ks37.keyleng = 2                    ' �L�[��
                                                    ' �L�[�t���O
    ZAIKO_Speck.ks37.keyflag = BtKfExt + BtKfDup + BtKfSeg
    ZAIKO_Speck.ks37.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks37.reserve = &H0                  ' �\��ς�
                                                    
    ZAIKO_Speck.ks38.keypos = 7                     ' �L�[�|�W�V����
    ZAIKO_Speck.ks38.keyleng = 2                    ' �L�[��
                                                    ' �L�[�t���O
    ZAIKO_Speck.ks38.keyflag = BtKfExt + BtKfDup + BtKfSeg
    ZAIKO_Speck.ks38.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks38.reserve = &H0                  ' �\��ς�
                                                    
    ZAIKO_Speck.ks39.keypos = 9                     ' �L�[�|�W�V����
    ZAIKO_Speck.ks39.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    ZAIKO_Speck.ks39.keyflag = BtKfExt + BtKfDup + BtKfSeg
    ZAIKO_Speck.ks39.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks39.reserve = &H0                  ' �\��ς�
                                                    
    ZAIKO_Speck.ks40.keypos = 10                    ' �L�[�|�W�V����
    ZAIKO_Speck.ks40.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    ZAIKO_Speck.ks40.keyflag = BtKfExt + BtKfDup + BtKfSeg
    ZAIKO_Speck.ks40.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks40.reserve = &H0                  ' �\��ς�
                                                    
    ZAIKO_Speck.ks41.keypos = 11                    ' �L�[�|�W�V����
    ZAIKO_Speck.ks41.keyleng = 20                   ' �L�[��
                                                    ' �L�[�t���O
    ZAIKO_Speck.ks41.keyflag = BtKfExt + BtKfDup + BtKfSeg
    ZAIKO_Speck.ks41.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks41.reserve = &H0                  ' �\��ς�
                                                        
    ZAIKO_Speck.ks42.keypos = 32                    ' �L�[�|�W�V����
    ZAIKO_Speck.ks42.keyleng = 8                    ' �L�[��
    ZAIKO_Speck.ks42.keyflag = BtKfExt + BtKfDup    ' �L�[�t���O
    ZAIKO_Speck.ks42.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks42.reserve = &H0                  ' �\��ς�
'---------------------------------------------------'
                                                    ' �L�[�U
    ZAIKO_Speck.ks43.keypos = 9                     ' �L�[�|�W�V����
    ZAIKO_Speck.ks43.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    ZAIKO_Speck.ks43.keyflag = BtKfExt + BtKfDup + BtKfSeg
    ZAIKO_Speck.ks43.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks43.reserve = &H0                  ' �\��ς�
                                                    
    ZAIKO_Speck.ks44.keypos = 10                    ' �L�[�|�W�V����
    ZAIKO_Speck.ks44.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    ZAIKO_Speck.ks44.keyflag = BtKfExt + BtKfDup + BtKfSeg
    ZAIKO_Speck.ks44.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks44.reserve = &H0                  ' �\��ς�
                                                    
    ZAIKO_Speck.ks45.keypos = 11                    ' �L�[�|�W�V����
    ZAIKO_Speck.ks45.keyleng = 20                   ' �L�[��
                                                    ' �L�[�t���O
    ZAIKO_Speck.ks45.keyflag = BtKfExt + BtKfDup + BtKfSeg
    ZAIKO_Speck.ks45.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks45.reserve = &H0                  ' �\��ς�
                                                    
    ZAIKO_Speck.ks46.keypos = 32                    ' �L�[�|�W�V����
    ZAIKO_Speck.ks46.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    ZAIKO_Speck.ks46.keyflag = BtKfExt + BtKfDup + BtKfSeg
    ZAIKO_Speck.ks46.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks46.reserve = &H0                  ' �\��ς�
                                                    
    ZAIKO_Speck.ks47.keypos = 1                     ' �L�[�|�W�V����
    ZAIKO_Speck.ks47.keyleng = 2                    ' �L�[��
    ZAIKO_Speck.ks47.keyflag = BtKfExt + BtKfDup + BtKfSeg
                                                    ' �L�[�t���O
    ZAIKO_Speck.ks47.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks47.reserve = &H0                  ' �\��ς�
                                                    
    ZAIKO_Speck.ks48.keypos = 3                     ' �L�[�|�W�V����
    ZAIKO_Speck.ks48.keyleng = 2                    ' �L�[��
                                                    ' �L�[�t���O
    ZAIKO_Speck.ks48.keyflag = BtKfExt + BtKfDup + BtKfSeg
    ZAIKO_Speck.ks48.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks48.reserve = &H0                  ' �\��ς�
                                                    
    ZAIKO_Speck.ks49.keypos = 5                     ' �L�[�|�W�V����
    ZAIKO_Speck.ks49.keyleng = 2                    ' �L�[��
                                                    ' �L�[�t���O
    ZAIKO_Speck.ks49.keyflag = BtKfExt + BtKfDup + BtKfSeg
    ZAIKO_Speck.ks49.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks49.reserve = &H0                  ' �\��ς�
                                                    
    ZAIKO_Speck.ks50.keypos = 7                     ' �L�[�|�W�V����
    ZAIKO_Speck.ks50.keyleng = 2                    ' �L�[��
    ZAIKO_Speck.ks50.keyflag = BtKfExt + BtKfDup    ' �L�[�t���O
    ZAIKO_Speck.ks50.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    ZAIKO_Speck.ks50.reserve = &H0                  ' �\��ς�

'---------------------------------------------------'
    sts = BTRV(BtOpCreate, ZAIKO_POS, ZAIKO_Speck, Len(ZAIKO_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�݌Ƀf�[�^")
        Exit Function
    End If
    ZAIKO_Create = False
End Function
Public Function ZAIKO_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �݌Ƀf�[�^�@�n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    ZAIKO_Open = True
                                            '�݌Ƀf�[�^�@�t���p�X�捞��
    sts = GetIni("FILE", ZAIKO_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [ZAIKO]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = ZAIKO_Create()        '�݌Ƀf�[�^�@�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�݌Ƀf�[�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�݌Ƀf�[�^")
                Exit Function
        End Select
    Loop
    ZAIKO_Open = False

End Function

