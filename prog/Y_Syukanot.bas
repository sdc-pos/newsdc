Attribute VB_Name = "Y_SYU"
Option Explicit
'********************************************************************
'*                                                                  *
'*              �o�ח\��f�[�^  �t�@�C����`                        �@*
'*                                                                  *
'*          CREATE 1997.06.02  S.Shibano                            *
'********************************************************************
'�t�@�C���h�c
Global Const Y_SYU_ID = "Y_SYU"

'�y�[�W�T�C�Y
Global Const Y_SYU_PG_SIZ% = 4096

'�|�W�V�����E�u���b�N
Global Y_SYU_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                              *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type Y_SYUREC_Tag
    KAN_KBN(0 To 0) As Byte         '�����敪
    DT_SYU(0 To 0) As Byte          '�f�[�^���
    YOTEI_QTY(0 To 7) As Byte       '�\�萔��
    FIX_QTY(0 To 7) As Byte         '�m�萔��
    NAIGAI(0 To 0) As Byte          '�����O
    JGYOBU(0 To 0) As Byte          '���ƕ��敪
    TEXT_NO(0 To 8) As Byte         '�e�L�X�g��
    CYOK_KBN(0 To 0) As Byte        '�����敪
    DEN_DT(0 To 7) As Byte          '�`�[���t
    IO_KBN(0 To 0) As Byte          '���o�ɋ敪
    PM_KBN(0 To 0) As Byte          '�ԍ��敪
    DEN_SYU(0 To 0) As Byte         '�`�[���
    DEN_NO(0 To 5) As Byte          '�`�[��
    CYU_KBN(0 To 0) As Byte         '�����敪
    HIN_GAI(0 To 12) As Byte        '�i�ԁi�O���j
    HIN_NAI(0 To 12) As Byte        '�i�ԁi�����j
    HIN_NAME(0 To 24) As Byte       '�i��
    YOSAN_FROM(0 To 4) As Byte      '�\�Z�P�ʁi���j
    YOSAN_TO(0 To 4) As Byte        '�\�Z�P�ʁi��j
    HOST_SOKO(0 To 1) As Byte       '�q�ɋ敪�iνāj
    HOST_TANA(0 To 7) As Byte       '�I�ԁiνāj
    SYUK_CODE(0 To 4) As Byte       '�x����^�o�א�
    SYUK_NAME(0 To 19) As Byte      '�x����^�o�א於
    KAN_DT(0 To 7) As Byte          '�������t
    PRINT_KBN(0 To 0) As Byte       '����敪
    HS_CYU_KBN(0 To 0) As Byte      '�����敪�iνāj
    SS_KBN(0 To 0) As Byte          '�r�r�敪
    SS_CODE(0 To 1) As Byte         '�r�r�R�[�h
    FILLER(0 To 10) As Byte         'FILLER
End Type

'�f�[�^�E�o�b�t�@
Global Y_SYUREC As Y_SYUREC_Tag

'�L�[��`
Type KEY0_Y_SYU            '�j�d�x�O
    JGYOBU(0 To 0) As Byte          '���ƕ��敪
    DEN_DT(0 To 7) As Byte          '�`�[���t
    DEN_NO(0 To 5) As Byte          '�`�[��
    SS_CODE(0 To 1) As Byte         '�r�r�R�[�h
End Type

Type KEY1_Y_SYU            '�j�d�x�P
    JGYOBU(0 To 0) As Byte          '���ƕ��敪
    KAN_KBN(0 To 0) As Byte         '�����敪
    CYU_KBN(0 To 0) As Byte         '�����敪
    SYUK_CODE(0 To 4) As Byte       '�x����^�o�א�
    SS_CODE(0 To 1) As Byte         '�r�r�R�[�h
    HIN_GAI(0 To 12) As Byte        '�i�ԁi�O���j
    DEN_DT(0 To 7) As Byte          '�`�[���t
    DEN_NO(0 To 5) As Byte          '�`�[��
End Type

Type KEY2_Y_SYU            '�j�d�x�Q
    JGYOBU(0 To 0) As Byte          '���ƕ��敪
    KAN_KBN(0 To 0) As Byte         '�����敪
    CYU_KBN(0 To 0) As Byte         '�����敪
    HOST_TANA(0 To 7) As Byte       '�I�ԁiνāj
    DEN_DT(0 To 7) As Byte          '�`�[���t
End Type

Type KEY3_Y_SYU            '�j�d�x�R
    JGYOBU(0 To 0) As Byte          '���ƕ��敪
    KAN_KBN(0 To 0) As Byte         '�����敪
    SYUK_CODE(0 To 4) As Byte       '�x����^�o�א�
    CYU_KBN(0 To 0) As Byte         '�����敪
    HIN_GAI(0 To 12) As Byte        '�i�ԁi�O���j
    DEN_DT(0 To 7) As Byte          '�`�[���t
    DEN_NO(0 To 5) As Byte          '�`�[��
    SS_CODE(0 To 1) As Byte         '�r�r�R�[�h
End Type

Type KEY4_Y_SYU            '�j�d�x�S
    JGYOBU(0 To 0) As Byte          '���ƕ��敪
    KAN_KBN(0 To 0) As Byte         '�����敪
    NAIGAI(0 To 0) As Byte          '�����O
    HIN_GAI(0 To 12) As Byte        '�i�ԁi�O���j
    DEN_DT(0 To 7) As Byte          '�`�[���t
End Type

Type KEY5_Y_SYU            '�j�d�x�T
    DEN_DT(0 To 7) As Byte          '�`�[���t
End Type

'�L�[�E�f�[�^
Global K0_Y_SYU As KEY0_Y_SYU
Global K1_Y_SYU As KEY1_Y_SYU
Global K2_Y_SYU As KEY2_Y_SYU
Global K3_Y_SYU As KEY3_Y_SYU
Global K4_Y_SYU As KEY4_Y_SYU
Global K5_Y_SYU As KEY5_Y_SYU

Type Y_SYU_FSpeck
    fs As BtFileSpeck               ' ̧�� ��߯��\����
    ks0 As BtKeySpeck               ' �� ��߯��\����
    ks1 As BtKeySpeck
    ks2 As BtKeySpeck
    ks3 As BtKeySpeck
    ks4 As BtKeySpeck
    ks5 As BtKeySpeck
    ks6 As BtKeySpeck
    ks7 As BtKeySpeck
    ks8 As BtKeySpeck
    ks9 As BtKeySpeck
    ks10 As BtKeySpeck
    ks11 As BtKeySpeck
    ks12 As BtKeySpeck
    ks13 As BtKeySpeck
    ks14 As BtKeySpeck
    ks15 As BtKeySpeck
    ks16 As BtKeySpeck
    ks17 As BtKeySpeck
    ks18 As BtKeySpeck
    ks19 As BtKeySpeck
    ks20 As BtKeySpeck
    ks21 As BtKeySpeck
    ks22 As BtKeySpeck
    ks23 As BtKeySpeck
    ks24 As BtKeySpeck
    ks25 As BtKeySpeck
    ks26 As BtKeySpeck
    ks27 As BtKeySpeck
    ks28 As BtKeySpeck
    ks29 As BtKeySpeck
    ks30 As BtKeySpeck
End Type

Global Y_SYU_Speck As Y_SYU_FSpeck

Private Function Y_SYU_Create() As Integer
'********************************************************************
'*                                                                  *
'*              �o�ח\��f�[�^  �b�q�d�`�s�d                        *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'*          CREATE 1997.06.02  S.Shibano                            *
'********************************************************************
Dim sts As Integer
Dim FullPath As String
Dim c As String * 128

    Y_SYU_Create = False
                                            '�o�ח\��f�[�^�t���p�X�捞��
    sts = GetIni("FILE", Y_SYU_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        Y_SYU_Create = True
        Exit Function
    End If
    
    FullPath = RTrim$(c)

    Y_SYU_Speck.fs.recoleng = Len(Y_SYUREC)         ' ���R�[�h��
    Y_SYU_Speck.fs.PageSize = Y_SYU_PG_SIZ          ' �y�[�W�T�C�Y
    Y_SYU_Speck.fs.idexnumb = 6                     ' �C���f�b�N�X��
    Y_SYU_Speck.fs.fileflag = 0                     ' �t�@�C���t���O
    Y_SYU_Speck.fs.reserve = &H0                    ' �\��ς�
                                                    ' �L�[�O
    Y_SYU_Speck.ks0.keypos = 20                     ' �L�[�|�W�V����
    Y_SYU_Speck.ks0.keyleng = 1                     ' �L�[��
    Y_SYU_Speck.ks0.keyflag = BtKfExt + BtKfSeg     ' �L�[�t���O
    Y_SYU_Speck.ks0.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    Y_SYU_Speck.ks0.reserve = &H0                   ' �\��ς�
                                                    ' �L�[�O
    Y_SYU_Speck.ks1.keypos = 31                     ' �L�[�|�W�V����
    Y_SYU_Speck.ks1.keyleng = 8                     ' �L�[��
    Y_SYU_Speck.ks1.keyflag = BtKfExt + BtKfSeg     ' �L�[�t���O
    Y_SYU_Speck.ks1.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    Y_SYU_Speck.ks1.reserve = &H0                   ' �\��ς�
                                                    ' �L�[�O
    Y_SYU_Speck.ks2.keypos = 42                     ' �L�[�|�W�V����
    Y_SYU_Speck.ks2.keyleng = 6                     ' �L�[��
    Y_SYU_Speck.ks2.keyflag = BtKfExt + BtKfSeg     ' �L�[�t���O
    Y_SYU_Speck.ks2.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    Y_SYU_Speck.ks2.reserve = &H0                   ' �\��ς�
                                                    ' �L�[�O
    Y_SYU_Speck.ks3.keypos = 156                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks3.keyleng = 2                     ' �L�[��
    Y_SYU_Speck.ks3.keyflag = BtKfExt               ' �L�[�t���O
    Y_SYU_Speck.ks3.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    Y_SYU_Speck.ks3.reserve = &H0                   ' �\��ς�
                                                    ' �L�[�P
    Y_SYU_Speck.ks4.keypos = 20                     ' �L�[�|�W�V����
    Y_SYU_Speck.ks4.keyleng = 1                     ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks4.keyflag = BtKfExt + BtKfSeg + BtKfChg
    Y_SYU_Speck.ks4.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    Y_SYU_Speck.ks4.reserve = &H0                   ' �\��ς�
                                                    ' �L�[�P
    Y_SYU_Speck.ks5.keypos = 1                      ' �L�[�|�W�V����
    Y_SYU_Speck.ks5.keyleng = 1                     ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks5.keyflag = BtKfExt + BtKfSeg + BtKfChg
    Y_SYU_Speck.ks5.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    Y_SYU_Speck.ks5.reserve = &H0                   ' �\��ς�
                                                    ' �L�[�P
    Y_SYU_Speck.ks6.keypos = 48                     ' �L�[�|�W�V����
    Y_SYU_Speck.ks6.keyleng = 1                     ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks6.keyflag = BtKfExt + BtKfSeg + BtKfChg
    Y_SYU_Speck.ks6.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    Y_SYU_Speck.ks6.reserve = &H0                   ' �\��ς�
                                                    ' �L�[�P
    Y_SYU_Speck.ks7.keypos = 120                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks7.keyleng = 5                     ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks7.keyflag = BtKfExt + BtKfSeg + BtKfChg
    Y_SYU_Speck.ks7.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    Y_SYU_Speck.ks7.reserve = &H0                   ' �\��ς�
                                                    ' �L�[�P
    Y_SYU_Speck.ks8.keypos = 156                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks8.keyleng = 2                     ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks8.keyflag = BtKfExt + BtKfSeg + BtKfChg
    Y_SYU_Speck.ks8.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    Y_SYU_Speck.ks8.reserve = &H0                   ' �\��ς�
                                                    ' �L�[�P
    Y_SYU_Speck.ks9.keypos = 49                     ' �L�[�|�W�V����
    Y_SYU_Speck.ks9.keyleng = 13                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks9.keyflag = BtKfExt + BtKfSeg + BtKfChg
    Y_SYU_Speck.ks9.keytype = Chr(BtKtString)       ' �L�[�^�C�v
    Y_SYU_Speck.ks9.reserve = &H0                   ' �\��ς�
                                                    ' �L�[�P
    Y_SYU_Speck.ks10.keypos = 31                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks10.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks10.keyflag = BtKfExt + BtKfSeg + BtKfChg
    Y_SYU_Speck.ks10.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks10.reserve = &H0                  ' �\��ς�
                                                    ' �L�[�P
    Y_SYU_Speck.ks11.keypos = 42                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks11.keyleng = 6                    ' �L�[��
    Y_SYU_Speck.ks11.keyflag = BtKfExt + BtKfChg    ' �L�[�t���O
    Y_SYU_Speck.ks11.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks11.reserve = &H0                  ' �\��ς�
                                                    ' �L�[�Q
    Y_SYU_Speck.ks12.keypos = 20                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks12.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks12.keyflag = BtKfExt + BtKfSeg + BtKfChg + BtKfDup
    Y_SYU_Speck.ks12.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks12.reserve = &H0                  ' �\��ς�
                                                    ' �L�[�Q
    Y_SYU_Speck.ks13.keypos = 1                     ' �L�[�|�W�V����
    Y_SYU_Speck.ks13.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks13.keyflag = BtKfExt + BtKfSeg + BtKfChg + BtKfDup
    Y_SYU_Speck.ks13.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks13.reserve = &H0                  ' �\��ς�
                                                    ' �L�[�Q
    Y_SYU_Speck.ks14.keypos = 48                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks14.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks14.keyflag = BtKfExt + BtKfSeg + BtKfChg + BtKfDup
    Y_SYU_Speck.ks14.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks14.reserve = &H0                  ' �\��ς�
                                                    ' �L�[�Q
    Y_SYU_Speck.ks15.keypos = 112                   ' �L�[�|�W�V����
    Y_SYU_Speck.ks15.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks15.keyflag = BtKfExt + BtKfSeg + BtKfChg + BtKfDup
    Y_SYU_Speck.ks15.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks15.reserve = &H0                  ' �\��ς�
                                                    ' �L�[�Q
    Y_SYU_Speck.ks16.keypos = 31                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks16.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks16.keyflag = BtKfExt + BtKfChg + BtKfDup
    Y_SYU_Speck.ks16.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks16.reserve = &H0                  ' �\��ς�
                                                    ' �L�[�R
    Y_SYU_Speck.ks17.keypos = 20                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks17.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks17.keyflag = BtKfExt + BtKfSeg + BtKfChg
    Y_SYU_Speck.ks17.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks17.reserve = &H0                  ' �\��ς�
                                                    ' �L�[�R
    Y_SYU_Speck.ks18.keypos = 1                     ' �L�[�|�W�V����
    Y_SYU_Speck.ks18.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks18.keyflag = BtKfExt + BtKfSeg + BtKfChg
    Y_SYU_Speck.ks18.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks18.reserve = &H0                  ' �\��ς�
                                                    ' �L�[�R
    Y_SYU_Speck.ks19.keypos = 120                   ' �L�[�|�W�V����
    Y_SYU_Speck.ks19.keyleng = 5                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks19.keyflag = BtKfExt + BtKfSeg + BtKfChg
    Y_SYU_Speck.ks19.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks19.reserve = &H0                  ' �\��ς�
                                                    ' �L�[�R
    Y_SYU_Speck.ks20.keypos = 48                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks20.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks20.keyflag = BtKfExt + BtKfSeg + BtKfChg
    Y_SYU_Speck.ks20.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks20.reserve = &H0                  ' �\��ς�
                                                    ' �L�[�R
    Y_SYU_Speck.ks21.keypos = 49                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks21.keyleng = 13                   ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks21.keyflag = BtKfExt + BtKfSeg + BtKfChg
    Y_SYU_Speck.ks21.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks21.reserve = &H0                  ' �\��ς�
                                                    ' �L�[�R
    Y_SYU_Speck.ks22.keypos = 31                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks22.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks22.keyflag = BtKfExt + BtKfSeg + BtKfChg
    Y_SYU_Speck.ks22.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks22.reserve = &H0                  ' �\��ς�
                                                    ' �L�[�R
    Y_SYU_Speck.ks23.keypos = 42                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks23.keyleng = 6                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks23.keyflag = BtKfExt + BtKfSeg + BtKfChg
    Y_SYU_Speck.ks23.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks23.reserve = &H0                  ' �\��ς�
                                                    ' �L�[�R
    Y_SYU_Speck.ks24.keypos = 156                   ' �L�[�|�W�V����
    Y_SYU_Speck.ks24.keyleng = 2                    ' �L�[��
    Y_SYU_Speck.ks24.keyflag = BtKfExt + BtKfChg    ' �L�[�t���O
    Y_SYU_Speck.ks24.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks24.reserve = &H0                  ' �\��ς�
                                                    ' �L�[�S
    Y_SYU_Speck.ks25.keypos = 20                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks25.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks25.keyflag = BtKfExt + BtKfSeg + BtKfChg + BtKfDup
    Y_SYU_Speck.ks25.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks25.reserve = &H0                  ' �\��ς�
                                                    ' �L�[�S
    Y_SYU_Speck.ks26.keypos = 1                     ' �L�[�|�W�V����
    Y_SYU_Speck.ks26.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks26.keyflag = BtKfExt + BtKfSeg + BtKfChg + BtKfDup
    Y_SYU_Speck.ks26.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks26.reserve = &H0                  ' �\��ς�
                                                    ' �L�[�S
    Y_SYU_Speck.ks27.keypos = 19                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks27.keyleng = 1                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks27.keyflag = BtKfExt + BtKfSeg + BtKfChg + BtKfDup
    Y_SYU_Speck.ks27.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks27.reserve = &H0                  ' �\��ς�
                                                    ' �L�[�S
    Y_SYU_Speck.ks28.keypos = 49                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks28.keyleng = 13                   ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks28.keyflag = BtKfExt + BtKfSeg + BtKfChg + BtKfDup
    Y_SYU_Speck.ks28.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks28.reserve = &H0                  ' �\��ς�
                                                    ' �L�[�S
    Y_SYU_Speck.ks29.keypos = 31                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks29.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks29.keyflag = BtKfExt + BtKfChg + BtKfDup
    Y_SYU_Speck.ks29.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks29.reserve = &H0                  ' �\��ς�
                                                    ' �L�[�T
    Y_SYU_Speck.ks30.keypos = 31                    ' �L�[�|�W�V����
    Y_SYU_Speck.ks30.keyleng = 8                    ' �L�[��
                                                    ' �L�[�t���O
    Y_SYU_Speck.ks30.keyflag = BtKfExt + BtKfChg + BtKfDup
    Y_SYU_Speck.ks30.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    Y_SYU_Speck.ks30.reserve = &H0                  ' �\��ς�

    sts = BTRV(BtOpCreate, Y_SYU_POS, Y_SYU_Speck, Len(Y_SYU_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�o�ח\��f�[�^")
        Y_SYU_Create = True
    End If
End Function

Function Y_SYU_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              �o�ח\��f�[�^  �n�o�d�m                            �@*
'*                                                                  *
'*      ��  ��:Open Mode(Btrieve�Q��)                               *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'*          CREATE 1997.06.02  S.Shibano                            *
'********************************************************************
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String
    
        Y_SYU_Open = False
                                            '�o�ח\��f�[�^�t���p�X�捞��
    sts = GetIni("FILE", Y_SYU_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        Y_SYU_Open = True
        Exit Function
    End If
    FullPath = RTrim$(c)
    
    Do
        sts = BTRV(BtOpOpen, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Function
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = Y_SYU_Create()        '�o�ח\��f�[�^�쐬
                If sts <> False Then
                    Y_SYU_Open = True
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�o�ח\��f�[�^")
                    Y_SYU_Open = True
                End If
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "�o�ח\��f�[�^")
                Y_SYU_Open = True
                Exit Function
        End Select
    Loop
End Function


