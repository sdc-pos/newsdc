Attribute VB_Name = "P_SHURIAGE"
Option Explicit

'********************************************************************
'*
'*              ���ޔ����ް�  �t�@�C����`
'*
'*          CREATE 2005.11.11
'********************************************************************
'�t�@�C���h�c
Public Const P_SHURIAGE_ID$ = "P_SHURIAGE"

'�y�[�W�T�C�Y
Private Const P_SHURIAGE_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public P_SHURIAGE_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Public Type P_SHURIAGE_REC_Tag
    
    URIAGE_NO(0 To 4)       As Byte         'ں��އ�
    URIAGE_DT(0 To 7)       As Byte         '����N����
    KEIJYO_YM(0 To 5)       As Byte         '�v��N��
    TORI_KBN(0 To 0)        As Byte         '�����敪
    TOKUI_CODE(0 To 4)      As Byte         '���Ӑ溰��
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '���ޕi��
    G_SYUSHI(0 To 2)        As Byte         '���x�P��
    G_HANBAI_KBN(0 To 1)    As Byte         '�̔��敪
    URIAGE_QTY(0 To 11)     As Byte         '���㐔��(S9(8)V99)
    TANKA(0 To 10)          As Byte         '�P��(9(8)V99)
    KINGAKU(0 To 8)         As Byte         '������z(S9(8))
    SEIKU_F(0 To 0)         As Byte         '�����׸�
        
    ZEI_KIN(0 To 8)         As Byte         '�����(S9(8))
    
    FILLER(0 To 19)         As Byte         'Filler
    UPD_DATETIME(0 To 13)   As Byte         '�X�V�@����

End Type
'�f�[�^�E�o�b�t�@
Public P_SHURIAGE_REC       As P_SHURIAGE_REC_Tag

'�L�[��`
Public Type KEY0_P_SHURIAGE                 '�j�d�x�O
    URIAGE_NO(0 To 4)       As Byte         'ں��އ�
End Type
    
Public Type KEY1_P_SHURIAGE                 '�j�d�x�P
    KEIJYO_YM(0 To 5)       As Byte         '�v��N��
    G_SYUSHI(0 To 2)        As Byte         '���x�P��
    TOKUI_CODE(0 To 4)      As Byte         '���Ӑ溰��
    URIAGE_DT(0 To 7)       As Byte         '����N����
    URIAGE_NO(0 To 4)       As Byte         'ں��އ�
End Type
    
    
'�L�[�E�f�[�^
Public K0_P_SHURIAGE        As KEY0_P_SHURIAGE
Public K1_P_SHURIAGE        As KEY1_P_SHURIAGE

Type P_SHURIAGE_FSpeck
    fs                      As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                     As BtKeySpeck   ' �� ��߯��\����
    ks1                     As BtKeySpeck   ' �� ��߯��\����
    ks2                     As BtKeySpeck   ' �� ��߯��\����
    ks3                     As BtKeySpeck   ' �� ��߯��\����
    ks4                     As BtKeySpeck   ' �� ��߯��\����
    ks5                     As BtKeySpeck   ' �� ��߯��\����
End Type

Private P_SHURIAGE_Speck    As P_SHURIAGE_FSpeck
Private Function P_SHURIAGE_Create() As Integer
'********************************************************************
'*
'*              ���ޔ����ް�  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    P_SHURIAGE_Create = True
                                            '���ޔ����ް��t���p�X�捞��
    sts = GetIni("FILE", P_SHURIAGE_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SHURIAGE]�ǂݍ��݃G���[")
        Exit Function
    End If

    FullPath = RTrim(c)

    P_SHURIAGE_Speck.fs.recoleng = Len(P_SHURIAGE_REC)  ' ���R�[�h��
    P_SHURIAGE_Speck.fs.PageSize = P_SHURIAGE_PG_SIZ    ' �y�[�W�T�C�Y
    P_SHURIAGE_Speck.fs.idexnumb = 2                    ' �C���f�b�N�X��
    P_SHURIAGE_Speck.fs.fileflag = 0                    ' �t�@�C���t���O
    P_SHURIAGE_Speck.fs.reserve = &H0                   ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    P_SHURIAGE_Speck.ks0.keypos = 1                     ' �L�[�|�W�V����
    P_SHURIAGE_Speck.ks0.keyleng = 5                    ' �L�[��
    P_SHURIAGE_Speck.ks0.keyflag = BtKfExt              ' �L�[�t���O
    P_SHURIAGE_Speck.ks0.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHURIAGE_Speck.ks0.reserve = &H0                  ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    
    '--------------------------------------------------- �L�[�P ��
    P_SHURIAGE_Speck.ks1.keypos = 14                    ' �L�[�|�W�V����
    P_SHURIAGE_Speck.ks1.keyleng = 6                    ' �L�[��
    P_SHURIAGE_Speck.ks1.keyflag = BtKfExt + BtKfChg + BtKfSeg  ' �L�[�t���O
    P_SHURIAGE_Speck.ks1.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHURIAGE_Speck.ks1.reserve = &H0                  ' �\��ς�
    
    P_SHURIAGE_Speck.ks2.keypos = 48                    ' �L�[�|�W�V����
    P_SHURIAGE_Speck.ks2.keyleng = 3                    ' �L�[��
    P_SHURIAGE_Speck.ks2.keyflag = BtKfExt + BtKfChg + BtKfSeg    ' �L�[�t���O
    P_SHURIAGE_Speck.ks2.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHURIAGE_Speck.ks2.reserve = &H0                  ' �\��ς�
    
    P_SHURIAGE_Speck.ks3.keypos = 21                    ' �L�[�|�W�V����
    P_SHURIAGE_Speck.ks3.keyleng = 5                    ' �L�[��
    P_SHURIAGE_Speck.ks3.keyflag = BtKfExt + BtKfChg + BtKfSeg   ' �L�[�t���O
    P_SHURIAGE_Speck.ks3.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHURIAGE_Speck.ks3.reserve = &H0                  ' �\��ς�
    
    P_SHURIAGE_Speck.ks4.keypos = 6                    ' �L�[�|�W�V����
    P_SHURIAGE_Speck.ks4.keyleng = 8                    ' �L�[��
    P_SHURIAGE_Speck.ks4.keyflag = BtKfExt + BtKfChg + BtKfSeg   ' �L�[�t���O
    P_SHURIAGE_Speck.ks4.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHURIAGE_Speck.ks4.reserve = &H0                  ' �\��ς�
    
    P_SHURIAGE_Speck.ks5.keypos = 1                     ' �L�[�|�W�V����
    P_SHURIAGE_Speck.ks5.keyleng = 5                    ' �L�[��
    P_SHURIAGE_Speck.ks5.keyflag = BtKfExt + BtKfChg             ' �L�[�t���O
    P_SHURIAGE_Speck.ks5.keytype = Chr(BtKtString)      ' �L�[�^�C�v
    P_SHURIAGE_Speck.ks5.reserve = &H0                  ' �\��ς�
    '--------------------------------------------------- �L�[�P ��
    
    
    sts = BTRV(BtOpCreate, P_SHURIAGE_POS, P_SHURIAGE_Speck, Len(P_SHURIAGE_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���ޔ����ް�")
        Exit Function
    End If
    
    P_SHURIAGE_Create = False

End Function

Public Function P_SHURIAGE_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���ޔ����ް�  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    P_SHURIAGE_Open = True
                                            '���ޔ����ް��t���p�X�捞��
    sts = GetIni("FILE", P_SHURIAGE_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SHURIAGE]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, P_SHURIAGE_POS, P_SHURIAGE_REC, Len(P_SHURIAGE_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_SHURIAGE_Create()   '���ޔ����ް��쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_SHURIAGE_POS, P_SHURIAGE_REC, Len(P_SHURIAGE_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���ޔ����ް�")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���ޔ����ް�")
                Exit Function
        End Select
    Loop
    
    P_SHURIAGE_Open = False

End Function

