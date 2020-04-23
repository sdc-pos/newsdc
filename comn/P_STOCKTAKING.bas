Attribute VB_Name = "P_STOCKTAKING"
Option Explicit

'********************************************************************
'*
'*              ���ޒI�����ް�  �t�@�C����`
'*
'*          CREATE 2006.02.15
'********************************************************************
'�t�@�C���h�c
Public Const P_STOCK_ID$ = "P_STOCK"

'�y�[�W�T�C�Y
Private Const P_STOCK_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public P_STOCK_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Public Type P_STOCK_REC_Tag
    
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '���ޕi��
    
    
    CODE(0 To 4)            As Byte         '�d���溰��
    TANKA(0 To 10)          As Byte         '�d���P�� 9(8)V99
    
    INPUT_DATE(0 To 7)      As Byte         '�o�^���t 2006.11.22
    
    G_SYUSHI(0 To 2)        As Byte         '���x�P��
    ZEN_ZAIKO_QTY(0 To 7)   As Byte         '�O���݌ɐ���
                            
    NYUKO_QTY(0 To 7)       As Byte         '���ɐ�
    SYUKO_QTY(0 To 7)       As Byte         '�o�ɐ�
    ZAIKO_QTY(0 To 7)       As Byte         '�݌ɐ�
    
    
    LAST_SYUKA_DT(0 To 7)   As Byte         '�ŏI�o�ד�
    LAST_SYUKA_QTY(0 To 7)  As Byte         '�ŏI�o�א���
    
    MOTO_ZAIKO_QTY(0 To 7)  As Byte         '�ďW�v�O
    MAEGARI_QTY(0 To 7)     As Byte         '�O�ؐ�
    
    SYUKA_NON_F(0 To 0)     As Byte         '�o�א��v�Z�L���@0:���Ȃ��@1:����
    
    
    
    ZEN_ZAIKO_KIN(0 To 7)   As Byte         '�O���݌ɐ���
    
    
    FILLER(0 To 5)         As Byte          '


End Type
'�f�[�^�E�o�b�t�@
Public P_STOCK_REC          As P_STOCK_REC_Tag

'�L�[��`
    
Public Type KEY0_P_STOCK                    '�j�d�x�O
    
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '���ޕi��
    CODE(0 To 4)            As Byte         '�d���溰��
    TANKA(0 To 10)          As Byte         '�d���P�� 9(8)V99
    
End Type
    
Public Type KEY1_P_STOCK                    '�j�d�x�P�@2006.11.22
    
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '���ޕi��
    
    INPUT_DATE(0 To 7)      As Byte         '�o�^���t 2006.11.22
    
    
    CODE(0 To 4)            As Byte         '�d���溰��
    TANKA(0 To 10)          As Byte         '�d���P�� 9(8)V99
    
End Type
    
    
    
    
    
'�L�[�E�f�[�^
Public K0_P_STOCK           As KEY0_P_STOCK
Public K1_P_STOCK           As KEY1_P_STOCK


Type P_STOCK_FSpeck
    fs                      As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                     As BtKeySpeck   ' �� ��߯��\����
    ks1                     As BtKeySpeck   ' �� ��߯��\����
    ks2                     As BtKeySpeck   ' �� ��߯��\����
    ks3                     As BtKeySpeck   ' �� ��߯��\����
    ks4                     As BtKeySpeck   ' �� ��߯��\����
    ks5                     As BtKeySpeck   ' �� ��߯��\����
    ks6                     As BtKeySpeck   ' �� ��߯��\����
    ks7                     As BtKeySpeck   ' �� ��߯��\����
    ks8                     As BtKeySpeck   ' �� ��߯��\����
    ks9                     As BtKeySpeck   ' �� ��߯��\����
    ks10                    As BtKeySpeck   ' �� ��߯��\����

End Type

Private P_STOCK_Speck       As P_STOCK_FSpeck
Private Function P_STOCK_Create() As Integer
'********************************************************************
'*
'*              ���ޒI�����ް�  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'*      ���x���Ƀt�@�C�����𕪂���  2007.11.13
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

Dim ret             As Long     '2007.11.13




    P_STOCK_Create = True
                                            '���ޒI�����ް��t���p�X�捞��
    sts = GetIni("FILE", P_STOCK_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_STOCK]�ǂݍ��݃G���[")
        Exit Function
    End If


    '2007.11.13
'    FullPath = Trim(c)
    ret = InStr(1, Trim(c), ".") - 1
    FullPath = Left(Trim(c), ret) & GLB_SYUSHI_F & Right(Trim(c), Len(Trim(c)) - ret)
    '2007.11.13



    P_STOCK_Speck.fs.recoleng = Len(P_STOCK_REC)        ' ���R�[�h��
    P_STOCK_Speck.fs.PageSize = P_STOCK_PG_SIZ          ' �y�[�W�T�C�Y
    P_STOCK_Speck.fs.idexnumb = 2                       ' �C���f�b�N�X��
    P_STOCK_Speck.fs.fileflag = 0                       ' �t�@�C���t���O
    P_STOCK_Speck.fs.reserve = &H0                      ' �\��ς�
    
    '--------------------------------------------------- �L�[�O ��
    P_STOCK_Speck.ks0.keypos = 1                        ' �L�[�|�W�V����
    P_STOCK_Speck.ks0.keyleng = 1                       ' �L�[��
    P_STOCK_Speck.ks0.keyflag = BtKfExt + BtKfSeg + BtKfChg      ' �L�[�t���O
    P_STOCK_Speck.ks0.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_STOCK_Speck.ks0.reserve = &H0                     ' �\��ς�
    
    P_STOCK_Speck.ks1.keypos = 2                        ' �L�[�|�W�V����
    P_STOCK_Speck.ks1.keyleng = 1                       ' �L�[��
    P_STOCK_Speck.ks1.keyflag = BtKfExt + BtKfSeg + BtKfChg      ' �L�[�t���O
    P_STOCK_Speck.ks1.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_STOCK_Speck.ks1.reserve = &H0                     ' �\��ς�
    
    P_STOCK_Speck.ks2.keypos = 3                        ' �L�[�|�W�V����
    P_STOCK_Speck.ks2.keyleng = 20                      ' �L�[��
    P_STOCK_Speck.ks2.keyflag = BtKfExt + BtKfSeg + BtKfChg      ' �L�[�t���O
    P_STOCK_Speck.ks2.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_STOCK_Speck.ks2.reserve = &H0                     ' �\��ς�
    
    P_STOCK_Speck.ks3.keypos = 23                       ' �L�[�|�W�V����
    P_STOCK_Speck.ks3.keyleng = 5                       ' �L�[��
    P_STOCK_Speck.ks3.keyflag = BtKfExt + BtKfSeg + BtKfChg      ' �L�[�t���O
    P_STOCK_Speck.ks3.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_STOCK_Speck.ks3.reserve = &H0                     ' �\��ς�
    
    P_STOCK_Speck.ks4.keypos = 28                       ' �L�[�|�W�V����
    P_STOCK_Speck.ks4.keyleng = 11                      ' �L�[��
    P_STOCK_Speck.ks4.keyflag = BtKfExt + BtKfChg                ' �L�[�t���O
    P_STOCK_Speck.ks4.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_STOCK_Speck.ks4.reserve = &H0                     ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    
    '--------------------------------------------------- �L�[�P ��
    P_STOCK_Speck.ks5.keypos = 1                        ' �L�[�|�W�V����
    P_STOCK_Speck.ks5.keyleng = 1                       ' �L�[��
                                                        ' �L�[�t���O
    P_STOCK_Speck.ks5.keyflag = BtKfExt + BtKfSeg + BtKfChg
    P_STOCK_Speck.ks5.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_STOCK_Speck.ks5.reserve = &H0                     ' �\��ς�
    
    
    P_STOCK_Speck.ks6.keypos = 2                        ' �L�[�|�W�V����
    P_STOCK_Speck.ks6.keyleng = 1                       ' �L�[��
                                                        ' �L�[�t���O
    P_STOCK_Speck.ks6.keyflag = BtKfExt + BtKfSeg + BtKfChg
    P_STOCK_Speck.ks6.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_STOCK_Speck.ks6.reserve = &H0                     ' �\��ς�
    
    P_STOCK_Speck.ks7.keypos = 3                        ' �L�[�|�W�V����
    P_STOCK_Speck.ks7.keyleng = 20                      ' �L�[��
                                                        ' �L�[�t���O
    P_STOCK_Speck.ks7.keyflag = BtKfExt + BtKfSeg + BtKfChg
    P_STOCK_Speck.ks7.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_STOCK_Speck.ks7.reserve = &H0                     ' �\��ς�
    
    P_STOCK_Speck.ks8.keypos = 39                       ' �L�[�|�W�V����
    P_STOCK_Speck.ks8.keyleng = 8                       ' �L�[��
                                                        ' �L�[�t���O
    P_STOCK_Speck.ks8.keyflag = BtKfExt + BtKfSeg + BtKfChg
    P_STOCK_Speck.ks8.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_STOCK_Speck.ks8.reserve = &H0                     ' �\��ς�
    
    P_STOCK_Speck.ks9.keypos = 23                       ' �L�[�|�W�V����
    P_STOCK_Speck.ks9.keyleng = 5                       ' �L�[��
                                                        ' �L�[�t���O
    P_STOCK_Speck.ks9.keyflag = BtKfExt + BtKfSeg + BtKfChg
    P_STOCK_Speck.ks9.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    P_STOCK_Speck.ks9.reserve = &H0                     ' �\��ς�
    
    P_STOCK_Speck.ks10.keypos = 28                      ' �L�[�|�W�V����
    P_STOCK_Speck.ks10.keyleng = 11                     ' �L�[��
    P_STOCK_Speck.ks10.keyflag = BtKfExt + BtKfChg      ' �L�[�t���O
    P_STOCK_Speck.ks10.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    P_STOCK_Speck.ks10.reserve = &H0                     ' �\��ς�
    
    
    
    
    
    
    
    
    sts = BTRV(BtOpCreate, P_STOCK_POS, P_STOCK_Speck, Len(P_STOCK_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���ޒI�����ް�")
        Exit Function
    End If
    
    P_STOCK_Create = False

End Function

Public Function P_STOCK_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���ޒI�����ް�  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*      ���x���Ƀt�@�C�����𕪂���  2007.11.13
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

Dim ret             As Long     '2007.11.13


    P_STOCK_Open = True
                                            '���ޒI���f�[�^�t���p�X�捞��
    sts = GetIni("FILE", P_STOCK_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_STOCK]�ǂݍ��݃G���[")
        Exit Function
    End If

    '2007.11.13
'    FullPath = Trim(c)
    ret = InStr(1, Trim(c), ".") - 1
    FullPath = Left(Trim(c), ret) & GLB_SYUSHI_F & Right(Trim(c), Len(Trim(c)) - ret)
    '2007.11.13


    Do
        sts = BTRV(BtOpOpen, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_STOCK_Create()   '���ޒI�����ް��쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���ޒI�����ް�")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���ޒI�����ް�")
                Exit Function
        End Select
    Loop
    
    P_STOCK_Open = False

End Function

