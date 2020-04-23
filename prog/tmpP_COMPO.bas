Attribute VB_Name = "tmpP_COMPO"
Option Explicit
'********************************************************************
'*                                                                  *
'*              �\���}�X�^  �t�@�C����`                            *
'*                                                                  *
'*          CREATE 2005.11.11                                       *
'********************************************************************
'�t�@�C���h�c
Public Const tmpP_COMPO_ID$ = "tmpP_COMPO"

'�y�[�W�T�C�Y
Private Const tmpP_COMPO_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public tmpP_COMPO_POS      As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Public Type tmpP_COMPOREC_Tag
    
    
    SHIMUKE(0 To 2)         As Byte         '�d������
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '�e�i��
    DATA_KBN(0 To 0)        As Byte         '�ް��敪
    SEQNO(0 To 2)           As Byte         '�ǔ�
    KO_SYUBETSU(0 To 1)     As Byte         '�q�@���
    KO_JGYOBU(0 To 0)       As Byte         '�q�@���ƕ�
    KO_NAIGAI(0 To 0)       As Byte         '�q�@�����O
    KO_HIN_GAI(0 To 19)     As Byte         '�q�@�i��
    KO_QTY(0 To 5)          As Byte         '�q�@����(999V99)
    KO_BIKOU(0 To 39)       As Byte         '�q�@���l
    FILLER(0 To 137)        As Byte         'Filler
    UPD_TANTO(0 To 4)       As Byte         '�X�V�@�S����
    UPD_DATETIME(0 To 13)   As Byte         '�X�V�@����

End Type
'�f�[�^�E�o�b�t�@
Public tmpP_COMPOREC         As tmpP_COMPOREC_Tag

'�L�[��`

Type KEY0_tmpP_COMPO                        '�j�d�x�O
    SHIMUKE(0 To 2)         As Byte         '�d������
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '�e�i��
    DATA_KBN(0 To 0)        As Byte         '�ް��敪
    SEQNO(0 To 2)           As Byte         '�ǔ�
End Type
    
'�L�[�E�f�[�^
Public K0_tmpP_COMPO        As KEY0_tmpP_COMPO

Type tmpP_COMPO_FSpeck
    fs                      As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                     As BtKeySpeck   ' �� ��߯��\����
    ks1                     As BtKeySpeck   ' �� ��߯��\����
    ks2                     As BtKeySpeck   ' �� ��߯��\����
    ks3                     As BtKeySpeck   ' �� ��߯��\����
    ks4                     As BtKeySpeck   ' �� ��߯��\����
    ks5                     As BtKeySpeck   ' �� ��߯��\����
End Type

Private tmpP_COMPO_Speck    As tmpP_COMPO_FSpeck

Private Function tmpP_COMPO_Create() As Integer
'********************************************************************
'*                                                                  *
'*              �\���}�X�^(�ꎞ�t�@�C��)�b�q�d�`�s�d                *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    tmpP_COMPO_Create = True
                                            '�\���}�X�^�t���p�X�捞��
    sts = GetIni("FILE", tmpP_COMPO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [tmpP_COMPO]�ǂݍ��݃G���[")
        Exit Function
    End If

    FullPath = RTrim(c)

    tmpP_COMPO_Speck.fs.recoleng = Len(tmpP_COMPOREC)       ' ���R�[�h��
    tmpP_COMPO_Speck.fs.PageSize = tmpP_COMPO_PG_SIZ        ' �y�[�W�T�C�Y
    tmpP_COMPO_Speck.fs.idexnumb = 1                        ' �C���f�b�N�X��
    tmpP_COMPO_Speck.fs.fileflag = 0                        ' �t�@�C���t���O
    tmpP_COMPO_Speck.fs.reserve = &H0                       ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    tmpP_COMPO_Speck.ks0.keypos = 1                         ' �L�[�|�W�V����
    tmpP_COMPO_Speck.ks0.keyleng = 3                        ' �L�[��
    tmpP_COMPO_Speck.ks0.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
    tmpP_COMPO_Speck.ks0.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    tmpP_COMPO_Speck.ks0.reserve = &H0                      ' �\��ς�
    
    tmpP_COMPO_Speck.ks1.keypos = 4                         ' �L�[�|�W�V����
    tmpP_COMPO_Speck.ks1.keyleng = 1                        ' �L�[��
    tmpP_COMPO_Speck.ks1.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
    tmpP_COMPO_Speck.ks1.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    tmpP_COMPO_Speck.ks1.reserve = &H0                      ' �\��ς�
    
    tmpP_COMPO_Speck.ks2.keypos = 5                         ' �L�[�|�W�V����
    tmpP_COMPO_Speck.ks2.keyleng = 1                        ' �L�[��
    tmpP_COMPO_Speck.ks2.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
    tmpP_COMPO_Speck.ks2.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    tmpP_COMPO_Speck.ks2.reserve = &H0                      ' �\��ς�
    
    tmpP_COMPO_Speck.ks3.keypos = 6                         ' �L�[�|�W�V����
    tmpP_COMPO_Speck.ks3.keyleng = 20                       ' �L�[��
    tmpP_COMPO_Speck.ks3.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
    tmpP_COMPO_Speck.ks3.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    tmpP_COMPO_Speck.ks3.reserve = &H0                      ' �\��ς�
    
    tmpP_COMPO_Speck.ks4.keypos = 26                        ' �L�[�|�W�V����
    tmpP_COMPO_Speck.ks4.keyleng = 1                        ' �L�[��
    tmpP_COMPO_Speck.ks4.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
    tmpP_COMPO_Speck.ks4.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    tmpP_COMPO_Speck.ks4.reserve = &H0                      ' �\��ς�
    
    tmpP_COMPO_Speck.ks5.keypos = 27                        ' �L�[�|�W�V����
    tmpP_COMPO_Speck.ks5.keyleng = 3                        ' �L�[��
    tmpP_COMPO_Speck.ks5.keyflag = BtKfExt                  ' �L�[�t���O
    tmpP_COMPO_Speck.ks5.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    tmpP_COMPO_Speck.ks5.reserve = &H0                      ' �\��ς�
    
    '--------------------------------------------------- �L�[�O ��
    sts = BTRV(BtOpCreate, tmpP_COMPO_POS, tmpP_COMPO_Speck, Len(tmpP_COMPO_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�\���}�X�^�i�ꎞ�t�@�C���j")
        Exit Function
    End If
    
    tmpP_COMPO_Create = False

End Function

Public Function tmpP_COMPO_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �\���}�X�^�i�ꎞ�t�@�C���j  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
Dim ans         As Integer


    tmpP_COMPO_Open = True
                                            '�\���}�X�^�t���p�X�捞��
    sts = GetIni("FILE", tmpP_COMPO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [tmpP_COMPO]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, tmpP_COMPO_POS, tmpP_COMPOREC, Len(tmpP_COMPOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                
                ans = MsgBox("���[���ŁA��Ɨp�t�@�C���g�p���ł��B", vbRetryCancel, "�m�F����")
                
                If ans = vbCancel Then
                    Exit Function
                End If
                Sleep (500&)
            Case BtErrFileNotFound
                sts = tmpP_COMPO_Create()      '�\���}�X�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, tmpP_COMPO_POS, tmpP_COMPOREC, Len(tmpP_COMPOREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�\���}�X�^�i�ꎞ�t�@�C���j")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�\���}�X�^�i�ꎞ�t�@�C���j")
                Exit Function
        End Select
    Loop
    
    tmpP_COMPO_Open = False

End Function

