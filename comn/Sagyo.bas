Attribute VB_Name = "SAGYO"
Option Explicit
'********************************************************************
'*                                                                  *
'*              ��ƊǗ��}�X�^  �t�@�C����`                        *
'*                                                                  *
'*          CREATE 1997.06.02  S.Shibano                            *
'*          UPDATE 2001.02.14  ��Ɛ������̂̍폜
'*                             ������\�����ڂ̕ύX
'*                             �����O�L���C������L���̍폜
'********************************************************************
'�t�@�C���h�c
Global Const SAGYO_ID = "SAGYO"

'�y�[�W�T�C�Y
Global Const SAGYO_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Global SAGYO_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type SAGYOREC_Tag
    BAR_TYPE(0 To 2)    As Byte     '��o�[�R�[�h�̌n
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    PARM(0 To 3)        As Byte     '�p�����[�^
    SAGYO_DNAME(0 To 15) As Byte    '�\������
    LCD1_TYPE(0 To 0)   As Byte     'LCD1�s�ڐ���
    LCD2_TYPE(0 To 0)   As Byte     'LCD2�s�ڐ���
    LCD3_TYPE(0 To 0)   As Byte     'LCD3�s�ڐ���
    LCD4_TYPE(0 To 0)   As Byte     'LCD4�s�ڐ���
    LCD2_DSP(0 To 15)   As Byte     'LCD2�s�ڕ\�����e
    LCD3_DSP(0 To 15)   As Byte     'LCD3�s�ڕ\�����e
    LCD4_DSP(0 To 15)   As Byte     'LCD4�s�ڕ\�����e
    LOCK_F(0 To 0)      As Byte     '�r���t���O
    FILLER(0 To 3)      As Byte     'FILLER
End Type
'�f�[�^�E�o�b�t�@
Global SAGYOREC As SAGYOREC_Tag

'�L�[��`

Type KEY0_SAGYO            '�j�d�x�O
    BAR_TYPE(0 To 2)    As Byte     '��o�[�R�[�h�̌n
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    PARM(0 To 3)        As Byte     '�p�����[�^
End Type

Type KEY0_SAGY1            '�j�d�x�P
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    BAR_TYPE(0 To 2)    As Byte     '��o�[�R�[�h�̌n
    NAIGAI(0 To 0)      As Byte     '�����O
    PARM(0 To 3)        As Byte     '�p�����[�^
End Type

'�L�[�E�f�[�^
Global K0_SAGYO As KEY0_SAGYO
Global K1_SAGYO As KEY0_SAGY1

Type SAGYO_FSpeck
    fs As BtFileSpeck               ' ̧�� ��߯��\����
    ks0 As BtKeySpeck               ' �� ��߯��\����
    ks1 As BtKeySpeck
    ks2 As BtKeySpeck
    ks3 As BtKeySpeck
End Type

Global SAGYO_Speck As SAGYO_FSpeck
 

Private Function SAGYO_Create() As Integer
'********************************************************************
'*                                                                  *
'*              ��ƊǗ��}�X�^  �b�q�d�`�s�d                        *
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

    SAGYO_Create = False
                                            '��ƊǗ��}�X�^�t���p�X�捞��
    sts = GetIni("FILE", SAGYO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        SAGYO_Create = True
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    SAGYO_Speck.fs.recoleng = Len(SAGYOREC)     ' ���R�[�h��
    SAGYO_Speck.fs.PageSize = SAGYO_PG_SIZ      ' �y�[�W�T�C�Y
    SAGYO_Speck.fs.idexnumb = 2                 ' �C���f�b�N�X��
    SAGYO_Speck.fs.fileflag = 0                 ' �t�@�C���t���O
    SAGYO_Speck.fs.reserve = &H0                ' �\��ς�
                                                ' �L�[�O
    SAGYO_Speck.ks0.keypos = 1                  ' �L�[�|�W�V����
                                                ' �L�[��
    SAGYO_Speck.ks0.keyleng = 3 + 1 + 1 + 4
    SAGYO_Speck.ks0.keyflag = BtKfExt           ' �L�[�t���O
    SAGYO_Speck.ks0.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    SAGYO_Speck.ks0.reserve = &H0               ' �\��ς�
                                                ' �L�[�P
    SAGYO_Speck.ks1.keypos = 4                  ' �L�[�|�W�V����
    SAGYO_Speck.ks1.keyleng = 1                 ' �L�[��
    SAGYO_Speck.ks1.keyflag = BtKfSeg 'BtKfExt           ' �L�[�t���O
    SAGYO_Speck.ks1.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    SAGYO_Speck.ks1.reserve = &H0               ' �\��ς�
    SAGYO_Speck.ks2.keypos = 1                  ' �L�[�|�W�V����
    SAGYO_Speck.ks2.keyleng = 3                 ' �L�[��
    SAGYO_Speck.ks2.keyflag = BtKfSeg 'BtKfExt           ' �L�[�t���O
    SAGYO_Speck.ks2.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    SAGYO_Speck.ks2.reserve = &H0               ' �\��ς�
    SAGYO_Speck.ks3.keypos = 5                  ' �L�[�|�W�V����
    SAGYO_Speck.ks3.keyleng = 5                 ' �L�[��
    SAGYO_Speck.ks3.keyflag = BtKfExt           ' �L�[�t���O
    SAGYO_Speck.ks3.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    SAGYO_Speck.ks3.reserve = &H0               ' �\��ς�
    
    sts = BTRV(BtOpCreate, SAGYO_POS, SAGYO_Speck, Len(SAGYO_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "��ƊǗ��}�X�^")
        SAGYO_Create = True
    End If
End Function

Function SAGYO_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              ��ƊǗ��}�X�^  �n�o�d�m                            *
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
    
    SAGYO_Open = False
                                            '��ƊǗ��}�X�^�t���p�X�捞��
    sts = GetIni("FILE", SAGYO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        SAGYO_Open = True
        Exit Function
    End If
    FullPath = RTrim$(c)
    
    Do
        sts = BTRV(BtOpOpen, SAGYO_POS, SAGYOREC, Len(SAGYOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Function
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = SAGYO_Create()        '��ƊǗ��}�X�^�쐬
                If sts <> False Then
                    SAGYO_Open = True
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, SAGYO_POS, SAGYOREC, Len(SAGYOREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "��ƊǗ��}�X�^")
                    SAGYO_Open = True
                End If
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "��ƊǗ��}�X�^")
                SAGYO_Open = True
                Exit Function
        End Select
    Loop
End Function
