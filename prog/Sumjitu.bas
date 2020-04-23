Attribute VB_Name = "SUMJ"
Option Explicit
'********************************************************************
'*                                                                  *
'*              ���o�׎��яW�v�f�[�^�@�t�@�C����`                          *
'*                                                                  *
'*          CREATE 1997.06.02  S.Shibano                            *
'********************************************************************
'�t�@�C���h�c
Global Const SUMJ_ID = "SUMJ"

'�y�[�W�T�C�Y
Global Const SUMJ_PG_SIZ% = 2048

'�|�W�V�����E�u���b�N
Global SUMJ_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                              *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type SUMJREC_Tag
    JGYOBU(0 To 0) As Byte          '���ƕ��敪
    NAIGAI(0 To 0) As Byte          '�����O
    HIN_GAI(0 To 19) As Byte        '�i�ԁi�O���j
    NYUKA_QTY(0 To 7) As Byte       '���ב���
    CHOKU_QTY(0 To 7) As Byte       '���ג�����
    TUK_QTY(0 To 7) As Byte         '���؂�o�א�
    HSP_QTY(0 To 7) As Byte         '��[�X�|�b�g�o�א� (�܂ߓ���)
    BOU_QTY(0 To 7) As Byte         '�f�Տo�א�
    KIN_QTY(0 To 7) As Byte         '�ً}�o�א�
    ZAI_PURA(0 To 7) As Byte        '�ݒ��i�{�j�o�ɐ�
    ZAI_MINA(0 To 7) As Byte        '�ݒ��i�|�j�o�ɐ�
    FILLER(0 To 9) As Byte          'FILLER
End Type

'�f�[�^�E�o�b�t�@
Global SUMJREC As SUMJREC_Tag

'�L�[��`
Type KEY0_SUMJ            '�j�d�x�O
    JGYOBU(0 To 0) As Byte          '���ƕ��敪
    NAIGAI(0 To 0) As Byte          '�����O
    HIN_GAI(0 To 19) As Byte        '�i�ԁi�O���j
End Type

'�L�[�E�f�[�^
Global K0_SUMJ As KEY0_SUMJ

Type SUMJ_FSpeck
    fs As BtFileSpeck               ' ̧�� ��߯��\����
    ks0 As BtKeySpeck               ' �� ��߯��\����
End Type

Global SUMJ_Speck As SUMJ_FSpeck

Private Function SUMJ_Create() As Integer
'********************************************************************
'*                                                                  *
'*              ���o�׎��яW�v�f�[�^�@�b�q�d�`�s�d                        *
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

    SUMJ_Create = False
                                            '���o�׎��яW�v�f�[�^�t���p�X�捞��
    sts = GetIni("FILE", SUMJ_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        SUMJ_Create = True
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    SUMJ_Speck.fs.recoleng = Len(SUMJREC)     ' ���R�[�h��
    SUMJ_Speck.fs.PageSize = SUMJ_PG_SIZ      ' �y�[�W�T�C�Y
    SUMJ_Speck.fs.idexnumb = 1                 ' �C���f�b�N�X��
    SUMJ_Speck.fs.fileflag = 0                 ' �t�@�C���t���O
    SUMJ_Speck.fs.reserve = &H0                ' �\��ς�
                                                ' �L�[�O
    SUMJ_Speck.ks0.keypos = 1                  ' �L�[�|�W�V����
    SUMJ_Speck.ks0.keyleng = 1 + 1 + 20        ' �L�[��
    SUMJ_Speck.ks0.keyflag = BtKfExt           ' �L�[�t���O
    SUMJ_Speck.ks0.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    SUMJ_Speck.ks0.reserve = &H0               ' �\��ς�

    sts = BTRV(BtOpCreate, SUMJ_POS, SUMJ_Speck, Len(SUMJ_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���o�׎��яW�v�f�[�^")
        SUMJ_Create = True
    End If
End Function

Function SUMJ_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              ���o�׎��яW�v�f�[�^�@�n�o�d�m                            *
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
    
    SUMJ_Open = False
                                            '���o�׎��яW�v�f�[�^�t���p�X�捞��
    sts = GetIni("FILE", SUMJ_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        SUMJ_Open = True
        Exit Function
    End If
    FullPath = RTrim$(c)
    
    Do
        sts = BTRV(BtOpOpen, SUMJ_POS, SUMJREC, Len(SUMJREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Function
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = SUMJ_Create()        '���o�׎��яW�v�f�[�^�쐬
                If sts <> False Then
                    SUMJ_Open = True
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, SUMJ_POS, SUMJREC, Len(SUMJREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���o�׎��яW�v�f�[�^")
                    SUMJ_Open = True
                End If
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "���o�׎��яW�v�f�[�^")
                SUMJ_Open = True
                Exit Function
        End Select
    Loop
End Function


