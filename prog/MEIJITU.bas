Attribute VB_Name = "MEIJ"
Option Explicit
'********************************************************************
'*                                                                  *
'*              ���o�ז��׃f�[�^�@�t�@�C����`                        *
'*                                                                  *
'*          CREATE 2001.05.15                                       *
'********************************************************************
'�t�@�C���h�c
Global Const MEIJ_ID = "MEIJ"

'�y�[�W�T�C�Y
Global Const MEIJ_PG_SIZ% = 2048

'�|�W�V�����E�u���b�N
Global MEIJ_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                              *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type MEIJREC_Tag
    IO_KBN(0 To 0)      As Byte     '���ׁ^�o�׋敪
    DEN_DT(0 To 7)      As Byte     '�`�[���t
    CYU_KBN(0 To 0)     As Byte     '�����敪
    NAIGAI(0 To 0)      As Byte     '�����O
    HIN_GAI(0 To 19)    As Byte     '�i�ԁi�O���j
    JITU_QTY(0 To 8)    As Byte     '���ѐ�
    FILLER(0 To 7)     As Byte     'FILLER
End Type

'�f�[�^�E�o�b�t�@
Public MEIJREC As MEIJREC_Tag

'�L�[��`
Type KEY0_MEIJ            '�j�d�x�O
    IO_KBN(0 To 0)      As Byte     '���ׁ^�o�׋敪
    DEN_DT(0 To 7)      As Byte     '�`�[���t
    CYU_KBN(0 To 0)     As Byte     '�����敪
    NAIGAI(0 To 0)      As Byte     '�����O
    HIN_GAI(0 To 19)    As Byte     '�i�ԁi�O���j
End Type

'�L�[�E�f�[�^
Public K0_MEIJ As KEY0_MEIJ

Type MEIJ_FSpeck
    fs As BtFileSpeck               ' ̧�� ��߯��\����
    ks0 As BtKeySpeck               ' �� ��߯��\����
End Type

Global MEIJ_Speck As MEIJ_FSpeck

Private Function MEIJ_Create() As Integer
'********************************************************************
'*
'*              ���o�ז��׃f�[�^�@�b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'*          CREATE 2001.05.15
'********************************************************************
Dim sts As Integer
Dim FullPath As String
Dim c As String * 128

    MEIJ_Create = False
                                            '���o�׎��яW�v�f�[�^�t���p�X�捞��
    sts = GetIni("FILE", MEIJ_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        MEIJ_Create = True
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    MEIJ_Speck.fs.recoleng = Len(MEIJREC)       ' ���R�[�h��
    MEIJ_Speck.fs.PageSize = MEIJ_PG_SIZ        ' �y�[�W�T�C�Y
    MEIJ_Speck.fs.idexnumb = 1                  ' �C���f�b�N�X��
    MEIJ_Speck.fs.fileflag = 0                  ' �t�@�C���t���O
    MEIJ_Speck.fs.reserve = &H0                 ' �\��ς�
                                                ' �L�[�O
    MEIJ_Speck.ks0.keypos = 1                   ' �L�[�|�W�V����
    MEIJ_Speck.ks0.keyleng = 1 + 8 + 1 + 1 + 20 ' �L�[��
    MEIJ_Speck.ks0.keyflag = BtKfExt            ' �L�[�t���O
    MEIJ_Speck.ks0.keytype = Chr(BtKtString)    ' �L�[�^�C�v
    MEIJ_Speck.ks0.reserve = &H0                ' �\��ς�

    sts = BTRV(BtOpCreate, MEIJ_POS, MEIJ_Speck, Len(MEIJ_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���o�ז��׃f�[�^")
        MEIJ_Create = True
    End If
End Function

Function MEIJ_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���o�ז��׃f�[�^�@�n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'*          CREATE 2001.05.15
'********************************************************************
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String
    
    MEIJ_Open = False
                                            '���o�ז��׃f�[�^�t���p�X�捞��
    sts = GetIni("FILE", MEIJ_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        MEIJ_Open = True
        Exit Function
    End If
    FullPath = RTrim$(c)
    
    Do
        sts = BTRV(BtOpOpen, MEIJ_POS, MEIJREC, Len(MEIJREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Function
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = MEIJ_Create()        '���o�׎��яW�v�f�[�^�쐬
                If sts <> False Then
                    MEIJ_Open = True
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, MEIJ_POS, MEIJREC, Len(MEIJREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���o�ז��׃f�[�^")
                    MEIJ_Open = True
                End If
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "���o�ז��׃f�[�^")
                MEIJ_Open = True
                Exit Function
        End Select
    Loop
End Function


