Attribute VB_Name = "MONTHLYQTY"
Option Explicit
'********************************************************************
'*                                                                  *
'*              �����Ϗo�א�(���ʏW�v)  �t�@�C����`                *
'*                                                                  *
'*          CREATE 2008.07.08                                       *
'********************************************************************
'�t�@�C���h�c
Public Const MONTHLYQTY_ID$ = "MONTHLYQTY"

'�y�[�W�T�C�Y
Public Const MONTHLYQTY_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public MONTHLYQTY_POS       As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type MONTHLYQTYREC_Tag
    DT(0 To 7)              As Byte         '���t
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '�i��(�O��)
    SyukaCnt(0 To 4)        As Byte         '�o�׉�
    SyukaQty(0 To 4)        As Byte         '�o�א���



End Type

'�f�[�^�E�o�b�t�@
Public MONTHLYQTYREC        As MONTHLYQTYREC_Tag

'�L�[��`

Type KEY0_MONTHLYQTY                    '�j�d�x�O
    DT(0 To 7)              As Byte         '���t
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '�i��(�O��)
End Type

'Type KEY1_MONTHLYQTY                    '�j�d�x�P
'    JGYOBU(0 To 0)          As Byte         '���ƕ�
'    NAIGAI(0 To 0)          As Byte         '�����O
'    HIN_GAI(0 To 19)        As Byte         '�i��(�O��)
'    DT(0 To 7)              As Byte         '���t
'End Type

'�L�[�E�f�[�^
Public K0_MONTHLYQTY        As KEY0_MONTHLYQTY
'Public K1_MONTHLYQTY        As KEY1_MONTHLYQTY

Type MONTHLYQTY_FSpeck
    fs  As BtFileSpeck          ' ̧�� ��߯��\����
    ks0 As BtKeySpeck           ' �� ��߯��\����
    ks1 As BtKeySpeck           ' �� ��߯��\����
    ks2 As BtKeySpeck           ' �� ��߯��\����
    ks3 As BtKeySpeck           ' �� ��߯��\����
'    ks4 As BtKeySpeck           ' �� ��߯��\����
'    ks5 As BtKeySpeck           ' �� ��߯��\����
'    ks6 As BtKeySpeck           ' �� ��߯��\����
'    ks7 As BtKeySpeck           ' �� ��߯��\����
End Type

Public MONTHLYQTY_Speck     As MONTHLYQTY_FSpeck
 
Private Function MONTHLYQTY_Create() As Integer
'********************************************************************
'*                                                                  *
'*              �����Ϗo�א�(���ʏW�v)  �b�q�d�`�s�d                      �@  *
'*                                                                  *
'*      ��  ��:�Ȃ�                                                 *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'*          CREATE 2004.02.14                                       *
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    MONTHLYQTY_Create = True
                                            '�����Ϗo�א�(���ʏW�v)�t���p�X�捞��
    sts = GetIni("FILE", MONTHLYQTY_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    MONTHLYQTY_Speck.fs.recoleng = Len(MONTHLYQTYREC)   ' ���R�[�h��
    MONTHLYQTY_Speck.fs.PageSize = MONTHLYQTY_PG_SIZ%   ' �y�[�W�T�C�Y
    MONTHLYQTY_Speck.fs.idexnumb = 1                    ' �C���f�b�N�X��
    MONTHLYQTY_Speck.fs.fileflag = 0                    ' �t�@�C���t���O
    MONTHLYQTY_Speck.fs.reserve = &H0                   ' �\��ς�
                                                        ' �L�[�O
    MONTHLYQTY_Speck.ks0.keypos = 1                         ' �L�[�|�W�V����
    MONTHLYQTY_Speck.ks0.keyleng = 8                        ' �L�[��
    MONTHLYQTY_Speck.ks0.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
    MONTHLYQTY_Speck.ks0.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    MONTHLYQTY_Speck.ks0.reserve = &H0                      ' �\��ς�
                                                        ' �L�[�O
    MONTHLYQTY_Speck.ks1.keypos = 9                         ' �L�[�|�W�V����
    MONTHLYQTY_Speck.ks1.keyleng = 1                        ' �L�[��
    MONTHLYQTY_Speck.ks1.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
    MONTHLYQTY_Speck.ks1.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    MONTHLYQTY_Speck.ks1.reserve = &H0                      ' �\��ς�
                                                        ' �L�[�O
    MONTHLYQTY_Speck.ks2.keypos = 10                        ' �L�[�|�W�V����
    MONTHLYQTY_Speck.ks2.keyleng = 1                        ' �L�[��
    MONTHLYQTY_Speck.ks2.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
    MONTHLYQTY_Speck.ks2.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    MONTHLYQTY_Speck.ks2.reserve = &H0                      ' �\��ς�
                                                        ' �L�[�O
    MONTHLYQTY_Speck.ks3.keypos = 11                        ' �L�[�|�W�V����
    MONTHLYQTY_Speck.ks3.keyleng = 20                       ' �L�[��
    MONTHLYQTY_Speck.ks3.keyflag = BtKfExt                  ' �L�[�t���O
    MONTHLYQTY_Speck.ks3.keytype = Chr(BtKtString)          ' �L�[�^�C�v
    MONTHLYQTY_Speck.ks3.reserve = &H0                      ' �\��ς�




'                                                        ' �L�[�P
'    MONTHLYQTY_Speck.ks4.keypos = 9                         ' �L�[�|�W�V����
'    MONTHLYQTY_Speck.ks4.keyleng = 1                        ' �L�[��
'    MONTHLYQTY_Speck.ks4.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
'    MONTHLYQTY_Speck.ks4.keytype = Chr(BtKtString)          ' �L�[�^�C�v
'    MONTHLYQTY_Speck.ks4.reserve = &H0                      ' �\��ς�
'                                                        ' �L�[�P
'    MONTHLYQTY_Speck.ks5.keypos = 10                        ' �L�[�|�W�V����
'    MONTHLYQTY_Speck.ks5.keyleng = 1                        ' �L�[��
'    MONTHLYQTY_Speck.ks5.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
'    MONTHLYQTY_Speck.ks5.keytype = Chr(BtKtString)          ' �L�[�^�C�v
'    MONTHLYQTY_Speck.ks5.reserve = &H0                      ' �\��ς�
'                                                        ' �L�[�P
'    MONTHLYQTY_Speck.ks6.keypos = 11                        ' �L�[�|�W�V����
'    MONTHLYQTY_Speck.ks6.keyleng = 20                       ' �L�[��
'    MONTHLYQTY_Speck.ks6.keyflag = BtKfExt + BtKfSeg        ' �L�[�t���O
'    MONTHLYQTY_Speck.ks6.keytype = Chr(BtKtString)          ' �L�[�^�C�v
'    MONTHLYQTY_Speck.ks6.reserve = &H0                      ' �\��ς�
'                                                        ' �L�[�P
'    MONTHLYQTY_Speck.ks7.keypos = 1                         ' �L�[�|�W�V����
'    MONTHLYQTY_Speck.ks7.keyleng = 8                        ' �L�[��
'    MONTHLYQTY_Speck.ks7.keyflag = BtKfExt                  ' �L�[�t���O
'    MONTHLYQTY_Speck.ks7.keytype = Chr(BtKtString)          ' �L�[�^�C�v
'    MONTHLYQTY_Speck.ks7.reserve = &H0                      ' �\��ς�



    sts = BTRV(BtOpCreate, MONTHLYQTY_POS, MONTHLYQTY_Speck, Len(MONTHLYQTY_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�����Ϗo�א�(���ʏW�v)")
    End If

    MONTHLYQTY_Create = False

End Function

Function MONTHLYQTY_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              �����Ϗo�א�(���ʏW�v)  �n�o�d�m                    *
'*                                                                  *
'*      ��  ��:Open Mode(Btrieve�Q��)                               *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'*          CREATE 2008.07.08                                       *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    MONTHLYQTY_Open = True
                                            '�����Ϗo�א�(���ʏW�v) �t���p�X�捞��
    sts = GetIni("FILE", MONTHLYQTY_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, MONTHLYQTY_POS, MONTHLYQTYREC, Len(MONTHLYQTYREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = MONTHLYQTY_Create()   '�����Ϗo�א�(���ʏW�v) �쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, MONTHLYQTY_POS, MONTHLYQTYREC, Len(MONTHLYQTYREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�����Ϗo�א�(���ʏW�v)")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�����Ϗo�א�(���ʏW�v)")
                Exit Function
        End Select
    Loop

    MONTHLYQTY_Open = False
    
End Function
