Attribute VB_Name = "country"
Option Explicit
'********************************************************************
'*
'*              Country�}�X�^ �t�@�C����`
'*
'*          CREATE 2010.09.01
'********************************************************************
'�t�@�C���h�c
Public Const Country_ID = "Country"

'�y�[�W�T�C�Y
Public Const Country_PG_SIZ% = 4096

'�|�W�V�����E�u���b�N
Public Country_POS  As POSBLK
'=
'=
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type CountryREC_Tag
    CountryCode(0 To 2)     As Byte     '���R�[�h
    CountryName(0 To 19)    As Byte     '�����P
    CountryName2(0 To 19)   As Byte     '�����Q
    

End Type
'�f�[�^�E�o�b�t�@
Public CountryREC           As CountryREC_Tag


'�L�[��`
Type KEY0_Country                       '�j�d�x�O
    CountryCode(0 To 2)     As Byte     '���R�[�h
End Type



'�L�[�E�f�[�^
Public K0_Country           As KEY0_Country

Private Type Country_FSpeck
    fs  As BtFileSpeck              ' ̧�� ��߯��\����
    ks0 As BtKeySpeck               ' �� ��߯��\����
End Type

Private Country_Speck       As Country_FSpeck
Private Function Country_Create() As Integer
'********************************************************************
'*
'*              Country�t�@�C��  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'*          CREATE 2010.09.01
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    Country_Create = True
                                            'Country�t�@�C���t���p�X�捞��
    sts = GetIni("FILE", Country_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI[Country] �ǂݍ��݃G���[")
        Exit Function
    End If

    FullPath = RTrim$(c)

    Country_Speck.fs.recoleng = Len(CountryREC)     ' ���R�[�h��
    Country_Speck.fs.PageSize = Country_PG_SIZ      ' �y�[�W�T�C�Y
    Country_Speck.fs.idexnumb = 1                   ' �C���f�b�N�X��
    Country_Speck.fs.fileflag = 0                   ' �t�@�C���t���O
    Country_Speck.fs.reserve = &H0                  ' �\��ς�

'---------------------------------------------------' �L�[�O
    Country_Speck.ks0.keypos = 1                    ' �L�[�|�W�V����
    Country_Speck.ks0.keyleng = 3                   ' �L�[��
                                                    ' �L�[�t���O
    Country_Speck.ks0.keyflag = BtKfExt + BtKfChg
    Country_Speck.ks0.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    Country_Speck.ks0.reserve = &H0                 ' �\��ς�

    
    
    sts = BTRV(BtOpCreate, Country_POS, Country_Speck, Len(Country_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "Country�}�X�^")
        Exit Function
    End If

    Country_Create = False

End Function

Function Country_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              Country�}�X�^  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'*          CREATE 2010.09.01
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    Country_Open = True
                                            'Country�t�@�C���t���p�X�捞��
    sts = GetIni("FILE", Country_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim$(c)

    Do
        sts = BTRV(BtOpOpen, Country_POS, CountryREC, Len(CountryREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = Country_Create()        'Country�t�@�C���쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, Country_POS, CountryREC, Len(CountryREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "Country�}�X�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "Country�}�X�^")
                Exit Function
        End Select
    Loop
    Country_Open = False
End Function

