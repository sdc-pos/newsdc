Attribute VB_Name = "MainF104019"
Option Explicit

'---------------------------------------------- *�X�V�p�݌Ƀ��[�N
'�|�W�V���j���O
Public wZAIKO_POS   As POSBLK
'�f�[�^�E�o�b�t�@
Public wZAIKOREC    As ZAIKOREC_Tag
'�L�[�E�f�[�^
Public K0_wZAIKO    As KEY0_ZAIKO
Public K1_wZAIKO    As KEY1_ZAIKO
Public K2_wZAIKO    As KEY2_ZAIKO





Sub Main()
    Last_JGYOBU = Trim(Command)

    F1040191.Show
End Sub

Public Function wZAIKO_Open(Mode As Integer) As Integer
'****************************************************
'*      �u�ړ������v    �݌ɂn�o�d�m����
'*
'*  �݌Ƀt�@�C����ʃ|�C���^�łn�o�d�m����
'*  (�Ăь��ŋN�����ɂP�x�����Ăяo��)

'*  �߂�l: false       :����
'*          true        :�ُ�
'*          SYS_CANCEL  :�X�V��ݾ�
'****************************************************
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String

Dim ans         As Integer
    
    
    wZAIKO_Open = True
                                '�݌Ƀf�[�^�@�t���p�X�捞��
    sts = GetIni("FILE", ZAIKO_ID, "SYS", c)
    
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, wZAIKO_POS, wZAIKOREC, Len(wZAIKOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
'-------------- �n�o�d�m�����ł̎g�p���́A�����グ���ɂP�񂾂��̂͂��Ȃ̂ŁA��ɉ�ʓ��͂Ƃ��A
'               ��ݾق́A�����̋N����ݾقƂ���B
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    wZAIKO_Open = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpOpen, "�݌Ƀf�[�^")
                Exit Function
        End Select
    Loop

    wZAIKO_Open = False

End Function

Public Function wZAIKO_CLOSE() As Integer

'****************************************************
'*      �u�ړ������v    �݌ɂb�k�n�r�d����
'*
'*  �݌Ƀt�@�C����ʃ|�C���^�łb�k�n�r�d����
'*  (�Ăь��ŏI�����ɂP�x�����Ăяo��)
'*  �߂�l: false       :����
'*          true        :�ُ�
'****************************************************
Dim sts As Integer
    
    wZAIKO_CLOSE = True
    
    sts = BTRV(BtOpClose, wZAIKO_POS, wZAIKOREC, Len(wZAIKOREC), K0_wZAIKO, Len(K0_wZAIKO), 0)
    
    Select Case sts
        Case BtNoErr, BtErrNoOpen
        Case Else
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
            Exit Function
    End Select

    wZAIKO_CLOSE = False

End Function

