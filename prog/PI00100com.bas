Attribute VB_Name = "PI00100com"
Option Explicit

Private Type Item_Key_tag
    JGYOBU  As String * 1
    NAIGAI  As String * 1
End Type

Public K_Item_Tbl() As Item_Key_tag   '�����ޕi�ڏ��
Public G_Item_Tbl() As Item_Key_tag   '�O�����ޕi�ڏ��



Private Type D_Item_Tbl_Tag
    SYUBETSU    As String * 2               '���
    JGYOBU      As String * 1               '���ƕ�
    NAIGAI      As String * 1               '�����O
    HIN_GAI     As String * 20              '�i��
    QTY         As Double                   '����
    SHIJI_QTY   As Double                   '���ʁi�w�����j
    BIKOU       As String * 40              '���l�i���͒l�j
    ID_NO       As String * 12              'ID_No(�o�ח\��ID_No)
End Type



Public D_Item_Tbl()     As D_Item_Tbl_Tag   '�����^�\���i�ڏ��


Public Taget_Key        As String * 8       '�X�V�Ώۂ̎w�}�[�� 2008.02.13

Public Taget_SHIMUKE_CODE_KEY _
                        As String * 2       '����Ώہ@�d������ 2008.02.02

Public Taget_Hin_key    As String * 20      '����Ώہ@�i��     2008.02.02
Public Taget_JGYOBU_key As String * 1       '����Ώہ@���ƕ�   2008.02.02
Public Taget_NAIGAI_key As String * 1       '����Ώہ@�����O   2008.02.02


Public Doukon_Tbl_No(0 To 19) _
                        As String * 1

Public Doukon_Start     As Integer          '��ʊJ�n�s��

Public POS_UMU          As Boolean

Public PRI_S_TANTO      As Boolean      '���x�^�S���҈�� OFF:����Ȃ� ON:�������
Public PRI_MAIN_BCR     As Boolean      'Ҳ��ް���� OFF:����Ȃ� ON:�������
Public PRI_BIKOU_BCR    As Boolean      '���l���@OFF�F���͒l�@ON:�o��BCR

'2011.08.04
'Public PRI_DOUKON       As Boolean     '���i�������@���� OFF:����Ȃ� ON:�������
Public PRI_DOUKON       As Integer      '���i�������@���� 0:�����E�������Ȃ� 1:������ 2:�����
'2011.08.04

Public PRI_NYUKO_IN     As Boolean      '���Ɋ�����@���� OFF:����Ȃ� ON:�������

Public PRI_INPUT_IN     As Boolean      '���͊�����@���� OFF:����Ȃ� ON:�������



Public PRI_SAGYO_DAY    As Boolean      '��Ɠ��^���ʁ^�S�� OFF:����Ȃ� ON:������� 2007.05.22
Public PRI_HINBAN_BIKOU As Boolean      '�����@�i�ԁ^���^���� OFF:����Ȃ� ON:������� 2007.05.22


Public JISEKI_TITLE     As Variant      '���ӂ̖��̃^�C�g��
Public TASEKI_TITLE     As Variant      '���ӂ̖��̃^�C�g��




Public HIN_INV          As Boolean      '���o�^�i�ԉ�


Public LabelPrint_F     As String       '2008.05.30




Public JISSEKI_DSP      As String * 1   '2008.08.19



Public chk_TORI_GENSANKOKU  As String * 20 '���Y���L�������p   2013.01.08


Public KAIKON_PRI       As Boolean      '�J���E���[�h�����E�S���h�~�E�̕\�� 2013.01.16


Public GENSANKOKU_MSG_F As Boolean      '���Y��ү���ޕ\���L��   2013.02.19



Public KAISYA_DEF_VALUE     As String   '��̫�ĉ�к���     2013.03.28
Public JIGYOBU_DEF_VALUE    As String   '��̫�Ď��ƕ�����   2013.03.28

Public NYUKA_KANSYOZAI  As Boolean      '���׎��ɏՍނ̕\�� 2013.11.05
    
    
    
Public PRINT_STOP_F     As Boolean      '������~�@2015.03.26
    
Public LABEL_PLUS        As Integer      '���x�����s���� 2015.04.02
    
Public GAI_BUHIN_CHK    As Boolean      '�C�O�����敪�`�F�b�N 2015.07.23
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   �O��C���N����
'Public Const Last_Update_Day$ = "(PI00010 2015.12.14 17:30)"



Public PI000104_Error_F     As Integer      '2019.03.14
Public PI000104_HIN_GAI     As String * 20  '2019.03.14
Public PI000104_OLD_HIN_GAI As String * 20  '2019.03.14

Public PI000104_CANCEL_F    As Integer      '2019.03.14

'---------------------------------------------- *���i���w�}�ް��i�e�j�ʃ|�C���^
'�|�W�V���j���O
Public wP_SSHIJI_O_POS  As POSBLK
'�f�[�^�E�o�b�t�@
Public wP_SSHIJI_O_REC  As P_SSHIJI_O_REC_Tag
'�L�[�E�f�[�^
Public K0_wP_SSHIJI_O   As KEY0_P_SSHIJI_O
Public K1_wP_SSHIJI_O   As KEY1_P_SSHIJI_O

' ------------------------------------------------------------------------
'       �w�肵�����x�̐��l�Ɏl�̌ܓ����܂��B
'
' @Param    dValue      �ۂߑΏۂ̔{���x���������_���B
' @Param    iDigits     �߂�l�̗L�������̐��x�B
' @Return               iDigits �ɓ��������x�̐��l�Ɏl�̌ܓ����ꂽ���l�B
' ------------------------------------------------------------------------
Public Function ToHalfAdjust(ByVal dValue As Currency, ByVal iDigits As Integer) As Currency
    Dim dCoef As Double

    dCoef = (10 ^ iDigits)

    If dValue > 0 Then
        ToHalfAdjust = Int(CDbl(dValue * dCoef + 0.5)) / dCoef
    Else
        ToHalfAdjust = Fix(CDbl(dValue * dCoef - 0.5)) / dCoef
    End If
End Function

Public Function File_Open_Proc() As Integer
'----------------------------------------------------------------------------
'               �t�@�C���@�n�o�d�m����
'           2015.03.13
'           2015.04.24 Sub --> Function
'----------------------------------------------------------------------------
                                
Dim sts     As Integer
                                
    File_Open_Proc = True
                                
    DoEvents
                                
Call LOG_OUT(LOG_F, "File �ăI�[�v������ �@�J�n")           '2015.03.26
                                
    sts = BTRV(BtOpReset, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
        Exit Function
    End If
                                
                                
                                
                                '���ԃ}�X�^�n�o�d�m
    If HATUBAN_Open(BtOpenNomal) Then
'        Unload PI000101
        Exit Function
    End If
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
'        Unload PI000101
        Exit Function
    End If

                                '���i���ٗp�i�ڃ}�X�^�n�o�d�m
    If L_ITEM_Open(BtOpenNomal) Then
'        Unload PI000101
        Exit Function
    End If

                                '�N���X�}�X�^�n�o�d�m
    If P_Class_Open(BtOpenNomal) Then
'        Unload PI000101
        Exit Function
    End If
                                '�R�[�h�}�X�^�n�o�d�m
    If P_CODE_Open(BtOpenNomal) Then
'        Unload PI000101
        Exit Function
    End If
                                '�\���}�X�^�n�o�d�m
    If P_COMPO_Open(BtOpenNomal) Then
'        Unload PI000101
        Exit Function
    End If
                                '�Ǘ��}�X�^�n�o�d�m
    If P_KANRI_Open(BtOpenNomal) Then
'        Unload PI000101
        Exit Function
    End If
                                '���i���w�}�i�q�j�ް��n�o�d�m
    If P_SSHIJI_K_Open(BtOpenNomal) Then
'        Unload PI000101
        Exit Function
    End If
                                '���i���w�}�i�e�j�ް��n�o�d�m
    If P_SSHIJI_O_Open(BtOpenNomal) Then
'        Unload PI000101
        Exit Function
    End If
                                '�S���҃}�X�^�n�o�d�m
    If TANTO_Open(BtOpenNomal) Then
'        Unload PI000101
        Exit Function
    End If
                                '�o�ח\���ް��n�o�d�m
    If Y_SYU_Open(BtOpenNomal) Then
'        Unload PI000101
        Exit Function
    End If
                                '�󕥐�}�X�^�n�o�d�m
    If P_UKEHARAI_Open(BtOpenNomal) Then
'        Unload PI000101
        Exit Function
    End If


    '2010.07.20 ��
                                '���Y���}�X�^�n�o�d�m
    If GENSAN_Open(BtOpenNomal) Then
'        Unload PI000101
        Exit Function
    End If
    '2010.07.20 ��
                                '�݌��ް��n�o�d�m
    If ZAIKO_Open(BtOpenNomal) Then
'        Unload PI000101
        Exit Function
    End If

                                '���i���w�}�i�e�jܰ��n�o�d�m
    If wP_SSHIJI_O_Open(BtOpenNomal) Then
'        Unload PI000101
        Exit Function
    End If

                                '���o�ɒP���ݒ�}�X�^�n�o�d�m   2008.09.20
    If SE_LOC_TANKA_M_Open(BtOpenNomal) Then
'        Unload PI000101
        Exit Function
    End If


    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  MT_2009.06.01
                                'PN�}�X�^�n�o�d�m
    If PN_M_Open(0) Then
'        Beep
'        MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
'        Unload PI000101
        Exit Function
    End If
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


Call LOG_OUT(LOG_F, "File �ăI�[�v������ �@����I��")           '2015.03.26

    File_Open_Proc = False

End Function

Public Function wP_SSHIJI_O_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���i���w�}(�e)���[�N  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    wP_SSHIJI_O_Open = True
                                            '���i���w�}(�e)�ް��t���p�X�捞��
    sts = GetIni("FILE", P_SSHIJI_O_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_SSHIJI_O]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, wP_SSHIJI_O_POS, wP_SSHIJI_O_REC, Len(wP_SSHIJI_O_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case Else
                Call File_Error(sts, BtOpOpen, "���i���w�}(�e)ܰ�")
                Exit Function
        End Select
    Loop

    wP_SSHIJI_O_Open = False

End Function

Sub Main()
    
    
    
Dim lngReturnValue      As Long
Dim strMyTitle          As String
Dim lngPrevHwnd         As Long
Dim lngTopHwnd          As Long
Dim lngThreadID1        As Long
Dim lngThreadID2        As Long
    
    
    
    




    ' 2�d�N���̏ꍇ�́A��O�Ɏ����Ă��Ď������g�͏I������
    strMyTitle = App.Title
    App.Title = "$" & App.Title
    lngPrevHwnd = FindWindow("ThunderRT6Main", strMyTitle)
    If lngPrevHwnd <> 0 Then
    lngTopHwnd = GetLastActivePopup(lngPrevHwnd)
    If IsIconic(lngTopHwnd) = WIN32API_TRUE Then
    lngReturnValue = ShowWindow(lngTopHwnd, SW_NORMAL)
    End If
    lngThreadID1 = GetWindowThreadProcessId(GetForegroundWindow(), ByVal 0&)
    lngThreadID2 = GetCurrentThreadId()
    lngReturnValue = AttachThreadInput(lngThreadID2, lngThreadID1, 1)
    lngReturnValue = SetForegroundWindow(lngTopHwnd)
    lngReturnValue = AttachThreadInput(lngThreadID2, lngThreadID1, 0)
    Exit Sub
    End If
    App.Title = strMyTitle










    PI000101.Show
End Sub

