Attribute VB_Name = "ODR1010G"
Option Explicit
'********************************************************************
'*
'*              �n�c�q�P�O�P�O�p�@���ʕϐ�
'*
'********************************************************************

Public ODR10102_Return As Integer         '�m�F��ʏI�����


Public GW_PURA(0 To 30)          As String       '�ݒ��i�{�j�R�[�h
Public GW_MAINA(0 To 30)         As String       '�ݒ��i�|�j�R�[�h


Public DIS_ORDR_NO      As String       '�e���i�@������
Public DIS_BUNNO        As String       '���[��
Public DIS_OYA_ITEM     As String       '�e���i�R�[�h
Public DIS_ORDR_QTY     As String       '��������
Public DIS_NOUKI        As String       '�e���i�@�����[��
Public DIS_OK_DT        As String       '�g���\��
Public DIS_KAITO        As String       '�e���i�@�񓚔[��
Public DIS_USE_YM       As String       '�g�p��
Public DIS_FIN_DT       As String       '�������t
Public DIS_KEY          As String       '�f�[�^�j�������

Public DIS2_QTY         As String       '��������
Public DIS2_KAITO       As String       '�e���i�@�񓚔[��

Public Key_SIMUKE       As String       '�d������
Public Key_JIGYOBU      As String       '���ƕ�
Public Key_NAIGAI       As String       '�����O
Public Key_USE_YM       As String       '�g�p���iYYYYMM)
Public Key_INS_NO       As String       '�o�^��
Public Key_HinGai       As String       '�e�i��
Public Key_ORDER_NO     As String       '�e�i�ԁ@������
Public Key_BUN_NO       As String       '���[��

Public Key_Ko_HinGai    As String       '�q�i��         2010/05/07�ǉ�
Public Key_Ko_JIGYOBU      As String    '�q�i�� ���ƕ�
Public Key_Ko_NAIGAI       As String    '�q�i�� �����O
Sub Main()
'2017.01.16 �ǉ�
    
    
Dim lngReturnValue      As Long
Dim strMyTitle          As String
Dim lngPrevHwnd         As Long
Dim lngTopHwnd          As Long
Dim lngThreadID1        As Long
Dim lngThreadID2        As Long
    
    
    
    
    Last_JGYOBU = Trim(Command)






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










    ODR10101.Show
End Sub

Function OUT_TP1(HIN_GAI As String) As Integer
'
'           �\���W�J�@���@���v�ʂe�o��
'

Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim W_SeqKey    As String
Dim W_Ko_HinCD  As String

Dim W_QTY       As Double
Dim W_STR       As String
Dim W_Date      As String

Dim Fin_Qty     As Double           '2008.04.30

Dim W_A_Nouki   As String

    OUT_TP1 = True
        
    '
    '�w�肳�ꂽ�e�i�Ԃ̍\���q���i�̑S�ĂɊւ��āA�W�J����ݒ肷��I
    '
    W_SeqKey = ""
    
    If Trim(HIN_GAI) = "AD-HEPSC010" Then
        W_SeqKey = ""
    End If
    
    If Trim(StrConv(ODR_ORDER_REC.KAITO_DT, vbUnicode)) = "" Then
        W_A_Nouki = "99999999"
    Else
        W_A_Nouki = StrConv(ODR_ORDER_REC.KAITO_DT, vbUnicode)
    End If
    
    '   �ŏ��Ɂu�e���R�[�h�v���擾�B
    
    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, GW_SIMUKE)
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, GW_JIGYOBU)
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, GW_NAIGAI)
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, HIN_GAI)
    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, "0")
    Call UniCode_Conv(K0_P_COMPO.SEQNO, "")
    com = BtOpGetGreaterEqual
    sts = BTRV(com, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                'Beep
                'MsgBox "�w�肳�ꂽ�H��������܂���B"
        Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
            yn = MsgBox("���Ŏg�p���ł��I<�\���e>" & Chr(13) & Chr(10) & _
                            "�@�Ď��s���܂����H", vbYesNo + vbExclamation, "�m�F����")
            If yn = vbNo Then
                Exit Function
            End If
        Case Else
            Call File_Error(sts, com, "P_COMPO")
            Exit Function
    End Select
    If sts <> BtNoErr Then
        MsgBox "�\�����@���o�^�I <" & HIN_GAI & ">", vbExclamation
        Exit Function
    End If
    
    If Trim(StrConv(P_COMPO_O_REC.HIN_GAI, vbUnicode)) = "AD-DLHS03A05" Then
        sts = BtNoErr
    End If
    
    '   ��������u�q���i���R�[�h�v��ǂ݂Ȃ���W�J�e���o�͂���B
    com = BtOpGetNext
    Do
        sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                yn = MsgBox("���Ŏg�p���ł��I<�\���e>" & Chr(13) & Chr(10) & _
                            "�@�Ď��s���܂����H", vbYesNo + vbExclamation, "�m�F����")
                If yn = vbNo Then Exit Function
            Case Else
                Call File_Error(sts, com, "P_COMPO")
                Exit Function
        End Select
        If sts <> BtNoErr Then
            Exit Do
        End If
        If Trim(StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode)) <> Trim(GW_SIMUKE) Then Exit Do
        If Trim(StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode)) <> Trim(GW_JIGYOBU) Then Exit Do
        If Trim(StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode)) <> Trim(GW_NAIGAI) Then Exit Do
        If Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(HIN_GAI) Then Exit Do
        'If Trim(StrConv(P_COMPO_O_REC.DATA_KBN, vbUnicode)) <> "0" Then Exit Do
        
        W_SeqKey = StrConv(P_COMPO_K_REC.SEQNO, vbUnicode)
        
        If CInt(W_SeqKey) <> 0 Then                 '�\�����i���R�[�h�H
            
            Call ODR_TEMP1_CLR
    
            Call UniCode_Conv(ODR_TP1_R.KAITO_DT, W_A_Nouki)
            Call UniCode_Conv(ODR_TP1_R.CYUMON_DT, Trim(StrConv(ODR_ORDER_REC.CYUMON_DT, vbUnicode)))
            Call UniCode_Conv(ODR_TP1_R.USE_YM, Trim(StrConv(ODR_ORDER_REC.USE_YM, vbUnicode)))
            
            Call UniCode_Conv(ODR_TP1_R.SHIMUKE, GW_SIMUKE)
            Call UniCode_Conv(ODR_TP1_R.JGYOBU, GW_JIGYOBU)
            Call UniCode_Conv(ODR_TP1_R.NAIGAI, GW_NAIGAI)
            Call UniCode_Conv(ODR_TP1_R.INS_NO, Trim(StrConv(ODR_ORDER_REC.INS_NO, vbUnicode)))
            Call UniCode_Conv(ODR_TP1_R.ORDER_NO, Trim(StrConv(ODR_ORDER_REC.ORDER_NO, vbUnicode)))
            Call UniCode_Conv(ODR_TP1_R.BUN_NO, Trim(StrConv(ODR_ORDER_REC.BUN_NO, vbUnicode)))
            Call UniCode_Conv(ODR_TP1_R.HIN_GAI, Trim(HIN_GAI))
            
            

'2008.04.04            Call UniCode_Conv(ODR_TP1_R.KO_NAIGAI, Trim(StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode)))
'2008.04.04            Call UniCode_Conv(ODR_TP1_R.KO_JGYOBU, Trim(StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode)))
            
            Call UniCode_Conv(ODR_TP1_R.KO_JGYOBU, SHIZAI)      '2008.04.04
            Call UniCode_Conv(ODR_TP1_R.KO_NAIGAI, NAIGAI_NAI)  '2008.04.04
            
            Call UniCode_Conv(ODR_TP1_R.KO_HIN_GAI, Trim(StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)))
            
            Call UniCode_Conv(ODR_TP1_R.KO_SYUBETSU, Trim(StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode)))
            If IsNull(StrConv(ODR_TP1_R.KO_SYUBETSU, vbUnicode)) Then
                Call UniCode_Conv(ODR_TP1_R.KO_SYUBETSU, "")
            End If
            If Left(StrConv(ODR_TP1_R.KO_SYUBETSU, vbUnicode), 1) < " " Then
                Call UniCode_Conv(ODR_TP1_R.KO_SYUBETSU, "")
            End If
            Call UniCode_Conv(ODR_TP1_R.KO_QTY, Trim(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)))
            
            
            
            W_QTY = CDbl(Trim(StrConv(ODR_ORDER_REC.ODR_QTY, vbUnicode)))
            W_QTY = W_QTY * CDbl(Trim(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)))
            W_STR = CStr(W_QTY)
            Call UniCode_Conv(ODR_TP1_R.ALL_QTY, W_STR)     '�W�J��
                        
            If W_QTY < 0 Then
                W_QTY = W_QTY * 1
            End If
            
            W_STR = CStr(Abs(W_QTY))
    
            Fin_Qty = 0
            
            If CDbl(Trim(StrConv(ODR_ORDER_REC.ODR_QTY, vbUnicode))) <= 0 Then
                
                '08.11.27�R�����g�ɁI
                'Call UniCode_Conv(ODR_TP1_R.USE_QTY, W_STR)                     '�g�p��
                
            
            Else
            
                'If Trim(StrConv(ODR_ORDER_REC.FIN_DT, vbUnicode)) = "" Then
                '    Call UniCode_Conv(ODR_TP1_R.NED_QTY, W_STR)                 '�K�v��
                '    Call UniCode_Conv(ODR_TP1_R.REQ_QTY, W_STR)                 '���v��
                'Else
                '                '�@�������@���@�J�z���F���v���i�݌Ɉ����Ώہj�Ƃ���I
                '    If Trim(StrConv(ODR_ORDER_REC.FIN_DT, vbUnicode)) <= GW_SHIMEBI Then
                '        Call UniCode_Conv(ODR_TP1_R.USE_QTY, W_STR)             '�g�p��
                '        Call UniCode_Conv(ODR_TP1_R.REQ_QTY, W_STR)             '���v��
                '    Else
                '                '�@�������@���@�J�z���F�����Ȃ��I
                '
                '
                '    End If
                'End If
                '08.11.27��L�����L�ɕύX
                If Trim(StrConv(ODR_ORDER_REC.FIN_DT, vbUnicode)) = "" Then
                    Call UniCode_Conv(ODR_TP1_R.NED_QTY, W_STR)                 '�K�v��
                    Call UniCode_Conv(ODR_TP1_R.REQ_QTY, W_STR)                 '���v��
                    
                    '08.12.12�ǉ�
                    Call UniCode_Conv(ODR_TP1_R.KAN_KB, "1")                    '�������P
                    Call UniCode_Conv(ODR_TP1_R.OK_DT, "")
                Else
                            '����
                    Call UniCode_Conv(ODR_TP1_R.USE_QTY, W_STR)                 '�g�p��
                    
                    '08.12.12�ǉ�
                    Call UniCode_Conv(ODR_TP1_R.KAN_KB, "0")                    '�������O
                    Call UniCode_Conv(ODR_TP1_R.OK_DT, Trim(StrConv(ODR_ORDER_REC.KUMI_OK_DT, vbUnicode)))
                    '   ���������A���X�̑g���\�����Z�b�g�I
                End If
            
            End If


            Call UniCode_Conv(ODR_TP1_R.UPDT_DT, Format(Date, "yyyymmdd"))
            Call UniCode_Conv(ODR_TP1_R.UPDT_TM, Format(Time, "hhmmss"))
            
            
            '2008/09/19 �i�ڂl�o�^�`�F�b�N
            Call UniCode_Conv(K0_ITEM.JGYOBU, RTrim(StrConv(ODR_TP1_R.KO_JGYOBU, vbUnicode)))
            Call UniCode_Conv(K0_ITEM.NAIGAI, RTrim(StrConv(ODR_TP1_R.KO_NAIGAI, vbUnicode)))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, RTrim(StrConv(ODR_TP1_R.KO_HIN_GAI, vbUnicode)))
                
            Do
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        
                Select Case sts
                    Case BtNoErr
                        
                        Exit Do
                        
                    Case BtErrKeyNotFound, BtErrEOF
                        W_STR = StrConv(ODR_TP1_R.KO_JGYOBU, vbUnicode)
                        W_STR = W_STR & "-" & StrConv(ODR_TP1_R.KO_NAIGAI, vbUnicode)
                        W_STR = W_STR & "-" & Trim(StrConv(ODR_TP1_R.KO_HIN_GAI, vbUnicode))
                        
                        'yn = MsgBox("�i�ږ��o�^�I<" & W_Str & ">" & Chr(13) & Chr(10) & _
                        '            "�@���s���܂����H", vbYesNo + vbDefaultButton1 + vbExclamation, "�m�F����")
                        yn = vbYes
                        If yn = vbNo Then Exit Function
                        Exit Do
                        
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        yn = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If yn <> vbYes Then
                            Exit Function
                        End If
                        
                        
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                        Exit Function
                End Select
            Loop
    
    
    
                                    '2008/09/19 �i�ڂl�����F�o�^���Ȃ��I
            If Trim(StrConv(ODR_TP1_R.KO_HIN_GAI, vbUnicode)) = "TEST" Then
                sts = BtNoErr
            End If
            
            
            If sts = BtNoErr Then
'                Call UniCode_Conv(ODR_TP1_R.FILLER, "ITEM���o�^")   '2010/06/15 �v�]�Œǉ��I
'            End If
            '2010/06/15     �ēx�A�W�J�i�o�́j���Ȃ��悤�ɏC���I
            '               ���̕��A�q���i�W�J���10103�ɍ\���l���e��\���I
            '
                Do
                    sts = BTRV(BtOpInsert, ODR_TP1_POS, ODR_TP1_R, Len(ODR_TP1_R), K0_ODR_TEMP1, Len(K0_ODR_TEMP1), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                            Sleep (500)
                        Case Else
                            Call File_Error(sts, BtOpInsert, "ODR_TEMP1")
                            Exit Function
                    End Select
                Loop
            End If
            
            Key_SIMUKE = GW_SIMUKE
            
'2008.04.04            Key_JIGYOBU = Trim(StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
'2008.04.04.            Key_NAIGAI = Trim(StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
            Key_JIGYOBU = SHIZAI        '2008.04.04
            Key_NAIGAI = NAIGAI_NAI     '2008.04.04
            
            GW_HINGAI_KO = Trim(StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
            
'2008.04.04            GW_JIGYOBU_KO = StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode)
'2008.04.04            GW_NAIGAI_KO = StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode)
            GW_JIGYOBU_KO = SHIZAI      '2008.04.04
            GW_NAIGAI_KO = NAIGAI_NAI   '2008.04.04
            
        End If
        
        com = BtOpGetNext
    Loop
        
    OUT_TP1 = False


End Function


Function SET_ALL() As Integer
'
'                                           TP1����ɍ݌ɏ����Z�b�g����eSUB
'
Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim W_S         As Integer
            
    SET_ALL = True
           
    GW_JIGYOBU_KO = ""
    GW_NAIGAI_KO = ""
    GW_HINGAI_KO = ""
    
    com = BtOpGetFirst
    Do
        Do
            sts = BTRV(com, ODR_TP1_POS, ODR_TP1_R, Len(ODR_TP1_R), K2_ODR_TEMP1, Len(K2_ODR_TEMP1), 2)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                    
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                    yn = MsgBox("���Ŏg�p���ł��I<���Ԃe�P>" & Chr(13) & Chr(10) & _
                                "�@�Ď��s���܂����H", vbYesNo + vbExclamation, "�m�F����")
                    If yn = vbNo Then Exit Function
                Case Else
                    Call File_Error(sts, com, "ODR_TEMP1")
                    Exit Function
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
        
        W_S = 0
        If GW_JIGYOBU_KO <> StrConv(ODR_TP1_R.KO_JGYOBU, vbUnicode) Then W_S = 1
        If GW_NAIGAI_KO <> StrConv(ODR_TP1_R.KO_NAIGAI, vbUnicode) Then W_S = 1
        If GW_HINGAI_KO <> StrConv(ODR_TP1_R.KO_HIN_GAI, vbUnicode) Then W_S = 1
        
        
        GW_JIGYOBU_KO = StrConv(ODR_TP1_R.KO_JGYOBU, vbUnicode)
        GW_NAIGAI_KO = StrConv(ODR_TP1_R.KO_NAIGAI, vbUnicode)
        GW_HINGAI_KO = StrConv(ODR_TP1_R.KO_HIN_GAI, vbUnicode)
        
        
        If W_S = 1 Then
            
                '       �݌ɏ��������E�o��
            If SET_I_ZAIKO() Then
                MsgBox "�݌ɏ��ݒ�G���[�I", vbExclamation
                Exit Function
            End If
            
                '       �������i���ɗ\��j���@�����E�o��
            If SET_ODR_ZAN() Then
                MsgBox "�����ɏ��ݒ�G���[�I", vbExclamation
                Exit Function
            End If
        
                '2008/09/10 �ݒ����̍݌ɏ����W�v
            If SET_ZAITEI() Then
                MsgBox "�ݒ����̍݌ɏ��ݒ�G���[�I", vbExclamation
                Exit Function
            End If
                    
            '2008/05/31 �����i�̍݌ɏ����W�v
            If SET_H_SEIHIN() Then
                MsgBox "�����i�̍݌ɏ��ݒ�G���[�I", vbExclamation
                Exit Function
            End If
        
        End If
        
        
        com = BtOpGetNext
    Loop
        
                '       �d�����я��@�����E�o��
    If SET_UKEIRE() Then
        MsgBox "�d�����я��ݒ�G���[�I", vbExclamation
        Exit Function
    End If
            
     
    SET_ALL = False

End Function
Function SET_I_ZAIKO() As Integer

        '       �݌ɏ��������E�o�́i�e�[�u���j
        
        '       �݌ɐ����O�ł��o�́i08/09/11�j
        
Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim W_STR       As String
Dim W_Edit      As String
Dim W_Zaiko     As Double
Dim W_QTY       As Double
    
    SET_I_ZAIKO = True

    Call UniCode_Conv(K0_ODR_TEMP2.KO_JGYOBU, GW_JIGYOBU_KO)        '���ƕ�
    Call UniCode_Conv(K0_ODR_TEMP2.KO_NAIGAI, GW_NAIGAI_KO)         '�����O
    Call UniCode_Conv(K0_ODR_TEMP2.KO_HIN_GAI, GW_HINGAI_KO)        '�q�i��
    Call UniCode_Conv(K0_ODR_TEMP2.IO_KB, "a")                      'io�敪
    Call UniCode_Conv(K0_ODR_TEMP2.USE_YM, GW_TOUGETU)              '�g�p��
    Call UniCode_Conv(K0_ODR_TEMP2.ANS_NOUKI_DT, "")                   '�Ώۓ��t   YYYYMMDD
    Call UniCode_Conv(K0_ODR_TEMP2.ORDER_NO, "")                    '������
    
    Do
        sts = BTRV(BtOpGetEqual, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, Len(K0_ODR_TEMP2), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                'MsgBox "�w�肳�ꂽ�H��������܂���B"
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                Sleep (500)
            Case Else
                Call File_Error(sts, BtOpGetEqual, "ODR_TEMP2")
                Exit Function
        End Select
    Loop
    
    If sts = BtNoErr Then           '�o�^�ς�
        SET_I_ZAIKO = False
        Exit Function
    End If
    
    W_Zaiko = 0
    
            '-------------------------------------------------- '���݌Ɋl��
            
    Call UniCode_Conv(K0_ITEM.JGYOBU, GW_JIGYOBU_KO)     '���ƕ�
    Call UniCode_Conv(K0_ITEM.NAIGAI, GW_NAIGAI_KO)      '�����O
    Call UniCode_Conv(K0_ITEM.HIN_GAI, GW_HINGAI_KO)     '�q�i��
    
    Do
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                Sleep (500)
            Case Else
                Call File_Error(sts, BtOpGetEqual, "ITEM")
                'Exit Function
        End Select
    Loop
    
    If sts = BtNoErr Then
                
        If Trim(StrConv(ITEMREC.G_ZEN_ZAIKO_QTY, vbUnicode)) <> "" Then
            If IsNumeric(StrConv(ITEMREC.G_ZEN_ZAIKO_QTY, vbUnicode)) Then
                W_Zaiko = CDbl(StrConv(ITEMREC.G_ZEN_ZAIKO_QTY, vbUnicode))
            End If
        End If
        
    End If
                        

    Call ODR_TEMP2_CLR
    
        '�q�@���ƕ�
    Call UniCode_Conv(ODR_TP2_R.KO_JGYOBU, GW_JIGYOBU_KO)
        '�q�@�����O
    Call UniCode_Conv(ODR_TP2_R.KO_NAIGAI, GW_NAIGAI_KO)
        '�q�i��
    Call UniCode_Conv(ODR_TP2_R.KO_HIN_GAI, GW_HINGAI_KO)
        'io�敪
    Call UniCode_Conv(ODR_TP2_R.IO_KB, "a")
        '�g�p��
    Call UniCode_Conv(ODR_TP2_R.USE_YM, GW_TOUGETU)
        '�[��
    Call UniCode_Conv(ODR_TP2_R.ANS_NOUKI_DT, "")
        '������
    Call UniCode_Conv(ODR_TP2_R.ORDER_NO, "")
        

    W_STR = CStr(W_Zaiko)
    If Trim(W_STR) = "" Then W_STR = "0"
        
    Call UniCode_Conv(ODR_TP2_R.ZAI_QTY, W_STR)     '�݌ɐ�
    Call UniCode_Conv(ODR_TP2_R.MOTO_QTY, W_STR)    '���X�̍݌ɐ�
        
    Call UniCode_Conv(ODR_TP2_R.UPDT_DT, Right(Format(Date, "yyyymmdd"), 6))
    Call UniCode_Conv(ODR_TP2_R.UPDT_TM, Left(Format(Time, "hhmmss"), 4))
    Do
        sts = BTRV(BtOpInsert, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, Len(K0_ODR_TEMP2), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                Sleep (500)
            Case Else
                Call File_Error(sts, BtOpInsert, "ODR_TEMP2")
                Exit Function
        End Select
    Loop
    
    SET_I_ZAIKO = False
    
End Function

Function SET_ODR_ZAN() As Integer

        '       �������i���ɗ\��j���@�������o��    io�敪����


Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim W_Sw        As Integer

Dim W_STR       As String
Dim W_Edit      As String
Dim W_Zaiko     As Double
Dim W_QTY       As Double

Dim W_Date      As String       '�����̂P��
Dim W_Today     As String       '�{��
Dim wkYYMMDD     As String

Dim W_ZENGETU   As String


Dim W_Kan_DT    As String

    SET_ODR_ZAN = True
    
    
    W_Today = Format(Date, "yyyymmdd")
    W_Date = Left(W_Today, 6) & "01"
    W_STR = Left(W_Date, 4) & "/" & Mid(W_Date, 5, 2) & "/" & Right(W_Date, 2)
    
    
    wkYYMMDD = Left(GW_TOUGETU, 4) & "/" & Mid(GW_TOUGETU, 5, 2) & "/01"
    
    W_ZENGETU = Left(Format(DateAdd("d", -1, wkYYMMDD), "yyyymmdd"), 6) & "01"
    
    
    W_Zaiko = 0
    Call UniCode_Conv(K1_P_SHORDER.JGYOBU, GW_JIGYOBU_KO)
    Call UniCode_Conv(K1_P_SHORDER.NAIGAI, GW_NAIGAI_KO)
    Call UniCode_Conv(K1_P_SHORDER.HIN_GAI, GW_HINGAI_KO)
    Call UniCode_Conv(K1_P_SHORDER.ORDER_DT, "")
    Call UniCode_Conv(K1_P_SHORDER.ORDER_NO, "")
    com = BtOpGetGreaterEqual
    Do
        yn = 0
        sts = BTRV(com, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K1_P_SHORDER, Len(K1_P_SHORDER), 1)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                    
            Case Else
                Call File_Error(sts, com, "P_SHORDER")
                Exit Function
        End Select
        If sts <> BtNoErr Then Exit Do
        
        If Trim(StrConv(P_SHORDER_REC.JGYOBU, vbUnicode)) <> Trim(GW_JIGYOBU_KO) Then Exit Do
        If Trim(StrConv(P_SHORDER_REC.NAIGAI, vbUnicode)) <> Trim(GW_NAIGAI_KO) Then Exit Do
        If Trim(StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode)) <> Trim(GW_HINGAI_KO) Then Exit Do
           
        '   2008/09 �p�`�̇��P�R�ɂ��~�ρI
        If Trim(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode)) = "" Then
            Call UniCode_Conv(P_SHORDER_REC.ANS_NOUKI_DT, StrConv(P_SHORDER_REC.KAN_DT, vbUnicode))
        End If


If Trim(GW_HINGAI_KO) = "C215" Then
                W_Sw = True
End If
               
        W_Sw = True
               
        If StrConv(P_SHORDER_REC.CANCEL_F, vbUnicode) = "1" Then W_Sw = False   '�L�����Z���H
            
            
            '   �g�p���F���ݒ�@���@�ΏۊO�I
        If Trim(StrConv(P_SHORDER_REC.USE_YM, vbUnicode)) = "" Then W_Sw = False
          
          
            '   �g�p�� �� ����@���@�ΏۊO�I
        If Trim(StrConv(P_SHORDER_REC.USE_YM, vbUnicode)) < GW_TOUGETU Then W_Sw = False
        
        
            '   �g�p���@���@�Q�O�P���@���@�ΏۊO�I  2008/12/02
        If Trim(StrConv(P_SHORDER_REC.USE_YM, vbUnicode)) <> "" Then
            W_STR = Left(StrConv(P_SHORDER_REC.USE_YM, vbUnicode), 4) & "/" & _
                        Right(StrConv(P_SHORDER_REC.USE_YM, vbUnicode), 2)
            If W_STR > GW_MAX_YYMM Then
                W_Sw = False
            End If
       End If
       
       
       '                2008.12.17 ����F����@�ǉ�
       If StrConv(P_SHORDER_REC.KAN_F, vbUnicode) = "1" Then W_Sw = False
       
            '   ������      '2008/09/10 �������̂ݑΏہI�H
        'If Trim(StrConv(P_SHORDER_REC.KAN_DT, vbUnicode)) <> "" Then
        '    W_Sw = False
        'End If
                       
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        'If W_Sw = True Then
        '    '2008/09/13�Ƃɂ����A�Ώې��͒������ɓ���I�i�p���`�̇�10���j
        '    W_Zaiko = CDbl(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode))
        '    '08.12.04   �d���c�̌v�Z�ɕύX�I                                '08.12.04
        '    '               �A���A���O�̏ꍇ�́u�O�i�[���j�v�Ƃ���B        '08.12.04
        '    W_Zaiko = CDbl(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)) - _
        '                CDbl(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode))
        '    If W_Zaiko < 0 Then
        '        W_Zaiko = 0
        '    End If
        '    If StrConv(P_SHORDER_REC.KAN_F, vbUnicode) <> "1" Then   '������
        '        '                            '�������@���@�����
        '        'If CDbl(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)) > CDbl(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)) Then
        '        '
        '        '    W_Zaiko = W_Zaiko + CDbl(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode))
        '        'Else
        '        '                            '�������@���@�����
        '        '    W_Zaiko = W_Zaiko + CDbl(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode))
        '        'End If
        '    Else
        '                                    '�����F�����
        '        'W_Zaiko = W_Zaiko + CDbl(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode))
        '
        '        '2008/11/22�u���� �� �g�p�������� �� ��������J�z���v�͎���ς݁������݌ɍ݌ɁI
        '        If Trim(StrConv(P_SHORDER_REC.USE_YM, vbUnicode)) = GW_TOUGETU Then
        '            'If Trim(StrConv(P_SHORDER_REC.KAN_DT, vbUnicode)) >= GW_SHIMEBI Then
        '            '2008.11.29                                     �s�������t�I            (*_*;
        '            If Trim(StrConv(P_SHORDER_REC.KAN_DT, vbUnicode)) < GW_SHIMEBI Then
        '                W_Zaiko = 0
        '            End If
        '        End If
        '        '2008.12.02 �d�������͓��t���薳���ŁA�����\�݌ɂƂ��Ȃ��I
        '        W_Zaiko = 0
        '
        '    End If
        'End If
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        '       2008.12.04��L�u���b�N�����L�ɕύX�I
        '
        '               �d���c���������������܂ł̎���������Z�B
        '
        W_Zaiko = 0
        If W_Sw = True Then
            W_Zaiko = CDbl(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode))
            
            Call UniCode_Conv(K0_P_SHUKEIRE.ORDER_NO, StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode))
            Call UniCode_Conv(K0_P_SHUKEIRE.SEQNO, "")
            com = BtOpGetGreaterEqual
            Do
                Do
                    sts = BTRV(com, P_SHUKEIRE_POS, P_SHUKEIRE_REC, Len(P_SHUKEIRE_REC), K0_P_SHUKEIRE, Len(K0_P_SHUKEIRE), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                            Exit Do
                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                            Sleep (500)
                        Case Else
                            Call File_Error(sts, com, "P_SHUKEIRE")
                            Exit Function
                    End Select
                Loop
                If sts <> BtNoErr Then Exit Do
                
                '   ����f�[�^�̌v�㌎�𔻒�ɉ����������H
                
                If Trim(StrConv(P_SHUKEIRE_REC.ORDER_NO, vbUnicode)) <> _
                        Trim(StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode)) Then Exit Do
                    
                
                If StrConv(P_SHUKEIRE_REC.KEIJYO_YM, vbUnicode) <= GW_TOUGETU Then
                    W_Zaiko = W_Zaiko - CDbl(StrConv(P_SHUKEIRE_REC.UKEIRE_QTY, vbUnicode))
                End If
                
                
                If W_Zaiko <= 0 Then Exit Do
                
                com = BtOpGetNext
            Loop
            
        End If
                    '   �����܂ŁI
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        
        If W_Zaiko < 0 Then
            W_Zaiko = 0
        End If
        
        If W_Zaiko <> 0 Then
        
            Call UniCode_Conv(K0_ODR_TEMP2.KO_JGYOBU, GW_JIGYOBU_KO)    '�q�@���ƕ�
            Call UniCode_Conv(K0_ODR_TEMP2.KO_NAIGAI, GW_NAIGAI_KO)     '�q�@�����O
            Call UniCode_Conv(K0_ODR_TEMP2.KO_HIN_GAI, GW_HINGAI_KO)    '�q�i��
            
            If Trim(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode)) = "" Then
                Call UniCode_Conv(K0_ODR_TEMP2.IO_KB, "g")                  'io�敪
            Else
                Call UniCode_Conv(K0_ODR_TEMP2.IO_KB, "f")                  'io�敪
            End If
            
            W_STR = StrConv(P_SHORDER_REC.USE_YM, vbUnicode)
            Call UniCode_Conv(K0_ODR_TEMP2.USE_YM, W_STR)               '�g�p��
                
            W_STR = Trim(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode))
            Call UniCode_Conv(K0_ODR_TEMP2.ANS_NOUKI_DT, W_STR)            '������
                
            W_STR = StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode)
            Call UniCode_Conv(K0_ODR_TEMP2.ORDER_NO, W_STR)             '������
                
            sts = BTRV(BtOpGetEqual, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, Len(K0_ODR_TEMP2), 0)
            Select Case sts
                Case BtNoErr
                    com = BtOpUpdate
                Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                        
                    Call ODR_TEMP2_CLR
                    com = BtOpInsert
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "ODR_TEMP2")
                    Exit Function
            End Select
                
                '�q�@���ƕ�
            Call UniCode_Conv(ODR_TP2_R.KO_JGYOBU, GW_JIGYOBU_KO)
                '�q�@�����O
            Call UniCode_Conv(ODR_TP2_R.KO_NAIGAI, GW_NAIGAI_KO)
                '�q�i��
            Call UniCode_Conv(ODR_TP2_R.KO_HIN_GAI, GW_HINGAI_KO)
            
            
                                                '2008/09 �񓚔[���̗L���ŋ敪���قȂ�I
                'io�敪         2008.05.02
            If Trim(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode)) = "" Then
                Call UniCode_Conv(ODR_TP2_R.IO_KB, "g")                  'io�敪
            Else
                Call UniCode_Conv(ODR_TP2_R.IO_KB, "f")                  'io�敪
            End If
                            
                
            W_STR = StrConv(P_SHORDER_REC.USE_YM, vbUnicode)
            Call UniCode_Conv(ODR_TP2_R.USE_YM, W_STR)                  '�g�p��
                    
If Trim(GW_HINGAI_KO) = "C029" Then
    Debug.Print
End If
                    
            W_STR = StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode)
                '�[��
            Call UniCode_Conv(ODR_TP2_R.ANS_NOUKI_DT, W_STR)
                '������
            Call UniCode_Conv(ODR_TP2_R.ORDER_NO, StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode))
                
                
            W_QTY = W_Zaiko + CDbl(StrConv(ODR_TP2_R.ZAI_QTY, vbUnicode))
                                
                                
            W_STR = CStr(W_QTY)
                
            Call UniCode_Conv(ODR_TP2_R.ZAI_QTY, W_STR)
            Call UniCode_Conv(ODR_TP2_R.MOTO_QTY, W_STR)
                
            Call UniCode_Conv(ODR_TP2_R.UPDT_DT, Right(Format(Date, "yyyymmdd"), 6))
            Call UniCode_Conv(ODR_TP2_R.UPDT_TM, Left(Format(Time, "hhmmss"), 4))
            
            If com <> BtOpUpdate Then
                Do
                    sts = BTRV(com, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, Len(K0_ODR_TEMP2), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                            Sleep (500)
                        Case Else
                            Call File_Error(sts, com, "ODR_TEMP2")
                            Exit Function
                    End Select
                Loop
            Else
                '   ����́A��������H�@(��;)
                W_QTY = CDbl(StrConv(ODR_TP2_R.ZAI_QTY, vbUnicode))
                
            End If
            
            
        End If
    
    
        W_Zaiko = 0
        com = BtOpGetNext
    Loop
    
    
    SET_ODR_ZAN = False
    
End Function

Function SET_UKEIRE() As Integer

        '       ������я��@�������o��        io�敪����

Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim W_Sw        As Integer

Dim W_STR       As String
Dim W_Zaiko     As Double

Dim W_Moto      As Double

    SET_UKEIRE = True
    
    W_Zaiko = 0
    
    Call UniCode_Conv(K1_P_SHUKEIRE.KEIJYO_YM, GW_TOUGETU)
    Call UniCode_Conv(K1_P_SHUKEIRE.ORDER_CODE, "")
    Call UniCode_Conv(K1_P_SHUKEIRE.UKEIRE_DT, "")
    
    com = BtOpGetGreaterEqual
    Do
        yn = 0
        sts = BTRV(com, P_SHUKEIRE_POS, P_SHUKEIRE_REC, Len(P_SHUKEIRE_REC), K1_P_SHUKEIRE, Len(K1_P_SHUKEIRE), 1)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                    
            Case Else
                Call File_Error(sts, com, "P_SHORDER")
                Exit Function
        End Select
        If sts <> BtNoErr Then Exit Do
        
        If Trim(StrConv(P_SHUKEIRE_REC.ORDER_NO, vbUnicode)) = "05975" Then
            sts = 0
        End If
        
                
                        '   �v��N���@���@�����@���@�I��
        If StrConv(P_SHUKEIRE_REC.KEIJYO_YM, vbUnicode) <> GW_TOUGETU Then Exit Do
        
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        'W_STR = Trim(StrConv(P_SHUKEIRE_REC.UKEIRE_DT, vbUnicode))
        '
        ''2008.12.02
        ''                   ����N���@�� �@�J�z�N���H
        'If Left(W_STR, 6) <> Left(GW_SHIMEBI, 6) Then
        '&'   W_STR = ""
        'End If
        '
        '                    '   ������@<=�@�J�z���H
        'If W_STR <= GW_SHIMEBI Then
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        '                                   2008/12/06  ��L�u���b�N�̎�����t����s�v�I�Ƃ����B
        
                        '   �������m�F
            Call UniCode_Conv(K0_P_SHORDER.ORDER_NO, StrConv(P_SHUKEIRE_REC.ORDER_NO, vbUnicode))
            
            sts = BTRV(BtOpGetEqual, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                        
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "P_SHORDER")
                    Exit Function
            End Select
            
            If sts = BtNoErr Then
                W_Sw = True
                If Trim(StrConv(P_SHORDER_REC.JGYOBU, vbUnicode)) <> Trim(GW_JIGYOBU_KO) Then
                    W_Sw = False
                End If
                
                If Trim(StrConv(P_SHORDER_REC.NAIGAI, vbUnicode)) <> Trim(GW_NAIGAI_KO) Then
                    W_Sw = False
                End If
                                '�L�����Z���H
                If StrConv(P_SHORDER_REC.CANCEL_F, vbUnicode) = "1" Then
                    W_Sw = False
                End If
                
                If Trim(StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode)) = "C215" Then
                    
                    W_STR = StrConv(P_SHUKEIRE_REC.UKEIRE_DT, vbUnicode)
                    sts = 0
                End If
                            '���ƕ��A�����O����v�@���@�W�J�f�[�^���̕i�ڂ̗L���m�F
                If W_Sw = True Then
                    Call UniCode_Conv(K2_ODR_TEMP1.KO_JGYOBU, StrConv(P_SHORDER_REC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(K2_ODR_TEMP1.KO_NAIGAI, StrConv(P_SHORDER_REC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(K2_ODR_TEMP1.KO_HIN_GAI, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))
                    Call UniCode_Conv(K2_ODR_TEMP1.SHIMUKE, "")
                    Call UniCode_Conv(K2_ODR_TEMP1.JGYOBU, "")
                    Call UniCode_Conv(K2_ODR_TEMP1.NAIGAI, "")
                    Call UniCode_Conv(K2_ODR_TEMP1.HIN_GAI, "")
                    Call UniCode_Conv(K2_ODR_TEMP1.ORDER_NO, "")
                    Call UniCode_Conv(K2_ODR_TEMP1.INS_NO, "")
                    Call UniCode_Conv(K2_ODR_TEMP1.BUN_NO, "")
        
                    sts = BTRV(BtOpGetGreaterEqual, ODR_TP1_POS, ODR_TP1_R, Len(ODR_TP1_R), _
                                                K2_ODR_TEMP1, Len(K2_ODR_TEMP1), 2)
                    Select Case sts
                        Case BtNoErr
                        
                        Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                        
                        
                        Case Else
                            Call File_Error(sts, BtOpGetGreaterEqual, "ODR_TEMP1")
                            'Exit Function
                    End Select
                    If sts <> BtNoErr Then W_Sw = False
                    
                    If Trim(StrConv(ODR_TP1_R.KO_JGYOBU, vbUnicode)) <> Trim(GW_JIGYOBU_KO) Then
                        W_Sw = False
                    End If
                    
                    If Trim(StrConv(ODR_TP1_R.KO_NAIGAI, vbUnicode)) <> Trim(GW_NAIGAI_KO) Then
                        W_Sw = False
                    End If
                    
                    If Trim(StrConv(ODR_TP1_R.KO_HIN_GAI, vbUnicode)) <> _
                                    Trim(StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode)) Then
                        W_Sw = False
                    End If
                    
        
                    '>>>>>>>>>>>>>>>>>>>>   ������сi������j�̗݌v
                    If W_Sw Then
                        W_Zaiko = CDbl(StrConv(P_SHUKEIRE_REC.UKEIRE_QTY, vbUnicode))
                        
                        If Trim(StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode)) = "C215" Then
                            W_Zaiko = W_Zaiko * 1
                        End If
                            
                        Call UniCode_Conv(K0_ODR_TEMP2.KO_JGYOBU, GW_JIGYOBU_KO)    '�q�@���ƕ�
                        Call UniCode_Conv(K0_ODR_TEMP2.KO_NAIGAI, GW_NAIGAI_KO)     '�q�@�����O
                        Call UniCode_Conv(K0_ODR_TEMP2.KO_HIN_GAI, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))    '�q�i��
                        Call UniCode_Conv(K0_ODR_TEMP2.IO_KB, "c")                  'io�敪
                                        
                        Call UniCode_Conv(K0_ODR_TEMP2.USE_YM, GW_TOUGETU)               '�g�p��
                            
                        Call UniCode_Conv(K0_ODR_TEMP2.ANS_NOUKI_DT, "")            '������
                            
                        Call UniCode_Conv(K0_ODR_TEMP2.ORDER_NO, "")             '������
                            
                        sts = BTRV(BtOpGetEqual, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, _
                                            Len(K0_ODR_TEMP2), 0)
                        Select Case sts
                            Case BtNoErr
                                com = BtOpUpdate
                            Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                                    
                                Call ODR_TEMP2_CLR
                                com = BtOpInsert
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "ODR_TEMP2")
                                Exit Function
                        End Select
                            
                            '�q�@���ƕ�
                        Call UniCode_Conv(ODR_TP2_R.KO_JGYOBU, GW_JIGYOBU_KO)
                            '�q�@�����O
                        Call UniCode_Conv(ODR_TP2_R.KO_NAIGAI, GW_NAIGAI_KO)
                            '�q�i��
                        Call UniCode_Conv(ODR_TP2_R.KO_HIN_GAI, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))
                            'io�敪         2008.05.02
                        Call UniCode_Conv(ODR_TP2_R.IO_KB, "c")
                            '�g�p��
                        Call UniCode_Conv(ODR_TP2_R.USE_YM, GW_TOUGETU)
                            '�[��
                        Call UniCode_Conv(ODR_TP2_R.ANS_NOUKI_DT, "")
                            '������
                        Call UniCode_Conv(ODR_TP2_R.ORDER_NO, "")
                           
                        W_Moto = CDbl(StrConv(ODR_TP2_R.MOTO_QTY, vbUnicode))
                        W_STR = CStr(W_Zaiko + W_Moto)
                            
                        Call UniCode_Conv(ODR_TP2_R.ZAI_QTY, W_STR)
                        Call UniCode_Conv(ODR_TP2_R.MOTO_QTY, W_STR)
                            
                        Call UniCode_Conv(ODR_TP2_R.UPDT_DT, Right(Format(Date, "yyyymmdd"), 6))
                        Call UniCode_Conv(ODR_TP2_R.UPDT_TM, Left(Format(Time, "hhmmss"), 4))
                            
                        Do
                            sts = BTRV(com, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, Len(K0_ODR_TEMP2), 0)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                                    Sleep (500)
                                Case Else
                                    Call File_Error(sts, com, "ODR_TEMP2")
                                    Exit Function
                            End Select
                        Loop
            
                    End If
                    
                End If
                
                
            End If
            
        'End If
        
        
        com = BtOpGetNext
    Loop
    

    SET_UKEIRE = False

End Function

Function SET_ZAITEI() As Integer

        '       �ݒ��i�}�j���@�������o��        io�敪����
        
        '       �ړ�����茟��
        
Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim W_Sw        As Integer

Dim W_STR       As String
Dim W_Zaiko     As Double
Dim W_QTY       As Double

Dim W_Date      As String       '�����̂P��
Dim W_Today     As String       '�{��
Dim wkYYMMDD     As String

Dim W_ZENGETU   As String


Dim W_Kan_DT    As String

Dim X_i         As Integer

    SET_ZAITEI = True
    
    W_Zaiko = 0
    
    
    W_Today = Format(Date, "yyyymmdd")                  '�{���iPC-DATE)
    
    Call UniCode_Conv(K1_IDO.JGYOBU, GW_JIGYOBU_KO)
    Call UniCode_Conv(K1_IDO.NAIGAI, GW_NAIGAI_KO)
    Call UniCode_Conv(K1_IDO.HIN_GAI, GW_HINGAI_KO)
    Call UniCode_Conv(K1_IDO.JITU_DT, GW_SHIMEBI)       '�ΏہF�J�z���ȍ~�I
    Call UniCode_Conv(K1_IDO.JITU_TM, "")
    com = BtOpGetGreaterEqual
    
    Do
        sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                    
            Case Else
                Call File_Error(sts, com, "IDO")
                Exit Function
        End Select
        
        
        If sts <> BtNoErr Then Exit Do
        
        If Trim(StrConv(IDOREC.JGYOBU, vbUnicode)) <> Trim(GW_JIGYOBU_KO) Then Exit Do
        If Trim(StrConv(IDOREC.NAIGAI, vbUnicode)) <> Trim(GW_NAIGAI_KO) Then Exit Do
        If Trim(StrConv(IDOREC.HIN_GAI, vbUnicode)) <> Trim(GW_HINGAI_KO) Then Exit Do
        
        
        If Trim(StrConv(IDOREC.JITU_DT, vbUnicode)) > W_Today Then Exit Do
        
        
    'SUMI_JITU_QTY(0 To 7)               As Byte     '���ѐ���(���i���ς�)
    'MI_JITU_QTY(0 To 7)                 As Byte     '���ѐ���(�����i)
        
        'W_QTY = 0
        'Select Case StrConv(IDOREC.RIRK_ID, vbUnicode)
        '    Case GW_PURA
        '        If IsNumeric(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) Then
        '            W_QTY = CDbl(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode))
        '        End If
        '        If IsNumeric(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)) Then
        '            W_QTY = W_QTY + CDbl(StrConv(IDOREC.MI_JITU_QTY, vbUnicode))
        '        End If
        '
        '    Case GW_MAINA
        '        If IsNumeric(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) Then
        '            W_QTY = CDbl(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode))
        '        End If
        '        If IsNumeric(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)) Then
        '            W_QTY = W_QTY + CDbl(StrConv(IDOREC.MI_JITU_QTY, vbUnicode))
        '        End If
        '
        '        W_QTY = W_QTY * -1
        '
        '    Case Else
        '        W_QTY = 0
        '
        'End Select
        '               2009/03/04  GW_PURA,GW_MAINA���e�[�u���ɂ����I
        W_QTY = 0
        For X_i = 0 To UBound(GW_PURA)
            If GW_PURA(X_i) <> "" Then
                If StrConv(IDOREC.RIRK_ID, vbUnicode) = GW_PURA(X_i) Then
                    If IsNumeric(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) Then
                        W_QTY = CDbl(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode))
                    End If
                    If IsNumeric(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)) Then
                        W_QTY = W_QTY + CDbl(StrConv(IDOREC.MI_JITU_QTY, vbUnicode))
                    End If
                End If
            End If
        Next X_i
        
        For X_i = 0 To UBound(GW_MAINA)
            If GW_MAINA(X_i) <> "" Then
                If StrConv(IDOREC.RIRK_ID, vbUnicode) = GW_MAINA(X_i) Then
                    If IsNumeric(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) Then
                        W_QTY = W_QTY - CDbl(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode))
                    End If
                    If IsNumeric(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)) Then
                        W_QTY = W_QTY - CDbl(StrConv(IDOREC.MI_JITU_QTY, vbUnicode))
                    End If
                End If
            End If
        Next X_i
        
        
        W_Zaiko = W_Zaiko + W_QTY
        
        com = BtOpGetNext
        
    Loop
    
    If W_Zaiko <> 0 Then
    
        
            Call UniCode_Conv(K0_ODR_TEMP2.KO_JGYOBU, GW_JIGYOBU_KO)    '�q�@���ƕ�
            Call UniCode_Conv(K0_ODR_TEMP2.KO_NAIGAI, GW_NAIGAI_KO)     '�q�@�����O
            Call UniCode_Conv(K0_ODR_TEMP2.KO_HIN_GAI, GW_HINGAI_KO)    '�q�i��
            Call UniCode_Conv(K0_ODR_TEMP2.IO_KB, "d")                  'io�敪
                            
            Call UniCode_Conv(K0_ODR_TEMP2.USE_YM, GW_TOUGETU)               '�g�p��
                
            Call UniCode_Conv(K0_ODR_TEMP2.ANS_NOUKI_DT, "")            '������
                
            Call UniCode_Conv(K0_ODR_TEMP2.ORDER_NO, "")             '������
                
            sts = BTRV(BtOpGetEqual, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, Len(K0_ODR_TEMP2), 0)
            Select Case sts
                Case BtNoErr
                    com = BtOpUpdate
                Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                        
                    Call ODR_TEMP2_CLR
                    com = BtOpInsert
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "ODR_TEMP2")
                    Exit Function
            End Select
                
                '�q�@���ƕ�
            Call UniCode_Conv(ODR_TP2_R.KO_JGYOBU, GW_JIGYOBU_KO)
                '�q�@�����O
            Call UniCode_Conv(ODR_TP2_R.KO_NAIGAI, GW_NAIGAI_KO)
                '�q�i��
            Call UniCode_Conv(ODR_TP2_R.KO_HIN_GAI, GW_HINGAI_KO)
                'io�敪         2008.05.02
            Call UniCode_Conv(ODR_TP2_R.IO_KB, "d")
                '�g�p��
            Call UniCode_Conv(ODR_TP2_R.USE_YM, GW_TOUGETU)
                '�[��
            Call UniCode_Conv(ODR_TP2_R.ANS_NOUKI_DT, "")
                '������
            Call UniCode_Conv(ODR_TP2_R.ORDER_NO, "")
                
            W_STR = CStr(W_Zaiko)
                
            Call UniCode_Conv(ODR_TP2_R.ZAI_QTY, W_STR)
            Call UniCode_Conv(ODR_TP2_R.MOTO_QTY, W_STR)
                
            Call UniCode_Conv(ODR_TP2_R.UPDT_DT, Right(Format(Date, "yyyymmdd"), 6))
            Call UniCode_Conv(ODR_TP2_R.UPDT_TM, Left(Format(Time, "hhmmss"), 4))
            If com <> BtOpUpdate Then
                Do
                    sts = BTRV(com, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, Len(K0_ODR_TEMP2), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                            Sleep (500)
                        Case Else
                            Call File_Error(sts, com, "ODR_TEMP2")
                            Exit Function
                    End Select
                Loop
            End If
            
    End If
    SET_ZAITEI = False

End Function


Function SET_H_SEIHIN() As Integer
                    '�����i���̏W��i�L�[�ŏW�v����j 'io�敪����

Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim W_Sw        As Integer

Dim W_STR       As String
Dim W_Zaiko     As Double
Dim W_QTY       As Double


    SET_H_SEIHIN = True
    
    
    Call ODR_TEMP2_CLR
    
    W_Zaiko = 0
        
    Call UniCode_Conv(K1_ODR_HANSEIHIN.KO_JGYOBU, GW_JIGYOBU_KO)
    Call UniCode_Conv(K1_ODR_HANSEIHIN.KO_NAIGAI, GW_NAIGAI_KO)
    Call UniCode_Conv(K1_ODR_HANSEIHIN.KO_HIN_GAI, GW_HINGAI_KO)
    
    If Trim(GW_HINGAI_KO) = "AD-HESB66AZ" Then
        W_Zaiko = 0
    End If
    
    com = BtOpGetGreaterEqual
    
    Do
                                            
        sts = BTRV(com, ODR_HANSEIHIN_POS, ODR_HANSEIHIN_K_REC, Len(ODR_HANSEIHIN_K_REC), _
                            K1_ODR_HANSEIHIN, Len(K1_ODR_HANSEIHIN), 1)
                            
                            
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                    
            Case Else
                Call File_Error(sts, com, "ODR_HANSEIHIN")
                Exit Do
        End Select
                
        If sts <> BtNoErr Then Exit Do
        
        If Trim(StrConv(ODR_HANSEIHIN_K_REC.KO_JGYOBU, vbUnicode)) > Trim(GW_JIGYOBU_KO) Then Exit Do
        If Trim(StrConv(ODR_HANSEIHIN_K_REC.KO_NAIGAI, vbUnicode)) > Trim(GW_NAIGAI_KO) Then Exit Do
        If Trim(StrConv(ODR_HANSEIHIN_K_REC.KO_HIN_GAI, vbUnicode)) > Trim(GW_HINGAI_KO) Then Exit Do
        
        W_Sw = True
        
        If Trim(StrConv(ODR_HANSEIHIN_K_REC.KO_JGYOBU, vbUnicode)) <> Trim(GW_JIGYOBU_KO) Then
            W_Sw = False
        End If
        
        If Trim(StrConv(ODR_HANSEIHIN_K_REC.KO_NAIGAI, vbUnicode)) <> Trim(GW_NAIGAI_KO) Then
            W_Sw = False
        End If
        
        If Trim(StrConv(ODR_HANSEIHIN_K_REC.KO_HIN_GAI, vbUnicode)) <> Trim(GW_HINGAI_KO) Then
            W_Sw = False
        End If
        
        
                    '2008/11/11 ���t�𔻒肵�Ȃ��I
''''        If Trim(StrConv(ODR_HANSEIHIN_K_REC.USE_YM, vbUnicode)) <> Trim(GW_TOUGETU) Then
''''            W_Sw = False
''''        End If
        
        
        If StrConv(ODR_HANSEIHIN_K_REC.SEQNO, vbUnicode) = "000" Then
            W_Sw = False            '�e���R�[�h
        Else
            If StrConv(ODR_HANSEIHIN_K_REC.ZAITEI_F, vbUnicode) = "1" Then
                W_Sw = False        '���݌ɓo�^�ς�
            End If
        End If
        
                
'        If CDbl(StrConv(ODR_HANSEIHIN_K_REC.USE_QTY, vbUnicode)) <= 0 Then
'            W_Sw = False            '�����i�̖߂�
'        End If
        
        W_QTY = 0
                                            '������     9(5)v9(2)
        If W_Sw = True Then
            If Trim(StrConv(ODR_HANSEIHIN_K_REC.USE_QTY, vbUnicode)) <> "" Then
                If IsNumeric(Trim(StrConv(ODR_HANSEIHIN_K_REC.USE_QTY, vbUnicode))) Then
                    W_QTY = CDbl(StrConv(ODR_HANSEIHIN_K_REC.USE_QTY, vbUnicode))
                End If
            End If
        End If
        W_Zaiko = W_Zaiko + W_QTY
            
        com = BtOpGetNext
    Loop
               
    If W_Zaiko <> 0 Then
                      
        Call UniCode_Conv(K0_ODR_TEMP2.KO_JGYOBU, GW_JIGYOBU_KO)   '�q�@���ƕ�
        Call UniCode_Conv(K0_ODR_TEMP2.KO_NAIGAI, GW_NAIGAI_KO)     '�q�@�����O
        Call UniCode_Conv(K0_ODR_TEMP2.KO_HIN_GAI, GW_HINGAI_KO)    '�q�i��
        Call UniCode_Conv(K0_ODR_TEMP2.IO_KB, "e")                          'io�敪 2008.05.02
        Call UniCode_Conv(K0_ODR_TEMP2.USE_YM, GW_TOUGETU)              '�g�p��
            
        Call UniCode_Conv(K0_ODR_TEMP2.ANS_NOUKI_DT, "")           '������
        Call UniCode_Conv(K0_ODR_TEMP2.ORDER_NO, "")            '������
            
            
            
        sts = BTRV(BtOpGetEqual, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, Len(K0_ODR_TEMP2), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
            Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                    'MsgBox "�w�肳�ꂽ�H��������܂���B"
                    
                Call ODR_TEMP2_CLR
                com = BtOpInsert
            Case Else
                Call File_Error(sts, BtOpGetEqual, "ODR_TEMP2")
                Exit Function
        End Select
            
            '�q�@���ƕ�
        Call UniCode_Conv(ODR_TP2_R.KO_JGYOBU, GW_JIGYOBU_KO)   '�q�@���ƕ�
            '�q�@�����O
        Call UniCode_Conv(ODR_TP2_R.KO_NAIGAI, GW_NAIGAI_KO)     '�q�@�����O
            '�q�i��
        Call UniCode_Conv(ODR_TP2_R.KO_HIN_GAI, GW_HINGAI_KO)    '�q�i��
            'io�敪
        Call UniCode_Conv(ODR_TP2_R.IO_KB, "e")
            '�g�p��
        Call UniCode_Conv(ODR_TP2_R.USE_YM, GW_TOUGETU)
            
            '2008/05/31 �[��   �����I�ɕύX�B
        Call UniCode_Conv(ODR_TP2_R.ANS_NOUKI_DT, "")
            
            '2008/05/31 �������͖����I�ɕύX�B
        Call UniCode_Conv(ODR_TP2_R.ORDER_NO, "")
        
        W_Zaiko = W_Zaiko + CDbl(Trim(StrConv(ODR_TP2_R.ZAI_QTY, vbUnicode)))
        W_STR = CStr(W_Zaiko)
    
        Call UniCode_Conv(ODR_TP2_R.ZAI_QTY, W_STR)
        Call UniCode_Conv(ODR_TP2_R.MOTO_QTY, W_STR)
            
        Call UniCode_Conv(ODR_TP2_R.UPDT_DT, Right(Format(Date, "yyyymmdd"), 6))
        Call UniCode_Conv(ODR_TP2_R.UPDT_TM, Left(Format(Time, "hhmmss"), 4))
            
        Do
            sts = BTRV(com, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, Len(K0_ODR_TEMP2), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                    Sleep (500)
                Case Else
                    Call File_Error(sts, com, "ODR_TEMP2")
                    Exit Function
            End Select
        Loop
            
    End If
        
    SET_H_SEIHIN = False


End Function




Function ZAN_CALC() As Integer

        '       �݌ɁA���v�ʁ�������񂪕\�����ꂽ�̂ŁA���t���ɍ��������c����ݒ�B

Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim W_STR       As String

Dim W_QTY       As Double
Dim W_Date      As String
Dim W_NOW       As String

    ZAN_CALC = True
    
    W_NOW = GW_TOUGETU & "01"
    com = BtOpGetFirst
    Do
        sts = BTRV(com, ODR_TP1_POS, ODR_TP1_R, Len(ODR_TP1_R), K3_ODR_TEMP1, Len(K3_ODR_TEMP1), 3)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                'Beep
                'MsgBox "�w�肳�ꂽ�H��������܂���B"
            Case Else
                Call File_Error(sts, com, "ODR_TEMP1")
                Exit Function
        End Select
        If sts <> BtNoErr Then Exit Do
        
                        '��Trim 2008/07/02�ǉ�
        GW_JIGYOBU_KO = Trim(StrConv(ODR_TP1_R.KO_JGYOBU, vbUnicode))
        GW_NAIGAI_KO = Trim(StrConv(ODR_TP1_R.KO_NAIGAI, vbUnicode))
        GW_HINGAI_KO = Trim(StrConv(ODR_TP1_R.KO_HIN_GAI, vbUnicode))
        
        If StrConv(ODR_TP1_R.KAN_KB, vbUnicode) = "0" Then
            W_STR = StrConv(ODR_TP1_R.HIN_GAI, vbUnicode)
        End If
        
        W_STR = StrConv(ODR_TP1_R.HIN_GAI, vbUnicode) & StrConv(ODR_TP1_R.ORDER_NO, vbUnicode)
        
        W_QTY = CDbl(StrConv(ODR_TP1_R.REQ_QTY, vbUnicode))     '�Ώې��F���v��
        
        '       08.12.12 ���L�ɕύX
        W_QTY = CDbl(StrConv(ODR_TP1_R.ALL_QTY, vbUnicode))     '�Ώې��F�W�J��
        
        
        
        W_Date = StrConv(ODR_TP1_R.CYUMON_DT, vbUnicode)
        
        'If W_Date >= W_NOW Then                         '�u�K�v���������v�́A������������B
        '2008/09/10         �����̔���͖��Ӗ��I
        
            If W_QTY = 0 Then
                W_Date = ""                             '���v�����O�@���@�����i�̎q���i
            Else
                If GW_HINGAI_KO = "B016" Then
                    W_QTY = W_QTY * 1
                End If
                
                
                                                    '2010/03/04 �}�C�i�X�͈������Ȃ��I �F�o�O�I�@(*_*;
                If W_QTY > 0 Then
                    If Zaiko_Hikiate(W_Date, W_QTY) Then
                        MsgBox "�݌Ɉ��������G���[�I", vbExclamation
                        Exit Function
                    End If
                End If
                
            End If
            
            
            '   08.12.12�ύX�F�����̎��Ɉ������ʂ̓��t��ݒ肷��B
            If StrConv(ODR_TP1_R.KAN_KB, vbUnicode) = "1" Then
                Call UniCode_Conv(ODR_TP1_R.OK_DT, W_Date)
            Else
                W_STR = StrConv(ODR_TP1_R.KAN_KB, vbUnicode)
            End If
            
            
            W_STR = CStr(W_QTY)
            
            Call UniCode_Conv(ODR_TP1_R.FUSOKU_QTY, W_STR)
            
            
            Call UniCode_Conv(ODR_TP1_R.UPDT_DT, Right(Format(Date, "yyyymmdd"), 6))
            Call UniCode_Conv(ODR_TP1_R.UPDT_TM, Left(Format(Time, "hhmmss"), 4))
            
            Do
                sts = BTRV(BtOpUpdate, ODR_TP1_POS, ODR_TP1_R, Len(ODR_TP1_R), K3_ODR_TEMP1, Len(K3_ODR_TEMP1), 3)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                        Sleep (500)
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "ODR_TEMP1")
                        Exit Function
                End Select
            Loop
            
        'End If
        
        com = BtOpGetNext
    Loop
    
    
    
    ZAN_CALC = False

End Function
Function SET_O_MAINA(HIN_GAI As String) As Integer
'
'                                 '       �e�����̃}�C�i�X�f�[�^���݌Ɍ��􂵂ŉ��Z����B
'           �\���W�J�@���@�݌ɐ����Z�I      io�敪����
'

Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim W_SeqKey    As String
Dim W_Ko_HinCD  As String

Dim Z_QTY       As Double
Dim W_QTY       As Double
Dim W_STR       As String
Dim W_Date      As String


Dim W_YOKU_YM   As String


    SET_O_MAINA = True
    
                                                                    
    'Z_QTY = CDbl(StrConv(ODR_ORDER_REC.ODR_QTY, vbUnicode)) * -1       '�e�̒�����
    Z_QTY = Abs(CDbl(StrConv(ODR_ORDER_REC.ODR_QTY, vbUnicode)))       '�e�̒�����
    '
    '�w�肳�ꂽ�e�i�Ԃ̍\���q���i�̑S�ĂɊւ��āA�W�J����ݒ肷��I
    '
    W_SeqKey = ""
    
    '   �ŏ��Ɂu�e���R�[�h�v���擾�B
    
    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, GW_SIMUKE)
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, GW_JIGYOBU)
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, GW_NAIGAI)
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, HIN_GAI)
    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, "0")
    Call UniCode_Conv(K0_P_COMPO.SEQNO, "")
    com = BtOpGetGreaterEqual
    sts = BTRV(com, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                'Beep
                'MsgBox "�w�肳�ꂽ�H��������܂���B"
        Case Else
            Call File_Error(sts, com, "P_COMPO")
            Exit Function
    End Select
    If sts <> BtNoErr Then
        MsgBox "�\�����@���o�^�I <" & HIN_GAI & ">", vbExclamation
        Exit Function
    End If
    
    '   ��������u�q���i���R�[�h�v��ǂ݂Ȃ���W�J�e���o�͂���B
    com = BtOpGetNext
    Do
        sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                'Beep
                'MsgBox "�w�肳�ꂽ�H��������܂���B"
            Case Else
                Call File_Error(sts, com, "P_COMPO")
                Exit Function
        End Select
        If sts <> BtNoErr Then Exit Do
        
        If Trim(StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode)) <> Trim(GW_SIMUKE) Then Exit Do
        If Trim(StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode)) <> Trim(GW_JIGYOBU) Then Exit Do
        If Trim(StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode)) <> Trim(GW_NAIGAI) Then Exit Do
        If Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(HIN_GAI) Then Exit Do
        'If Trim(StrConv(P_COMPO_O_REC.DATA_KBN, vbUnicode)) <> "0" Then Exit Do
        
        W_SeqKey = StrConv(P_COMPO_K_REC.SEQNO, vbUnicode)
        
        If CInt(W_SeqKey) <> 0 Then                 '�\�����i���R�[�h�H
            Call UniCode_Conv(K0_ODR_TEMP2.KO_JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))    '�q�@���ƕ�
            Call UniCode_Conv(K0_ODR_TEMP2.KO_NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))    '�q�@�����O
            
            
            '   2008/10/06  2008.04.04�̓W�J�����̏C���Ɠ���ɂ����I
            Call UniCode_Conv(K0_ODR_TEMP2.KO_JGYOBU, SHIZAI)    '�q�@���ƕ�
            Call UniCode_Conv(K0_ODR_TEMP2.KO_NAIGAI, NAIGAI_NAI)    '�q�@�����O
            
            Call UniCode_Conv(K0_ODR_TEMP2.KO_HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))  '�q�@�i��
            
            Call UniCode_Conv(K0_ODR_TEMP2.IO_KB, "b")             'io�敪
            
                                               '�g�p��
            Call UniCode_Conv(K0_ODR_TEMP2.USE_YM, StrConv(ODR_ORDER_REC.USE_YM, vbUnicode))
            
            Call UniCode_Conv(K0_ODR_TEMP2.ANS_NOUKI_DT, "")           '������
            Call UniCode_Conv(K0_ODR_TEMP2.ORDER_NO, "")            '������
            
            sts = BTRV(BtOpGetEqual, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, Len(K0_ODR_TEMP2), 0)
            Select Case sts
                Case BtNoErr
                    com = BtOpUpdate
                Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                    'MsgBox "�w�肳�ꂽ�H��������܂���B"
                    
                    Call ODR_TEMP2_CLR
                    '�q�@���ƕ�
                    Call UniCode_Conv(ODR_TP2_R.KO_JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
                    '�q�@�����O
                    Call UniCode_Conv(ODR_TP2_R.KO_NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
                    
                    
                    '   2008/10/06  2008.04.04�̓W�J�����̏C���Ɠ���ɂ����I
                    '�q�@���ƕ�
                    Call UniCode_Conv(ODR_TP2_R.KO_JGYOBU, SHIZAI)
                    '�q�@�����O
                    Call UniCode_Conv(ODR_TP2_R.KO_NAIGAI, NAIGAI_NAI)

                    '�q�i��
                    Call UniCode_Conv(ODR_TP2_R.KO_HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                    'io�敪
                    Call UniCode_Conv(ODR_TP2_R.IO_KB, "b")
                    '�g�p��
                    Call UniCode_Conv(ODR_TP2_R.USE_YM, StrConv(ODR_ORDER_REC.USE_YM, vbUnicode))
                    
                    com = BtOpInsert
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "ODR_TEMP2")
                    Exit Function
            End Select
                                    
            If Trim(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) = "" Or _
                Not IsNumeric(Trim(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode))) Then
                    '���̂悤�Ȉُ�f�[�^�̏ꍇ�́u�P�v�Ƃ݂Ȃ��B
                W_QTY = Z_QTY * 1
            Else
                W_QTY = Z_QTY * CDbl(Trim(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)))
            End If
            W_QTY = W_QTY + CDbl(Trim(StrConv(ODR_TP2_R.ZAI_QTY, vbUnicode)))
            
            
            W_STR = CStr(W_QTY)
            
            If Trim(W_STR) = "" Then W_STR = "0"
            Call UniCode_Conv(ODR_TP2_R.ZAI_QTY, W_STR)     '�݌ɐ�
            Call UniCode_Conv(ODR_TP2_R.MOTO_QTY, W_STR)    '���X�̍݌ɐ�
            
            Call UniCode_Conv(ODR_TP2_R.UPDT_DT, Right(Format(Date, "yyyymmdd"), 6))
            Call UniCode_Conv(ODR_TP2_R.UPDT_TM, Left(Format(Time, "hhmmss"), 4))
            Do
                sts = BTRV(com, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, Len(K0_ODR_TEMP2), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                        Sleep (500)
                    Case Else
                        Call File_Error(sts, com, "ODR_TEMP2")
                        Exit Function
                End Select
            Loop
            
        End If
        
        com = BtOpGetNext
    Loop
    
    
    SET_O_MAINA = False
    
End Function

Function Zaiko_Hikiate(OK_DT As String, W_QTY As Double) As Integer

        '       �݌ɁA���v�ʁ�������񂪕\�����ꂽ�̂ŁA���t���ɍ��������c����ݒ�B

Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim W_STR       As String

Dim W_Key       As String
Dim W_Date      As String

Dim W_K_io      As String
Dim W_K_DT      As String
Dim W_K_No      As String

Dim W_IN        As Double


If Trim(GW_HINGAI_KO) = "B016" Then
    Debug.Print
End If


    Zaiko_Hikiate = True
    
    W_Date = OK_DT
    
    
    OK_DT = ""
    W_K_io = ""
    W_K_DT = ""
    W_K_No = ""
    com = BtOpGetGreater
    Do
    
        Call UniCode_Conv(K0_ODR_TEMP2.KO_JGYOBU, GW_JIGYOBU_KO)    '�q�@���ƕ�
        Call UniCode_Conv(K0_ODR_TEMP2.KO_NAIGAI, GW_NAIGAI_KO)     '�q�@�����O
        Call UniCode_Conv(K0_ODR_TEMP2.KO_HIN_GAI, GW_HINGAI_KO)    '�q�i��
        Call UniCode_Conv(K0_ODR_TEMP2.IO_KB, W_K_io)               'io�敪
        Call UniCode_Conv(K0_ODR_TEMP2.ANS_NOUKI_DT, W_K_DT)        '������
        Call UniCode_Conv(K0_ODR_TEMP2.ORDER_NO, W_K_No)            '������
        
        'com = BtOpGetGreater
        sts = BTRV(com, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, Len(K0_ODR_TEMP2), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                'MsgBox "�w�肳�ꂽ�H��������܂���B"
            Case Else
                Call File_Error(sts, com, "ODR_TEMP2")
                Exit Function
        End Select
        If sts <> BtNoErr Then Exit Do
        
        
        If Trim(StrConv(ODR_TP2_R.KO_JGYOBU, vbUnicode)) <> Trim(GW_JIGYOBU_KO) Then Exit Do
        If Trim(StrConv(ODR_TP2_R.KO_NAIGAI, vbUnicode)) <> Trim(GW_NAIGAI_KO) Then Exit Do
        If Trim(StrConv(ODR_TP2_R.KO_HIN_GAI, vbUnicode)) <> Trim(GW_HINGAI_KO) Then Exit Do
        
        
        
        
'''2008.04.10        If StrConv(ODR_TP2_R.CYUMON_DT, vbUnicode) > W_DATE Then Exit Do
        
        
If Trim(GW_HINGAI_KO) = "B123" Then
    If StrConv(ODR_TP1_R.USE_YM, vbUnicode) = "200812" Then
    Debug.Print
    End If
End If
        
        
        W_K_io = StrConv(ODR_TP2_R.IO_KB, vbUnicode)
        W_K_DT = StrConv(ODR_TP2_R.ANS_NOUKI_DT, vbUnicode)
        W_K_No = StrConv(ODR_TP2_R.ORDER_NO, vbUnicode)
        
        W_IN = CDbl(StrConv(ODR_TP2_R.ZAI_QTY, vbUnicode))
        
        
        '           �񓚔[�����󔒁@���@�ΏۊO�I�H
        If StrConv(ODR_TP2_R.IO_KB, vbUnicode) = "g" Then
                
            W_IN = 0
        
        End If
        
        
        '2008.12.02     �d���c�̓��A�g�p���̈قȂ�d���c�͈����ΏۊO�Ƃ���I�H
        'If StrConv(ODR_TP2_R.IO_KB, vbUnicode) = "f" Then
        '    If StrConv(ODR_TP2_R.USE_YM, vbUnicode) <> "" Then
        '        If StrConv(ODR_TP2_R.USE_YM, vbUnicode) <> StrConv(ODR_TP1_R.USE_YM, vbUnicode) Then
        '            W_IN = 0
        '        End If
        '    End If
        'End If
        
        '2009.07.13
        '           ��L�́u�d���c�قȂ饥��v�̓~�X�I
        '               �u�قȂ�v�ł͂Ȃ��A�u�g�p�����d���c�̎g�p���v����Ȃ��ƁA�O���c���g�p����Ȃ��I�H
        If StrConv(ODR_TP2_R.IO_KB, vbUnicode) = "f" Then
            If StrConv(ODR_TP2_R.USE_YM, vbUnicode) <> "" Then
                If StrConv(ODR_TP2_R.USE_YM, vbUnicode) > StrConv(ODR_TP1_R.USE_YM, vbUnicode) Then
                    W_IN = 0
                End If
            End If
        End If
        
        
        
        If W_IN > 0 Then
            
            If W_QTY <= W_IN Then
                W_IN = W_IN - W_QTY
                W_QTY = 0
                OK_DT = Trim(StrConv(ODR_TP2_R.ANS_NOUKI_DT, vbUnicode))
                
                If Trim(OK_DT) = "" Then
                    OK_DT = Format(Date, "yyyymmdd")        '�݌Ƀf�[�^�ŉ\�I
                End If
                
            Else
                W_QTY = W_QTY - W_IN
                W_IN = 0
                
            End If
            
            
            W_STR = CStr(W_IN)
            
            Call UniCode_Conv(ODR_TP2_R.ZAI_QTY, W_STR)
            
            Call UniCode_Conv(ODR_TP2_R.UPDT_DT, Right(Format(Date, "yyyymmdd"), 6))
            Call UniCode_Conv(ODR_TP2_R.UPDT_TM, Left(Format(Time, "hhmmss"), 4))
    
            Do
                sts = BTRV(BtOpUpdate, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, Len(K0_ODR_TEMP2), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                        Sleep (500)
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "ODR_TEMP2")
                        Exit Do
                End Select
            Loop
        
        End If
        
        
        If W_QTY <= 0 Then              '08/11/14 �u���v���u���v�ɕύX�I�@(*_*;
            Exit Do
        End If
        
        
        com = BtOpGetNext '+ BtSNoWait
    Loop
    
    
    Zaiko_Hikiate = False

End Function

Function OK_DT_SRCH(OK_DT As String) As Integer
'
'           �g���\���@����
'

Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

    OK_DT_SRCH = True
    OK_DT = ""
        
If "TEST-1" = Trim(Key_HinGai) Then
    Debug.Print
End If
        
        '
    '�w�肳�ꂽ�e�i�Ԃ̍\���q���i�̑S�ĂɊւ��āA�݌ɁE���������m�F����I�H�@(��;)
    '
    Call UniCode_Conv(K1_ODR_TEMP1.SHIMUKE, Key_SIMUKE)    '�d������
    Call UniCode_Conv(K1_ODR_TEMP1.JGYOBU, Key_JIGYOBU)         '���ƕ�
    Call UniCode_Conv(K1_ODR_TEMP1.NAIGAI, Key_NAIGAI)          '�����O
    Call UniCode_Conv(K1_ODR_TEMP1.HIN_GAI, Key_HinGai)         '�e�i��
    Call UniCode_Conv(K1_ODR_TEMP1.ORDER_NO, Key_ORDER_NO)      '�e�i�ԁ@������
    Call UniCode_Conv(K1_ODR_TEMP1.INS_NO, Key_INS_NO)          '�o�^��
    Call UniCode_Conv(K1_ODR_TEMP1.BUN_NO, Key_BUN_NO)          '���[��
    Call UniCode_Conv(K1_ODR_TEMP1.OK_DT, "")                   '��葵����
        
    com = BtOpGetGreaterEqual
    Do
        sts = BTRV(com, ODR_TP1_POS, ODR_TP1_R, Len(ODR_TP1_R), K1_ODR_TEMP1, Len(K1_ODR_TEMP1), 1)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                
                
            Case Else
                Call File_Error(sts, com, "ODR_TEMP1")
                Exit Function
        End Select
        If sts <> BtNoErr Then Exit Do
        If Trim(StrConv(ODR_TP1_R.SHIMUKE, vbUnicode)) <> Trim(Key_SIMUKE) Then Exit Do
        If Trim(StrConv(ODR_TP1_R.JGYOBU, vbUnicode)) <> Trim(Key_JIGYOBU) Then Exit Do
        If Trim(StrConv(ODR_TP1_R.NAIGAI, vbUnicode)) <> Trim(Key_NAIGAI) Then Exit Do
        If Trim(StrConv(ODR_TP1_R.HIN_GAI, vbUnicode)) <> Trim(Key_HinGai) Then Exit Do
        If Trim(StrConv(ODR_TP1_R.ORDER_NO, vbUnicode)) <> Trim(Key_ORDER_NO) Then Exit Do
        If Trim(StrConv(ODR_TP1_R.INS_NO, vbUnicode)) <> Trim(Key_INS_NO) Then Exit Do
        If Trim(StrConv(ODR_TP1_R.BUN_NO, vbUnicode)) <> Trim(Key_BUN_NO) Then Exit Do
        
        
        'OK_DT = Trim(StrConv(ODR_TP1_R.OK_DT, vbUnicode))
        
        'If Trim(OK_DT) = "" Then Exit Do            '�݌ɕs���̃f�[�^�L��I�I
        
        '2008/12.16
        '               ���L�ɕύX
        If Trim(StrConv(ODR_TP1_R.OK_DT, vbUnicode)) = "" Then
            OK_DT = ""
            Exit Do
        End If
        
        If OK_DT < Trim(StrConv(ODR_TP1_R.OK_DT, vbUnicode)) Then
            OK_DT = Trim(StrConv(ODR_TP1_R.OK_DT, vbUnicode))
        End If
        
        com = BtOpGetNext
    Loop
    
    OK_DT_SRCH = False


End Function


Function OUT_KENTO() As Integer
'
'           �e�g�p���ʁA�q���i���ƂɁA�����݌ɐ��`�K�v���ȂǊe�퍀�ڂ�ݒ�
'
Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer
Dim W_STR       As String
Dim W_Dbl       As Double
Dim W_QTY       As Double
Dim W_ODR       As Double

Dim X_i         As Integer
Dim X_j         As Integer

Dim W_From      As String
Dim W_To        As String
Dim W_Key1      As String
Dim W_Key2      As String
Dim W_Key3      As String
Dim W_Key4      As String

Dim W_YYMM      As String


Dim LAST_ORDER_DT   As String       '2016.12.14
Dim LAST_ORDER_QTY  As String       '2016.12.14
Dim i               As Integer      '2016.12.14



    OUT_KENTO = True
    
        
    W_From = Left(GW_TOUGETU, 4) & "/" & Right(GW_TOUGETU, 2) & "/01" '��̔N���iyyyymm�j
    W_To = ""
    
    
            '���������e Close ���@��LOpen �� Close �� KILL �� ��LOpen
    
    If ODR_KENTO_Open(BtOpenExec) Then
        MsgBox "�����𒆒f���܂��B", vbExclamation
        Exit Function
    End If
    
    sts = BTRV(BtOpClose, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_KENTO")
        End If
    End If
    
    If ODR_KENTO_KILL Then
        Exit Function
    End If
    
    If ODR_KENTO_Open(BtOpenExec) Then
        MsgBox "�����𒆒f���܂��B", vbExclamation
        Exit Function
    End If
    
    Call ODR_KENTO_CLR
    W_STR = Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:nn:ss")
    Call UniCode_Conv(ODR_KNT_R.ITEM_NM, W_STR)
    
    sts = BTRV(BtOpInsert, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
    Select Case sts
        Case BtNoErr
                    
        Case Else
            Call File_Error(sts, BtOpInsert, "ODR_KENTO")
            Exit Function
    End Select
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                                '   �����݌ɂ̏o��
    com = BtOpGetFirst
    Do
        Do
            sts = BTRV(com, ODR_ZK_POS, ODR_ZK_R, Len(ODR_ZK_R), K0_ODR_ZK, Len(K0_ODR_ZK), 0)
            Select Case sts
                Case BtNoErr
                    com = BtOpUpdate
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF
                    
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                    Sleep (500)
                Case Else
                    Call File_Error(sts, com, "ODR_ZAIKO")
                    Exit Do
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
        
        Call ODR_KENTO_CLR
        
        If Trim(StrConv(ODR_ZK_R.KO_HIN_GAI, vbUnicode)) = "B015" Then
            sts = BtNoErr
        End If
            
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(ODR_ZK_R.KO_JGYOBU, vbUnicode))    '���ƕ�
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(ODR_ZK_R.KO_NAIGAI, vbUnicode))      '�����O
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(ODR_ZK_R.KO_HIN_GAI, vbUnicode))     '�q�i��
                
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                    '*********************************************************************
                            'TEST�I�ɕҏW�I (^_^;)
                    '�i��
                Call UniCode_Conv(ITEMREC.HIN_NAME, "���o�^")
                    '�������b�g
                W_STR = String(UBound(ITEMREC.G_SHIIRE_TBL(0).LOT), "0") & "1"
                Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).LOT, W_STR)
                    '�d����
                Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).CODE, "")
                    
                    '�d���P��
                W_STR = String(UBound(ITEMREC.G_SHIIRE_TBL(0).TANKA), "0") & "1"
                Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).TANKA, W_STR)
                
                sts = BtNoErr
                
            Case Else
                Call File_Error(sts, BtOpGetEqual, "ITEM")
                'Exit Function
        End Select
    
        If sts = BtNoErr Then
                
            Call UniCode_Conv(ODR_KNT_R.ITEM_NM, Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode)))
                
            
            
'>>>>>>>>>>>>>>>>>>>>>>>>>> �ŐV���������g�p����    2016.12.14
            
            If IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(0).LOT, vbUnicode)) Then
                Call UniCode_Conv(ODR_KNT_R.LOT_QTY, StrConv(ITEMREC.G_SHIIRE_TBL(0).LOT, vbUnicode))
            Else
                Call UniCode_Conv(ODR_KNT_R.LOT_QTY, "00000000.00")
            End If
                
                
                
If Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) = "C200" Then
    Debug.Print
End If
                
                
            LAST_ORDER_DT = ""
            LAST_ORDER_QTY = ""


            For i = 0 To 2
                If StrConv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_DT, vbUnicode) > LAST_ORDER_DT Then
                    LAST_ORDER_DT = StrConv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_DT, vbUnicode)
                    LAST_ORDER_QTY = StrConv(ITEMREC.G_SHIIRE_TBL(i).LOT, vbUnicode)
                End If
            Next i

            If IsNumeric(LAST_ORDER_QTY) Then
                Call UniCode_Conv(ODR_KNT_R.LOT_QTY, LAST_ORDER_QTY)
            Else
                Call UniCode_Conv(ODR_KNT_R.LOT_QTY, "00000000.00")
            End If

'>>>>>>>>>>>>>>>>>>>>>>>>>> �ŐV���������g�p����    2016.12.14
                
                
                
            Call UniCode_Conv(ODR_KNT_R.SECT, StrConv(ITEMREC.G_SHIIRE_TBL(0).CODE, vbUnicode))
                
            If IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode)) Then
                Call UniCode_Conv(ODR_KNT_R.TANKA, StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode))
            Else
                Call UniCode_Conv(ODR_KNT_R.TANKA, "00000000.00")
            End If
            
            '�ꊇ�����敪�̐ݒ�i09.05.22)
            Call UniCode_Conv(ODR_KNT_R.IKKATU_MK, StrConv(ITEMREC.AVE_SYUKA, vbUnicode))
                        
        End If
        
        
        
        Call UniCode_Conv(ODR_KNT_R.KO_JGYOBU, StrConv(ODR_ZK_R.KO_JGYOBU, vbUnicode))
        Call UniCode_Conv(ODR_KNT_R.KO_NAIGAI, StrConv(ODR_ZK_R.KO_NAIGAI, vbUnicode))
        Call UniCode_Conv(ODR_KNT_R.KO_HIN_GAI, StrConv(ODR_ZK_R.KO_HIN_GAI, vbUnicode))
        

        For X_i = 0 To UBound(ODR_ZK_R.ALL_ZAI)
            W_To = Left(Format(DateAdd("m", X_i, W_From), "yyyymmdd"), 6)
            Call UniCode_Conv(ODR_KNT_R.USE_YM, W_To)
            
            W_STR = CStr(CDbl(Trim(StrConv(ODR_ZK_R.ALL_ZAI(X_i).Z_QTY, vbUnicode))))
            Call UniCode_Conv(ODR_KNT_R.ZAI_QTY, W_STR)
        
            sts = BTRV(BtOpInsert, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
            Select Case sts
                Case BtNoErr
                    
                Case Else
                    Call File_Error(sts, BtOpInsert, "ODR_KENTO")
                    Exit Do
            End Select
               
        Next X_i
        
        com = BtOpGetNext
    Loop
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                                '   �W�J�ς݃f�[�^�iTP1�j���W�J���E���v���E�g�p���@�o��
    
    com = BtOpGetFirst
    Do
        Do
            sts = BTRV(com, ODR_TP1_POS, ODR_TP1_R, Len(ODR_TP1_R), K2_ODR_TEMP1, Len(K2_ODR_TEMP1), 2)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                    
                    
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                    yn = MsgBox("���Ŏg�p���ł��I<���ԏ��v�ʂe>" & Chr(13) & Chr(10) & _
                                "�@�Ď��s���܂����H", vbYesNo + vbExclamation, "�m�F����")
                    If yn = vbNo Then Exit Do
                Case Else
                    Call File_Error(sts, com, "ODR_TEMP1")
                    Exit Do
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
        
        Call UniCode_Conv(K0_ODR_KENTO.USE_YM, StrConv(ODR_TP1_R.USE_YM, vbUnicode))
        Call UniCode_Conv(K0_ODR_KENTO.KO_JGYOBU, StrConv(ODR_TP1_R.KO_JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ODR_KENTO.KO_NAIGAI, StrConv(ODR_TP1_R.KO_NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ODR_KENTO.KO_HIN_GAI, StrConv(ODR_TP1_R.KO_HIN_GAI, vbUnicode))
        
        
        If Trim(StrConv(ODR_TP1_R.KO_HIN_GAI, vbUnicode)) = "B015" Then
            sts = BtNoErr
        End If
        
        
        Do
            sts = BTRV(BtOpGetEqual, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
            Select Case sts
                Case BtNoErr
                    com = BtOpUpdate
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF
                    Call ODR_KENTO_CLR
                    
                    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(ODR_TP1_R.KO_JGYOBU, vbUnicode))    '���ƕ�
                    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(ODR_TP1_R.KO_NAIGAI, vbUnicode))      '�����O
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(ODR_TP1_R.KO_HIN_GAI, vbUnicode))     '�q�i��
                            
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        
                        Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                                '*********************************************************************
                                        'TEST�I�ɕҏW�I (^_^;)
                                '�i��
                            Call UniCode_Conv(ITEMREC.HIN_NAME, "���o�^")
                                '�������b�g
                            W_STR = String(UBound(ITEMREC.G_SHIIRE_TBL(0).LOT), "0") & "1"
                            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).LOT, W_STR)
                                '�d����
                            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).CODE, "")
                                
                                '�d���P��
                            W_STR = String(UBound(ITEMREC.G_SHIIRE_TBL(0).TANKA), "0") & "1"
                            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).TANKA, W_STR)
                            
                            sts = BtNoErr
                            
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "ITEM")
                            'Exit Function
                    End Select
                
                    If sts = BtNoErr Then
                            
                        Call UniCode_Conv(ODR_KNT_R.ITEM_NM, Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode)))
                            
                        If IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(0).LOT, vbUnicode)) Then
                            Call UniCode_Conv(ODR_KNT_R.LOT_QTY, StrConv(ITEMREC.G_SHIIRE_TBL(0).LOT, vbUnicode))
                        Else
                            Call UniCode_Conv(ODR_KNT_R.LOT_QTY, "00000000.00")
                        End If
                            
                        Call UniCode_Conv(ODR_KNT_R.SECT, StrConv(ITEMREC.G_SHIIRE_TBL(0).CODE, vbUnicode))
                            
                        If IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode)) Then
                            Call UniCode_Conv(ODR_KNT_R.TANKA, StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode))
                        Else
                            Call UniCode_Conv(ODR_KNT_R.TANKA, "00000000.00")
                        End If
                        
                    End If
                                
                    
                    Call UniCode_Conv(ODR_KNT_R.USE_YM, StrConv(ODR_TP1_R.USE_YM, vbUnicode))
                    Call UniCode_Conv(ODR_KNT_R.KO_JGYOBU, StrConv(ODR_TP1_R.KO_JGYOBU, vbUnicode))
                    Call UniCode_Conv(ODR_KNT_R.KO_NAIGAI, StrConv(ODR_TP1_R.KO_NAIGAI, vbUnicode))
                    Call UniCode_Conv(ODR_KNT_R.KO_HIN_GAI, StrConv(ODR_TP1_R.KO_HIN_GAI, vbUnicode))
                    com = BtOpInsert
                    sts = BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                    Sleep (500)
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "ODR_KENTO")
                    Exit Do
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
        
                                            '�W�J��
        W_QTY = CDbl(Trim(StrConv(ODR_TP1_R.ALL_QTY, vbUnicode)))
        W_Dbl = W_QTY + CDbl(Trim(StrConv(ODR_KNT_R.ALL_QTY, vbUnicode)))
        W_STR = CStr(W_Dbl)
        Call UniCode_Conv(ODR_KNT_R.ALL_QTY, W_STR)
        
                                            '�g�p��
        W_QTY = CDbl(Trim(StrConv(ODR_TP1_R.USE_QTY, vbUnicode)))
        
        
    
        
        W_Dbl = W_QTY + CDbl(Trim(StrConv(ODR_KNT_R.USE_QTY, vbUnicode)))
'debug2019



If Trim(StrConv(ODR_KNT_R.KO_HIN_GAI, vbUnicode)) = "D550" Then
    Debug.Print StrConv(ODR_KNT_R.USE_YM, vbUnicode) & " " & W_Dbl & " " & CDbl(Trim(StrConv(ODR_KNT_R.USE_QTY, vbUnicode)))
End If
        
        W_STR = CStr(W_Dbl)
        Call UniCode_Conv(ODR_KNT_R.USE_QTY, W_STR)
    
                                            '�K�v��
        W_QTY = CDbl(Trim(StrConv(ODR_TP1_R.NED_QTY, vbUnicode)))
        W_Dbl = W_QTY + CDbl(Trim(StrConv(ODR_KNT_R.NED_QTY, vbUnicode)))
        W_STR = CStr(W_Dbl)
        Call UniCode_Conv(ODR_KNT_R.NED_QTY, W_STR)
    
    
        '08.11.27   ���������O�Ή�
                                    '�W�J��
        If CDbl(Trim(StrConv(ODR_TP1_R.ALL_QTY, vbUnicode))) < 0 Then
                                                '�W�J�������Βl�i�{�j�ɂ���I
            W_QTY = Abs(CDbl(Trim(StrConv(ODR_TP1_R.ALL_QTY, vbUnicode))))
            W_Dbl = W_QTY + CDbl(Trim(StrConv(ODR_KNT_R.MINASHI1, vbUnicode)))
            W_STR = CStr(W_Dbl)
            Call UniCode_Conv(ODR_KNT_R.MINASHI1, W_STR)
            
            
            
        End If
    
    
    
    
        Do
            sts = BTRV(com, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
            Select Case sts
                Case BtNoErr
                    
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                    Sleep (500)
                Case Else
                    Call File_Error(sts, com, "ODR_KENTO")
                    Exit Do
            End Select
        Loop
    
        
        com = BtOpGetNext
    Loop
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            '   �W�J�ς݃f�[�^�iTP2�j��蔼���i�A�ݒ��A�d���c�Ȃǁ@�o��
    
    com = BtOpGetFirst
    Do
        Do
            sts = BTRV(com, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K1_ODR_TEMP2, Len(K1_ODR_TEMP2), 1)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                    
                    
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                    yn = MsgBox("���Ŏg�p���ł��I<TEMP2>" & Chr(13) & Chr(10) & _
                                "�@�Ď��s���܂����H", vbYesNo + vbExclamation, "�m�F����")
                    If yn = vbNo Then Exit Do
                Case Else
                    Call File_Error(sts, com, "ODR_TEMP2")
                    Exit Do
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
        
        
        If Trim(StrConv(ODR_TP2_R.KO_HIN_GAI, vbUnicode)) = "B015" Then
            sts = BtNoErr
        End If
        
        
        Call UniCode_Conv(K0_ODR_KENTO.USE_YM, StrConv(ODR_TP2_R.USE_YM, vbUnicode))
        Call UniCode_Conv(K0_ODR_KENTO.KO_JGYOBU, StrConv(ODR_TP2_R.KO_JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ODR_KENTO.KO_NAIGAI, StrConv(ODR_TP2_R.KO_NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ODR_KENTO.KO_HIN_GAI, StrConv(ODR_TP2_R.KO_HIN_GAI, vbUnicode))
    
        Do
            sts = BTRV(BtOpGetEqual, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
            Select Case sts
                Case BtNoErr
                    com = BtOpUpdate
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF
                    Call ODR_KENTO_CLR
                    
                    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(ODR_TP2_R.KO_JGYOBU, vbUnicode))    '���ƕ�
                    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(ODR_TP2_R.KO_NAIGAI, vbUnicode))      '�����O
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(ODR_TP2_R.KO_HIN_GAI, vbUnicode))     '�q�i��
                            
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        
                        Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                                '*********************************************************************
                                        'TEST�I�ɕҏW�I (^_^;)
                                '�i��
                            Call UniCode_Conv(ITEMREC.HIN_NAME, "���o�^")
                                '�������b�g
                            W_STR = String(UBound(ITEMREC.G_SHIIRE_TBL(0).LOT), "0") & "1"
                            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).LOT, W_STR)
                                '�d����
                            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).CODE, "")
                                
                                '�d���P��
                            W_STR = String(UBound(ITEMREC.G_SHIIRE_TBL(0).TANKA), "0") & "1"
                            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).TANKA, W_STR)
                            
                            sts = BtNoErr
                            
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "ITEM")
                            'Exit Function
                    End Select
                
                    If sts = BtNoErr Then
                            
                        Call UniCode_Conv(ODR_KNT_R.ITEM_NM, Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode)))
                            
                        If IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(0).LOT, vbUnicode)) Then
                            Call UniCode_Conv(ODR_KNT_R.LOT_QTY, StrConv(ITEMREC.G_SHIIRE_TBL(0).LOT, vbUnicode))
                        Else
                            Call UniCode_Conv(ODR_KNT_R.LOT_QTY, "00000000.00")
                        End If
                            
                        Call UniCode_Conv(ODR_KNT_R.SECT, StrConv(ITEMREC.G_SHIIRE_TBL(0).CODE, vbUnicode))
                            
                        If IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode)) Then
                            Call UniCode_Conv(ODR_KNT_R.TANKA, StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode))
                        Else
                            Call UniCode_Conv(ODR_KNT_R.TANKA, "00000000.00")
                        End If
                        
                    End If
                    
                    Call UniCode_Conv(ODR_KNT_R.USE_YM, StrConv(ODR_TP2_R.USE_YM, vbUnicode))
                    Call UniCode_Conv(ODR_KNT_R.KO_JGYOBU, StrConv(ODR_TP2_R.KO_JGYOBU, vbUnicode))
                    Call UniCode_Conv(ODR_KNT_R.KO_NAIGAI, StrConv(ODR_TP2_R.KO_NAIGAI, vbUnicode))
                    Call UniCode_Conv(ODR_KNT_R.KO_HIN_GAI, StrConv(ODR_TP2_R.KO_HIN_GAI, vbUnicode))
                    com = BtOpInsert
                    sts = BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                    Sleep (500)
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "ODR_KENTO")
                    Exit Do
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
        
        W_QTY = CDbl(Trim(StrConv(ODR_TP2_R.MOTO_QTY, vbUnicode)))
        
        If Trim(StrConv(ODR_TP2_R.KO_HIN_GAI, vbUnicode)) = "AD-KZ061X" Then
            W_QTY = W_QTY * 1
        End If
        
    'ZAI_QTY(0 To 10)            As Byte         '�����݌ɐ�     9(8)v9(2)
    'MAI_QTY(0 To 10)            As Byte         '�s����         9(8)v9(2)
    'ODR_QTY(0 To 10)            As Byte         '������         9(8)v9(2)
    'SHI_QTY(0 To 10)            As Byte         '�d���c��       9(8)v9(2)
    'HANSEIHIN_QTY(0 To 10)      As Byte         '�����i��       9(8)v9(2)
    'ZAITEI_QTY(0 To 10)         As Byte         '�ݒ��}��       9(8)v9(2)
    'KAITO(0 To 7)               As Byte         '�񓚔[��
    'ZAN_CNT(0 To 2)             As Byte         '�d���c�@����
    
        Select Case StrConv(ODR_TP2_R.IO_KB, vbUnicode)
            
            Case "a"            '�݌�
                
                'W_Dbl = W_QTY + CDbl(Trim(StrConv(ODR_KNT_R.ZAI_QTY, vbUnicode)))
                'W_Str = CStr(W_Dbl)
                'Call UniCode_Conv(ODR_KNT_R.ZAI_QTY, W_Str)
                
                W_YYMM = StrConv(ODR_TP2_R.USE_YM, vbUnicode)
                
                If Trim(StrConv(ODR_TP2_R.KO_HIN_GAI, vbUnicode)) = "K142" Then
                    W_STR = ""
                End If
                
                If W_YYMM = GW_TOUGETU Then
                    W_QTY = CDbl(Trim(StrConv(ODR_TP2_R.MOTO_QTY, vbUnicode)))
                    W_STR = CStr(W_QTY)
                    Call UniCode_Conv(ODR_KNT_R.ITEM_Z_QTY, W_STR)
                End If
            
            Case "b"            '�e�������O
                
                '2008/10/06
                'W_Dbl = W_QTY + CDbl(Trim(StrConv(ODR_KNT_R.ZAI_QTY, vbUnicode)))
                'W_Str = CStr(W_Dbl)
                'Call UniCode_Conv(ODR_KNT_R.ZAI_QTY, W_Str)
                
            
            Case "c"            '�d���ς�
            
                W_Dbl = W_QTY + CDbl(Trim(StrConv(ODR_KNT_R.MINASHI2, vbUnicode)))
                W_STR = CStr(W_Dbl)
                'Call UniCode_Conv(ODR_KNT_R.ZAI_QTY, W_Str)
                
                                '2008.12.02
                Call UniCode_Conv(ODR_KNT_R.MINASHI2, W_STR)
            
            Case "d"            '�ݒ�
                
                W_Dbl = W_QTY + CDbl(Trim(StrConv(ODR_KNT_R.ZAITEI_QTY, vbUnicode)))
                W_STR = CStr(W_Dbl)
                Call UniCode_Conv(ODR_KNT_R.ZAITEI_QTY, W_STR)
            
            Case "e"            '�����i
            
                W_Dbl = W_QTY + CDbl(Trim(StrConv(ODR_KNT_R.HANSEIHIN_QTY, vbUnicode)))
                W_STR = CStr(W_Dbl)
                Call UniCode_Conv(ODR_KNT_R.HANSEIHIN_QTY, W_STR)
            
            Case "f"            '�����c
            
                W_Dbl = W_QTY + CDbl(Trim(StrConv(ODR_KNT_R.SHI_QTY1, vbUnicode)))
                W_STR = CStr(W_Dbl)
                Call UniCode_Conv(ODR_KNT_R.SHI_QTY1, W_STR)
                
                
                'If Trim(StrConv(ODR_TP2_R.ANS_NOUKI_DT, vbUnicode)) <> "" Then
                If Trim(StrConv(ODR_KNT_R.KAITO, vbUnicode)) = "" Then
                    Call UniCode_Conv(ODR_KNT_R.KAITO, StrConv(ODR_TP2_R.ANS_NOUKI_DT, vbUnicode))
                End If
                    
                If Trim(StrConv(ODR_TP2_R.ANS_NOUKI_DT, vbUnicode)) < StrConv(ODR_KNT_R.KAITO, vbUnicode) Then
                    Call UniCode_Conv(ODR_KNT_R.KAITO, StrConv(ODR_TP2_R.ANS_NOUKI_DT, vbUnicode))
                End If
                    
                W_Dbl = CDbl(Trim(StrConv(ODR_KNT_R.ZAN_CNT, vbUnicode))) + 1
                W_STR = CStr(W_Dbl)
                Call UniCode_Conv(ODR_KNT_R.ZAN_CNT, W_STR)
                'End If
                
            Case "g"            '�����c�i�񓚔[�������j
            
                W_Dbl = W_QTY + CDbl(Trim(StrConv(ODR_KNT_R.SHI_QTY2, vbUnicode)))
                W_STR = CStr(W_Dbl)
                Call UniCode_Conv(ODR_KNT_R.SHI_QTY2, W_STR)
                
            Case Else
         
         
         
        End Select
            
        Do
            sts = BTRV(com, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
            Select Case sts
                Case BtNoErr
                    
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                    Sleep (500)
                Case Else
                    Call File_Error(sts, com, "ODR_KENTO")
                    Exit Do
            End Select
        Loop
        
        
        com = BtOpGetNext
    Loop
        
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    '�s�����̌v�Z
    
    com = BtOpGetFirst
    Do
        Do
            sts = BTRV(com, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                    
                    
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                    yn = MsgBox("���Ŏg�p���ł��I<ODR_KENTO>" & Chr(13) & Chr(10) & _
                                "�@�Ď��s���܂����H", vbYesNo + vbExclamation, "�m�F����")
                    If yn = vbNo Then Exit Do
                Case Else
                    Call File_Error(sts, com, "ODR_KENTO")
                    Exit Do
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
        W_Key1 = StrConv(ODR_KNT_R.USE_YM, vbUnicode)
        W_Key2 = StrConv(ODR_KNT_R.KO_JGYOBU, vbUnicode)
        W_Key3 = StrConv(ODR_KNT_R.KO_NAIGAI, vbUnicode)
        W_Key4 = StrConv(ODR_KNT_R.KO_HIN_GAI, vbUnicode)
        
        If Trim(W_Key4) = "B533" Then
            sts = BtNoErr
        End If
        
        
        W_Dbl = CDbl(Trim(StrConv(ODR_KNT_R.MAI_QTY, vbUnicode)))           '���X�̕s����
        
        W_Dbl = W_Dbl - CDbl(Trim(StrConv(ODR_KNT_R.NED_QTY, vbUnicode)))   '�|�K�v��
        
        W_Dbl = W_Dbl + CDbl(Trim(StrConv(ODR_KNT_R.ZAI_QTY, vbUnicode)))   '�{�݌ɐ�
        W_Dbl = W_Dbl + CDbl(Trim(StrConv(ODR_KNT_R.ODR_QTY, vbUnicode)))   '�{������
        W_Dbl = W_Dbl + CDbl(Trim(StrConv(ODR_KNT_R.SHI_QTY1, vbUnicode)))   '�{�d���c��
        
        
        '2008/12/10 �g�p�������Z�I
        W_Dbl = W_Dbl - CDbl(Trim(StrConv(ODR_KNT_R.USE_QTY, vbUnicode)))   '�|�g�p��
        
        
        '       �����i���́A�����݌ɂɉ��Z����Ă���I
        '       ������x���Z����ƁA�_�u�b�ĉ��Z���Ă��܂��I
        'W_Dbl = W_Dbl + CDbl(Trim(StrConv(ODR_KNT_R.HANSEIHIN_QTY, vbUnicode)))   '�{�����i��
        
        
        '       �ݒ��}���́A�����݌ɂɉ��Z����Ă���I
        '       ������x���Z����ƁA�_�u�b�ĉ��Z���Ă��܂��I
        'W_Dbl = W_Dbl + CDbl(Trim(StrConv(ODR_KNT_R.ZAITEI_QTY, vbUnicode)))   '�{�ݒ��}��
        
        
        W_Dbl = W_Dbl + CDbl(Trim(StrConv(ODR_KNT_R.SHI_QTY2, vbUnicode)))   '�{�d���c���i�񓚔[�������j
        
        W_STR = CStr(W_Dbl)
        Call UniCode_Conv(ODR_KNT_R.MAI_QTY, W_STR)         '�s����
        
        
        
        '�������̌v�Z
        If IsNumeric(Trim(StrConv(ODR_KNT_R.LOT_QTY, vbUnicode))) Then
            W_QTY = CDbl(Trim(StrConv(ODR_KNT_R.LOT_QTY, vbUnicode)))
            If W_QTY = 0 Then W_QTY = 1
        Else
            W_QTY = 1
        End If
        
        If W_Dbl < 0 Then
            W_Dbl = W_Dbl * -1
            W_ODR = 0
            Do
                If W_ODR >= W_Dbl Then Exit Do
                W_ODR = W_ODR + W_QTY
            Loop
            W_STR = CStr(W_ODR)
            Call UniCode_Conv(ODR_KNT_R.ODR_QTY, W_STR)
        End If
        
        
        W_Dbl = CDbl(StrConv(ODR_KNT_R.MAI_QTY, vbUnicode))
        
        Do
            sts = BTRV(BtOpUpdate, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
            Select Case sts
                Case BtNoErr
                    
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                    Sleep (500)
                Case Else
                    Call File_Error(sts, BtOpUpdate, "ODR_KENTO")
                    Exit Do
            End Select
        Loop
        
    com = 0
    If com <> 0 Then
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �����ȍ~�̕s�������v�Z
        Call UniCode_Conv(K1_ODR_KENTO.KO_JGYOBU, W_Key2)
        Call UniCode_Conv(K1_ODR_KENTO.KO_NAIGAI, W_Key3)
        Call UniCode_Conv(K1_ODR_KENTO.KO_HIN_GAI, W_Key4)
        Call UniCode_Conv(K1_ODR_KENTO.USE_YM, W_Key1)
        com = BtOpGetGreater
        Do
            Do
                sts = BTRV(com, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K1_ODR_KENTO, Len(K1_ODR_KENTO), 1)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                        
                        
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                        yn = MsgBox("���Ŏg�p���ł��I<ODR_KENTO>" & Chr(13) & Chr(10) & _
                                    "�@�Ď��s���܂����H", vbYesNo + vbExclamation, "�m�F����")
                        If yn = vbNo Then Exit Do
                    Case Else
                        Call File_Error(sts, com, "ODR_KENTO")
                        Exit Do
                End Select
            Loop
            If sts <> BtNoErr Then Exit Do
            If StrConv(ODR_KNT_R.KO_JGYOBU, vbUnicode) <> W_Key2 Then Exit Do
            If StrConv(ODR_KNT_R.KO_NAIGAI, vbUnicode) <> W_Key3 Then Exit Do
            If StrConv(ODR_KNT_R.KO_HIN_GAI, vbUnicode) <> W_Key4 Then Exit Do
            
            If StrConv(ODR_KNT_R.USE_YM, vbUnicode) > W_Key1 Then
                W_QTY = CDbl(Trim(StrConv(ODR_KNT_R.MAI_QTY, vbUnicode))) + W_Dbl
                W_STR = CStr(W_QTY)
                Call UniCode_Conv(ODR_KNT_R.MAI_QTY, W_STR)
                
                
                
                
                Do
                    sts = BTRV(BtOpUpdate, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
                    Select Case sts
                        Case BtNoErr
                            
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                            Sleep (500)
                        Case Else
                            Call File_Error(sts, BtOpUpdate, "ODR_KENTO")
                            Exit Do
                    End Select
                Loop
            End If
            
            com = BtOpGetNext
        Loop
    End If
        
        Call UniCode_Conv(K0_ODR_KENTO.USE_YM, W_Key1)
        Call UniCode_Conv(K0_ODR_KENTO.KO_JGYOBU, W_Key2)
        Call UniCode_Conv(K0_ODR_KENTO.KO_NAIGAI, W_Key3)
        Call UniCode_Conv(K0_ODR_KENTO.KO_HIN_GAI, W_Key4)
        
        com = BtOpGetGreater
    Loop

    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    '�ꊇ�����̏ꍇ�A���������O�Ƃ���B     '09.05.22
    com = BtOpGetFirst
    Do
        Do
            sts = BTRV(com, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                    
                    
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                    yn = MsgBox("���Ŏg�p���ł��I<ODR_KENTO>" & Chr(13) & Chr(10) & _
                                "�@�Ď��s���܂����H", vbYesNo + vbExclamation, "�m�F����")
                    If yn = vbNo Then Exit Do
                Case Else
                    Call File_Error(sts, com, "ODR_KENTO")
                    Exit Do
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
        
        
        '�ꊇ       2009.05.22
        'If StrConv(ODR_KNT_R.IKKATU_MK, vbUnicode) = "1" Then
        '    Call UniCode_Conv(ODR_KNT_R.ODR_QTY, "0")
        'End If
        
        Do
            sts = BTRV(BtOpUpdate, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
            Select Case sts
                Case BtNoErr
                    
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                    Sleep (500)
                Case Else
                    Call File_Error(sts, BtOpUpdate, "ODR_KENTO")
                    Exit Do
            End Select
        Loop
        

        com = BtOpGetNext
    Loop

    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

    

    OUT_KENTO = False
    
    sts = BTRV(BtOpClose, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_KENTO")
        End If
    End If
    
End Function
Function GESSYO_SET() As Integer
'
'           �q���i���ƂɁA�g�p���P�ʂ̌����݌ɐݒ�
'
Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer
Dim W_STR       As String
Dim W_Dbl       As Double
Dim W_QTY       As Double

Dim W_Key1      As String
Dim W_Key2      As String

Dim X_i         As Integer
Dim X_j         As Integer

Dim W_From      As String
Dim W_To        As String


    GESSYO_SET = True
                                
    W_Key1 = ""
    W_Key2 = ""
    
    Call ODR_ZAIKO_CLR
    
    W_From = Left(GW_TOUGETU, 4) & "/" & Right(GW_TOUGETU, 2) & "/01" '��̔N���iyyyymm�j
    W_To = ""
    
                            '�����݌ɂe Close ���@��LOpen �� Close �� KILL �� ��LOpen
                   
    sts = BTRV(BtOpClose, ODR_ZK_POS, ODR_ZK_R, Len(ODR_ZK_R), K0_ODR_ZK, Len(K0_ODR_ZK), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_ZAIKO")
        End If
    End If
    
    
    If ODR_ZAIKO_Open(BtOpenExec) Then
        MsgBox "�����𒆒f���܂��B", vbExclamation
        Exit Function
    End If
    sts = BTRV(BtOpClose, ODR_ZK_POS, ODR_ZK_R, Len(ODR_ZK_R), K0_ODR_ZK, Len(K0_ODR_ZK), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_ZAIKO")
        End If
    End If
    
    If ODR_ZAIKO_KILL Then
        Exit Function
    End If
    
    If ODR_ZAIKO_Open(BtOpenExec) Then
        MsgBox "�����𒆒f���܂��B", vbExclamation
        Exit Function
    End If
    
    '2008/10/17
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Call ODR_ZAIKO_CLR
    Call UniCode_Conv(ODR_ZK_R.FILLER, W_From)
    
    sts = BTRV(BtOpInsert, ODR_ZK_POS, ODR_ZK_R, Len(ODR_ZK_R), K0_ODR_ZK, Len(K0_ODR_ZK), 0)
    If sts <> BtNoErr Then
        MsgBox "�����݌Ɂ@����ǉ����s�I", vbExclamation
        sts = BTRV(BtOpClose, ODR_ZK_POS, ODR_ZK_R, Len(ODR_ZK_R), K0_ODR_ZK, Len(K0_ODR_ZK), 0)
        If sts Then
            If sts <> BtErrNoOpen Then
                Call File_Error(sts, BtOpClose, "ODR_ZAIKO")
            End If
        End If
        If ODR_ZAIKO_Open(BtOpenNomal) Then
            MsgBox "�����𒆒f���܂��B", vbExclamation
        End If

        Exit Function
    End If
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    
    
    com = BtOpGetFirst
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            '   �g�p�����󔒂̏ꍇ�A��N����ݒ肷��B
    Do
        Do
            sts = BTRV(com, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K1_ODR_TEMP2, Len(K1_ODR_TEMP2), 1)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                    yn = MsgBox("���Ŏg�p���ł��I<���ԏ��v�ʂe>" & Chr(13) & Chr(10) & _
                                "�@�Ď��s���܂����H", vbYesNo + vbExclamation, "�m�F����")
                    If yn = vbNo Then Exit Do
                Case Else
                    Call File_Error(sts, com, "ODR_TEMP2")
                    Exit Do
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
        
        If Trim(StrConv(ODR_TP2_R.USE_YM, vbUnicode)) = "" Then
            Call UniCode_Conv(ODR_TP2_R.USE_YM, GW_TOUGETU)
            Do
                sts = BTRV(BtOpUpdate, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K1_ODR_TEMP2, Len(K1_ODR_TEMP2), 1)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                        Sleep (500)
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "ODR_TEMP2")
                        Exit Do
                End Select
            Loop
            If sts <> BtNoErr Then Exit Do
        End If
        
        com = BtOpGetNext
    Loop
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
          
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                    '   �擪����ǂ݁A�g�p���P�ʂ̌����݌ɂ�ݒ肷��B
                    '   �g�p�����������f�[�^�́A�݌ɐ������Z����B
    com = BtOpGetFirst
    Do
        Do
            sts = BTRV(com, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K1_ODR_TEMP2, Len(K1_ODR_TEMP2), 1)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                    yn = MsgBox("���Ŏg�p���ł��I<���ԏ��v�ʂe>" & Chr(13) & Chr(10) & _
                                "�@�Ď��s���܂����H", vbYesNo + vbExclamation, "�m�F����")
                    If yn = vbNo Then Exit Do
                Case Else
                    Call File_Error(sts, com, "ODR_TEMP2")
                    Exit Do
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
        
        If Trim(StrConv(ODR_TP2_R.KO_HIN_GAI, vbUnicode)) = "B015" Then
            sts = BtNoErr
        End If
        If Trim(StrConv(ODR_TP2_R.IO_KB, vbUnicode)) = "b" Then
            sts = BtNoErr
        End If
        
        
        '2008/11/14
        Call UniCode_Conv(K0_ODR_ZK.KO_JGYOBU, StrConv(ODR_TP2_R.KO_JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ODR_ZK.KO_NAIGAI, StrConv(ODR_TP2_R.KO_NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ODR_ZK.KO_HIN_GAI, StrConv(ODR_TP2_R.KO_HIN_GAI, vbUnicode))
        Do
            sts = BTRV(BtOpGetEqual, ODR_ZK_POS, ODR_ZK_R, Len(ODR_ZK_R), K0_ODR_ZK, Len(K0_ODR_ZK), 0)
            Select Case sts
                Case BtNoErr
                    com = BtOpUpdate
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                    Call ODR_ZAIKO_CLR
                    Call UniCode_Conv(ODR_ZK_R.KO_JGYOBU, StrConv(ODR_TP2_R.KO_JGYOBU, vbUnicode))
                    Call UniCode_Conv(ODR_ZK_R.KO_NAIGAI, StrConv(ODR_TP2_R.KO_NAIGAI, vbUnicode))
                    Call UniCode_Conv(ODR_ZK_R.KO_HIN_GAI, StrConv(ODR_TP2_R.KO_HIN_GAI, vbUnicode))
                    com = BtOpInsert
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                    yn = MsgBox("���Ŏg�p���ł��I<ODR_ZAIKO>" & Chr(13) & Chr(10) & _
                                "�@�Ď��s���܂����H", vbYesNo + vbExclamation, "�m�F����")
                    If yn = vbNo Then Exit Function
                Case Else
                    Call File_Error(sts, com, "ODR_ZAIKO")
                    Exit Function
            End Select
        Loop
        
        W_STR = StrConv(ODR_TP2_R.USE_YM, vbUnicode)
        W_To = Left(W_STR, 4) & "/" & Right(W_STR, 2) & "/01"
        X_i = DateDiff("m", W_From, W_To)
             
        '2008/12/01 ���L�����ǉ��I
        If X_i >= UBound(ODR_ZK_R.ALL_ZAI) Then
            'MsgBox "[" & W_From & "] �` [" & W_To & "]  �g�p���@���Ԑݒ�ُ�I�H", vbExclamation
        
        Else
                    
                    '   �d���c�́A�����̌����݌ɂɔ��f�I�I                  '2008/09/27
            If StrConv(ODR_TP2_R.IO_KB, vbUnicode) >= "f" Then
                X_i = X_i + 1
            End If
                    
                    
            W_Dbl = CDbl(Trim(StrConv(ODR_TP2_R.MOTO_QTY, vbUnicode)))
            W_Dbl = W_Dbl + CDbl(Trim(StrConv(ODR_ZK_R.ALL_ZAI(X_i).Z_QTY, vbUnicode)))
                    
            W_STR = CStr(W_Dbl)
            Call Numeric_Check(EDIT_ONLY, UBound(ODR_ZK_R.ALL_ZAI(0).Z_QTY) + 1, 2, NEGA_ENA, ZSUP_DIS, _
                                COMA_DIS, CStr(W_Dbl), W_STR)
                    
            For X_j = X_i To UBound(ODR_ZK_R.ALL_ZAI)
                W_Dbl = CDbl(Trim(StrConv(ODR_TP2_R.MOTO_QTY, vbUnicode)))
                W_QTY = W_Dbl + CDbl(Trim(StrConv(ODR_ZK_R.ALL_ZAI(X_j).Z_QTY, vbUnicode)))
                W_STR = CStr(W_QTY)
                Call UniCode_Conv(ODR_ZK_R.ALL_ZAI(X_j).Z_QTY, W_STR)            '�݌ɐ�     9(5)v9(2)
            Next X_j
        
        End If
        
        Do
            sts = BTRV(com, ODR_ZK_POS, ODR_ZK_R, Len(ODR_ZK_R), K0_ODR_ZK, Len(K0_ODR_ZK), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                    Sleep (500)
                Case Else
                    Call File_Error(sts, com, "ODR_ZAIKO")
                    Exit Do
            End Select
        Loop
        
        '>>>>>>>>>>>>>  2008/11/14 ���L���R�����g�ɁI
        '
        '
        '                                                    '�񓚔[���̖����d���͏��O!08/09/19
        '                                                    '�܂ށI�I�I     2008/09/27
        ''If StrConv(ODR_TP2_R.IO_KB, vbUnicode) <> "g" Then
        '
        '    W_Key1 = StrConv(ODR_TP2_R.KO_JGYOBU, vbUnicode)
        '    W_Key1 = W_Key1 & StrConv(ODR_TP2_R.KO_NAIGAI, vbUnicode)
        '    W_Key1 = W_Key1 & StrConv(ODR_TP2_R.KO_HIN_GAI, vbUnicode)
        '    W_Key1 = W_Key1 & StrConv(ODR_TP2_R.USE_YM, vbUnicode)
        '
        '    If W_Key1 = W_Key2 Then
        '
        '        W_Str = StrConv(ODR_TP2_R.USE_YM, vbUnicode)
        '        W_To = Left(W_Str, 4) & "/" & Right(W_Str, 2) & "/01"
        '        X_i = DateDiff("m", W_From, W_To)
        '
        '
        '        '   �d���c�́A�����̌����݌ɂɔ��f�I�I                  '2008/09/27
        '        If StrConv(ODR_TP2_R.IO_KB, vbUnicode) >= "f" Then
        '            X_i = X_i + 1
        '        End If
        '
        '
        '
        '        W_Dbl = CDbl(Trim(StrConv(ODR_TP2_R.MOTO_QTY, vbUnicode)))
        '        W_Dbl = W_Dbl + CDbl(Trim(StrConv(ODR_ZK_R.ALL_ZAI(X_i).Z_QTY, vbUnicode)))
        '
        '        W_Str = CStr(W_Dbl)
        '        Call Numeric_Check(EDIT_ONLY, UBound(ODR_ZK_R.ALL_ZAI(0).Z_QTY) + 1, 2, NEGA_ENA, ZSUP_DIS, _
        '                    COMA_DIS, CStr(W_Dbl), W_Str)
        '
        '        For X_j = X_i To UBound(ODR_ZK_R.ALL_ZAI)
        '            W_Dbl = CDbl(Trim(StrConv(ODR_TP2_R.MOTO_QTY, vbUnicode)))
        '            W_QTY = W_Dbl + CDbl(Trim(StrConv(ODR_ZK_R.ALL_ZAI(X_j).Z_QTY, vbUnicode)))
        '            W_Str = CStr(W_QTY)
        '            Call UniCode_Conv(ODR_ZK_R.ALL_ZAI(X_j).Z_QTY, W_Str)            '�݌ɐ�     9(5)v9(2)
        '        Next X_j
        '
        '    Else
        '        If W_Key2 <> "" Then
        '            Do
        '                sts = BTRV(BtOpInsert, ODR_ZK_POS, ODR_ZK_R, Len(ODR_ZK_R), K0_ODR_ZK, Len(K0_ODR_ZK), 0)
        '                Select Case sts
        '                    Case BtNoErr
        '                        Exit Do
        '                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
        '                        Sleep (500)
        '                    Case Else
        '                        Call File_Error(sts, BtOpInsert, "ODR_ZAIKO")
        '                        Exit Do
        '                End Select
        '            Loop
        '            If sts <> BtNoErr Then
        '                MsgBox "�����݌Ɂ@�ǉ����s�I", vbExclamation
        '                'Exit Do
        '            End If
        '
        '        End If
        '
        '        Call ODR_ZAIKO_CLR
        '        '�q�@���ƕ�
        '        Call UniCode_Conv(ODR_ZK_R.KO_JGYOBU, StrConv(ODR_TP2_R.KO_JGYOBU, vbUnicode))
        '        '�q�@�����O
        '        Call UniCode_Conv(ODR_ZK_R.KO_NAIGAI, StrConv(ODR_TP2_R.KO_NAIGAI, vbUnicode))
        '        '�q�i��
        '        Call UniCode_Conv(ODR_ZK_R.KO_HIN_GAI, StrConv(ODR_TP2_R.KO_HIN_GAI, vbUnicode))
        '
        '        W_Dbl = CDbl(Trim(StrConv(ODR_TP2_R.MOTO_QTY, vbUnicode)))
        '
        '
        '        W_Str = StrConv(ODR_TP2_R.USE_YM, vbUnicode)
        '        W_To = Left(W_Str, 4) & "/" & Right(W_Str, 2) & "/01"
        '
        '        X_i = DateDiff("m", W_From, W_To)
        '
        '        W_Str = CStr(W_Dbl)
        '        Call Numeric_Check(EDIT_ONLY, UBound(ODR_ZK_R.ALL_ZAI(0).Z_QTY) + 1, 2, NEGA_ENA, ZSUP_DIS, _
        '                    COMA_DIS, CStr(W_Dbl), W_Str)
        '        W_Str = CStr(W_Dbl)
        '        For X_j = X_i To UBound(ODR_ZK_R.ALL_ZAI)
        '            Call UniCode_Conv(ODR_ZK_R.ALL_ZAI(X_j).Z_QTY, W_Str)            '�݌ɐ�     9(5)v9(2)
        '        Next X_j
        '
        '    End If
        '
        ''End If
        
        
        W_Key2 = W_Key1
        com = BtOpGetNext
        
    Loop
    
    If W_Key2 <> "" Then


        Do
            sts = BTRV(BtOpInsert, ODR_ZK_POS, ODR_ZK_R, Len(ODR_ZK_R), K0_ODR_ZK, Len(K0_ODR_ZK), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                    Sleep (500)
                Case Else
                    Call File_Error(sts, BtOpInsert, "ODR_ZAIKO")
                    Exit Do
            End Select
        Loop
        If sts <> BtNoErr Then
            MsgBox "�����݌Ɂ@�ǉ����s�I", vbExclamation
            'Exit Do
        End If
            
    End If
    
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                    '   �擪����ǂ݁A�g�p���P�ʂ̕K�v�����v�Z���A
                    '   �g�p���̗����݌ɂ��猸�Z���Č����݌ɂ��v�Z����B
    com = BtOpGetFirst
    
    Do
        Do
            sts = BTRV(com, ODR_TP1_POS, ODR_TP1_R, Len(ODR_TP1_R), K2_ODR_TEMP1, Len(K2_ODR_TEMP1), 2)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                    
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                    yn = MsgBox("���Ŏg�p���ł��I<���ԏ��v�ʂe>" & Chr(13) & Chr(10) & _
                                "�@�Ď��s���܂����H", vbYesNo + vbExclamation, "�m�F����")
                    If yn = vbNo Then Exit Do
                Case Else
                    Call File_Error(sts, com, "ODR_TEMP1")
                    Exit Do
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
        
        Call UniCode_Conv(K0_ODR_ZK.KO_JGYOBU, StrConv(ODR_TP1_R.KO_JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ODR_ZK.KO_NAIGAI, StrConv(ODR_TP1_R.KO_NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ODR_ZK.KO_HIN_GAI, StrConv(ODR_TP1_R.KO_HIN_GAI, vbUnicode))
        W_STR = StrConv(ODR_TP1_R.KO_HIN_GAI, vbUnicode)
        Do
            sts = BTRV(BtOpGetEqual, ODR_ZK_POS, ODR_ZK_R, Len(ODR_ZK_R), K0_ODR_ZK, Len(K0_ODR_ZK), 0)
            Select Case sts
                Case BtNoErr
                    com = BtOpUpdate
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF
                    Call ODR_ZAIKO_CLR
                    Call UniCode_Conv(ODR_ZK_R.KO_JGYOBU, StrConv(ODR_TP1_R.KO_JGYOBU, vbUnicode))
                    Call UniCode_Conv(ODR_ZK_R.KO_NAIGAI, StrConv(ODR_TP1_R.KO_NAIGAI, vbUnicode))
                    Call UniCode_Conv(ODR_ZK_R.KO_HIN_GAI, StrConv(ODR_TP1_R.KO_HIN_GAI, vbUnicode))
                    com = BtOpInsert
                    sts = BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                    Sleep (500)
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "ODR_ZAIKO")
                    Exit Do
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
                     
                     
        W_Dbl = CDbl(Trim(StrConv(ODR_TP1_R.NED_QTY, vbUnicode)))   '���v��
        If W_Dbl = 0 Then
            W_STR = StrConv(ODR_TP1_R.HIN_GAI, vbUnicode)
            W_Dbl = 0
        End If
        
        W_STR = StrConv(ODR_TP1_R.USE_YM, vbUnicode)
        W_To = Left(W_STR, 4) & "/" & Right(W_STR, 2) & "/01"
        X_i = DateDiff("m", W_From, W_To) + 1                   '�����̌����݌ɂ��猸�Z
          
             
        '2008/12/01 ���L�����ǉ��I
        If X_i >= UBound(ODR_ZK_R.ALL_ZAI) Then
            'MsgBox "[" & W_From & "] �` [" & W_To & "]  �g�p���@���Ԑݒ�ُ�I�H", vbExclamation
        
        Else
                        
              
            For X_j = X_i To UBound(ODR_ZK_R.ALL_ZAI)
                W_Dbl = CDbl(Trim(StrConv(ODR_TP1_R.ALL_QTY, vbUnicode)))   '�W�J��
                W_QTY = CDbl(StrConv(ODR_ZK_R.ALL_ZAI(X_j).Y_QTY, vbUnicode))
                W_QTY = W_QTY + W_Dbl
                
                W_STR = CStr(W_QTY)
                Call UniCode_Conv(ODR_ZK_R.ALL_ZAI(X_j).Y_QTY, W_STR)       '���W�J��
            
                
                W_Dbl = CDbl(Trim(StrConv(ODR_TP1_R.REQ_QTY, vbUnicode)))   '�K�v��
                
                W_Dbl = W_Dbl + CDbl(Trim(StrConv(ODR_TP1_R.USE_QTY, vbUnicode)))   '�g�p�� 2008/12/10
                
                W_QTY = CDbl(StrConv(ODR_ZK_R.ALL_ZAI(X_j).Z_QTY, vbUnicode))
                W_QTY = W_QTY - W_Dbl                                       '�݌ɐ��@�\�@�K�v��
                
                W_STR = CStr(W_QTY)
                Call UniCode_Conv(ODR_ZK_R.ALL_ZAI(X_j).Z_QTY, W_STR)       '�݌ɐ�
            
            Next X_j
                           
        End If
        
        
        sts = BTRV(com, ODR_ZK_POS, ODR_ZK_R, Len(ODR_ZK_R), K0_ODR_ZK, Len(K0_ODR_ZK), 0)
        Select Case sts
            Case BtNoErr
                
            Case Else
                Call File_Error(sts, com, "ODR_ZAIKO")
                Exit Do
        End Select
        
        com = BtOpGetNext
    Loop
    
    'End If
    
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                            '�����݌ɂe Close �� ���pOpen
    sts = BTRV(BtOpClose, ODR_ZK_POS, ODR_ZK_R, Len(ODR_ZK_R), K0_ODR_ZK, Len(K0_ODR_ZK), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_ZAIKO")
        End If
    End If
    If ODR_ZAIKO_Open(BtOpenNomal) Then
        MsgBox "�����𒆒f���܂��B", vbExclamation
        Exit Function
    End If
    
    GESSYO_SET = False
    
End Function

