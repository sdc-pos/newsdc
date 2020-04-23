Attribute VB_Name = "P_SAGYO_LOG_OUTPUT"
Option Explicit
Public Function P_SAGYO_LOG_OUTPUT_PROC(TANTO_CODE As String, _
                                    WEL_ID As String, _
                                    JGYOBU As String, _
                                    NAIGAI As String, _
                                    MENU_NO As String, _
                                    YOIN As String, _
                                    Optional HIN_GAI As String = "                    ", _
                                    Optional SUMI_QTY As Long = 0, _
                                    Optional MI_QTY As Long = 0, _
                                    Optional FROM_LOCATION As String = "        ", _
                                    Optional TO_LOCATION As String = "        ", _
                                    Optional ID_NO As String = "        ", _
                                    Optional MTS As String = "        ", _
                                    Optional SS As String = "        ", _
                                    Optional RETRY As Integer = 10, _
                                    Optional SHIJI_No As String = "        ", _
                                    Optional HIN_CHECK_LABEL_CNT As String = "   ", _
                                    Optional HIN_CHECK_GENPIN_CNT As String = "   ", _
                                    Optional wkMTS As String = "        ", _
                                    Optional JAN_CODE As String = "                    ", _
                                    Optional MEMO As String = "                                        ", _
                                    Optional HIN_CHECK_GAISOU_CNT As String = "   ", _
                                    Optional HINBAN_DAMMY As String = "                    ") As Integer
'****************************************************
'*      ��ƃ��O�o�͏����X�V
'*
'*  ��ƃ��O�̏o�͂��s���B
'*  (�����̐ݒ�~�X�͂�����ł̓`�F�b�N���Ȃ�)
'*  �����F  �S����(�ȗ��s��)
'*          ID(�ȗ��s��)
'*          ���ƕ��i�ȗ��s�j
'*          �����O�i�ȗ��s�j
'*          �ƭ��ԍ��i�ȗ��s�j
'*          �v��(�ȗ��s��)
'*          �O���i�ԁi�ȗ��� TOP�ƭ����j
'*          ���i���ςݎ��ѐ��i=0���Ƃ���A�����̂ݏo�́j
'*          �����i���ѐ��i=0���Ƃ���A�����̂ݏo�́j
'*          FROM�I�iXXXXXXXX(�q�ɇ�+��+�A+�i)�ȗ���
'*          TO�I�iXXXXXXXX(�q�ɇ�+��+�A+�i)�ȗ��j
'*          �`�[ID(�ȗ���)
'*          MTS(�ȗ���)
'*          SS(�ȗ���)
'*          ���g���C(�ȗ��� �P����:1=��ʃ��b�Z�[�W�L 0:���C�Q����:���g���C��(0�`9 0:����))
'*          �w�}�[��            2010.09.03
'*          �i���������ٌ���    2010.09.03
'*          �i���������i�[����  2010.09.03
'*          ������
'*          JAN����             2011.08.18
'*          ����                2014.07.01
'*          ���i�������O������  2015.11.07
'*
'*  �߂�l: false       :����
'*          true        :�p���\�Ȉُ�
'*          SYS_ERR     :�p���ł��Ȃ��ُ�
'*          SYS_CANCEL  :�X�V��ݾ�
'*
'****************************************************
Dim sts                 As Integer
Dim Sumi_Zaiko_Qty      As Long
Dim Mi_Zaiko_Qty        As Long

Dim RETRY_CNT           As Integer
Dim MESG_FLG            As Integer
Dim RETRY_SU            As Integer
    
Dim ans                 As Integer
                                            
    P_SAGYO_LOG_OUTPUT_PROC = True
                                            
    MESG_FLG = CInt(Mid(Format(RETRY, "00"), 1, 1))
    RETRY_SU = CInt(Mid(Format(RETRY, "00"), 2, 1))
                                            
                                        '��ƃ��O�o��
    Call UniCode_Conv(P_SAGYO_LOG_REC.JITU_DT, Format(Now, "YYYYMMDD"))     '���ѓ��t
    Call UniCode_Conv(P_SAGYO_LOG_REC.JITU_TM, Format(Now, "HHMMSS"))       '���ю���
    
    Call UniCode_Conv(P_SAGYO_LOG_REC.TANTO_CODE, TANTO_CODE)               '�S���Һ���
    Call UniCode_Conv(P_SAGYO_LOG_REC.WEL_ID, WEL_ID)                       '�[��ID
    Call UniCode_Conv(P_SAGYO_LOG_REC.JGYOBU, JGYOBU)                       '���ƕ�
    Call UniCode_Conv(P_SAGYO_LOG_REC.NAIGAI, NAIGAI)                       '�����O
    Call UniCode_Conv(P_SAGYO_LOG_REC.MENU_NO, MENU_NO)                     'TOP�ƭ�
    Call UniCode_Conv(P_SAGYO_LOG_REC.RIRK_ID, YOIN)                        '��Ɨv��
    
    Call UniCode_Conv(P_SAGYO_LOG_REC.ID_NO, ID_NO)                         '�`�[ID
    Call UniCode_Conv(P_SAGYO_LOG_REC.HIN_GAI, HIN_GAI)                     '�i��
    If Trim(HINBAN_DAMMY) = "." Then                                        '2017.10.30
        Call UniCode_Conv(P_SAGYO_LOG_REC.HIN_GAI, ".")
    End If
                                                                            '���i���ϕ����ѐ���
    
    If SUMI_QTY >= 0 Then
        Call UniCode_Conv(P_SAGYO_LOG_REC.SUMI_JITU_QTY, Format(SUMI_QTY, "00000000"))
    Else
        Call UniCode_Conv(P_SAGYO_LOG_REC.SUMI_JITU_QTY, Format(SUMI_QTY, "0000000"))
    End If
                                                                            
                                                                            '�����i�����ѐ���
    If MI_QTY >= 0 Then
        Call UniCode_Conv(P_SAGYO_LOG_REC.MI_JITU_QTY, Format(MI_QTY, "00000000"))
    Else
        Call UniCode_Conv(P_SAGYO_LOG_REC.MI_JITU_QTY, Format(MI_QTY, "0000000"))
    End If
    Call UniCode_Conv(P_SAGYO_LOG_REC.MUKE_CODE, MTS)                       'MTS
    Call UniCode_Conv(P_SAGYO_LOG_REC.SS_CODE, SS)                          'SS
    
    Call UniCode_Conv(P_SAGYO_LOG_REC.FROM_SOKO, Mid(FROM_LOCATION, 1, 2))  'FROM �I��
    Call UniCode_Conv(P_SAGYO_LOG_REC.FROM_RETU, Mid(FROM_LOCATION, 3, 2))  'FROM �I��
    Call UniCode_Conv(P_SAGYO_LOG_REC.FROM_REN, Mid(FROM_LOCATION, 5, 2))   'FROM �I��
    Call UniCode_Conv(P_SAGYO_LOG_REC.FROM_DAN, Mid(FROM_LOCATION, 7, 2))   'FROM �I��
        
    Call UniCode_Conv(P_SAGYO_LOG_REC.TO_SOKO, Mid(TO_LOCATION, 1, 2))      'TO �I��
    Call UniCode_Conv(P_SAGYO_LOG_REC.TO_RETU, Mid(TO_LOCATION, 3, 2))      'TO �I��
    Call UniCode_Conv(P_SAGYO_LOG_REC.TO_REN, Mid(TO_LOCATION, 5, 2))       'TO �I��
    Call UniCode_Conv(P_SAGYO_LOG_REC.TO_DAN, Mid(TO_LOCATION, 7, 2))       'TO �I��
                                                                            '�o�͌���۸���
    Call UniCode_Conv(P_SAGYO_LOG_REC.PRG_ID, StrConv(App.EXEName, vbUpperCase))
        
    Call UniCode_Conv(P_SAGYO_LOG_REC.WORK_TM, "")
        
        
    Call UniCode_Conv(P_SAGYO_LOG_REC.SHIJI_No, SHIJI_No)                   '�w���� 2010.09.03
                                                                            '�i���������ٌ��� 2010.09.03
    Call UniCode_Conv(P_SAGYO_LOG_REC.HIN_CHECK_LABEL_CNT, HIN_CHECK_LABEL_CNT)
                                                                            '�i���������i�[���� 2010.09.03
    Call UniCode_Conv(P_SAGYO_LOG_REC.HIN_CHECK_GENPIN_CNT, HIN_CHECK_GENPIN_CNT)
        
        
        
    Call UniCode_Conv(P_SAGYO_LOG_REC.FILLER, "")
        
    '2011.01.19
    If Trim(wkMTS) <> "" Then
        Call UniCode_Conv(P_SAGYO_LOG_REC.MUKE_CODE, wkMTS)                     'MTS
    End If
        
        
        
    '2011.08.18
    Call UniCode_Conv(P_SAGYO_LOG_REC.JAN_CODE, JAN_CODE)                   'JAN����    2011.08.18
        
        
    '2014.07.01
    Call UniCode_Conv(P_SAGYO_LOG_REC.MEMO, MEMO)                           '����    2014.07.01
        
    '2015.11.07
    Call UniCode_Conv(P_SAGYO_LOG_REC.HIN_CHECK_GAISOU_CNT, HIN_CHECK_GAISOU_CNT)
        
        
    RETRY_CNT = 0
    Do
        
        sts = BTRV(BtOpInsert, P_SAGYO_LOG_POS, P_SAGYO_LOG_REC, Len(P_SAGYO_LOG_REC), K0_P_SAGYO_LOG, Len(K0_P_SAGYO_LOG), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        '���g���C�񐔃`�F�b�N
                If RETRY_SU <> 0 Then
                    
                    RETRY_CNT = RETRY_CNT + 1
                    If RETRY_CNT > RETRY_SU Then
                                        '�񐔃I�[�o�[
                        Call File_Error(sts, BtOpInsert, "��ƃ��O", 0)
                        P_SAGYO_LOG_OUTPUT_PROC = SYS_CANCEL
                        Exit Function
                    
                    End If
                
                End If
                
                If MESG_FLG = 0 Then
'                    DoEvents
                    If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                        DoEvents                                                    '2016.01.26
                    End If                                                          '2016.01.26
                Else
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_SAGYO_LOG.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        P_SAGYO_LOG_OUTPUT_PROC = SYS_CANCEL
                        Exit Function
                    End If
                End If
            
            
            Case BtErrDEAD_LOCK '�f�b�h���b�N   2010.11.10
                P_SAGYO_LOG_OUTPUT_PROC = SYS_CANCEL
                Exit Function
            
            
            
            Case Else
                Call File_Error(sts, BtOpInsert, "���۸�")
                P_SAGYO_LOG_OUTPUT_PROC = SYS_ERR
                Exit Function
        End Select
    Loop
                            '����I��
    P_SAGYO_LOG_OUTPUT_PROC = False

End Function
