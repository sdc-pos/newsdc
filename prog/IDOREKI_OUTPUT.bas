Attribute VB_Name = "IDOREKI_OUTPUT"
Option Explicit
'---------------------------------------------- *�X�V�p�i�ڃ��[�N
'�|�W�V���j���O
Public wITEM_POS    As POSBLK
'�f�[�^�E�o�b�t�@
Public wITEMREC     As ITEMREC_Tag
'�L�[�E�f�[�^
Public K0_wITEM     As KEY0_ITEM

Public Wel_S_SHOUHI             As String * 2       '�uWEL ���ޏ���v�̗v�� 2007.06.28          2015.03.03 �ړ�
Public Wel_S_SHOUHI2            As String * 2      '�uWEL ���ޏ���(�V)�v�̗v�� 2015.02.21                  �ړ�


Public Function IDOREKI_OUTPUT_PROC(FROM_LOCATION As String, _
                                    TO_LOCATION As String, _
                                    JGYOBU As String, _
                                    NAIGAI As String, _
                                    HIN_GAI As String, _
                                    NYUKA_DT As String, _
                                    YOIN As String, _
                                    SUMI_JITU_QTY As Long, _
                                    MI_JITU_QTY As Long, _
                                    ID As String, _
                                    TANTO_CODE As String, _
                                    Optional RETRY As Integer = 10, _
                                    Optional CYU_KBN As String = " ", _
                                    Optional MEMO As String = "          ", _
                                    Optional Ins_DateTime As String = "              ", _
                                    Optional SHIIRE_CODE As String = "     ", _
                                    Optional SHIIRE_TANKA As String = "           ", _
                                    Optional KEIJYO_YM As String = "      ", _
                                    Optional MENU_NO As String = "  ", _
                                    Optional MUKE_CODE As String = "        ", _
                                    Optional SS_CODE As String = "        ", _
                                    Optional ID_NO As String = "        ", _
                                    Optional BIN_NO As String = "  ", _
                                    Optional DEN_NO As String = "      ", _
                                    Optional DEN_YMD As String = "        ", Optional LOG_MODE As Integer = 0, Optional GENSANKOKU As String = "                    ", Optional SHIIRE_WORK_CENTER As String = "        ", Optional ID_NO2 As String = "            ", Optional YOSAN_FROM As String = "     ", Optional YOSAN_TO As String = "     ", Optional wkMTS As String = "        ", Optional SEK_TEI_LABELID As String = "             ", Optional HINBAN_DAMMY As String = "                    ") As Integer
'****************************************************
'*      �݌Ɉړ����f�[�^�X�V
'*
'*  �݌Ɉړ����̍X�V���s���B
'*  (�����̐ݒ�~�X�͂�����ł̓`�F�b�N���Ȃ�)
'*  �����F  FROM�I�iXXXXXXXX(�q�ɇ�+��+�A+�i)�ȗ��j  ��FROM/TO���ꂩ�K�{
'*          TO�I�iXXXXXXXX(�q�ɇ�+��+�A+�i)�ȗ��j    ��FROM/TO���ꂩ�K�{
'*          ���ƕ��i�ȗ��s�j
'*          �����O�i�ȗ��s�j
'*          �O���i�ԁi�ȗ��s�j
'*          ���ד�(YYYYMMDD �ȗ��s��)
'*          �v��(�ȗ��s��)
'*          ���i���ςݎ��ѐ��i=0���Ƃ���A�����̂ݏo�́j
'*          �����i���ѐ��i=0���Ƃ���A�����̂ݏo�́j
'*          ID(�K�{)
'*          �S����
'*          ���g���C(�ȗ��� �P����:1=��ʃ��b�Z�[�W�L 0:���C�Q����:���g���C��(0�`9 0:����))
'*          �����敪�i�ȗ��j
'*          ����(�ȗ��� �����ɏo�͂��郁�����e)
'*          �ް��ǉ������i�ȗ��@�ް��쐬�������ꌳ�j
'*          �d���溰�ށi���ޗp�ȗ��@2006.01.05�j
'*          �d���P���i���ޗp�ȗ��@2006.01.05�j
'*          �v��N���i���ޗp�ȗ��@2006.01.05�j
'*          TOP�ƭ�(�����p 2006.01.30)
'*          ������(�����p 2006.01.30)
'*          ������(�����p 2006.01.30)
'*          �`�[ID(�����p 2006.01.30)
'*          �և�   (���PC�o�� 2007.05.16)
'*
'*
'*          ���O�o�̓��[�h (0:�����ŏo�� 1:�����ł͏o�͂��Ȃ�)
'*  �߂�l: false       :����
'*          true        :�p���\�Ȉُ�
'*          SYS_ERR     :�p���ł��Ȃ��ُ�
'*          SYS_CANCEL  :�X�V��ݾ�
'*
'*  ���o�ח\��^�݌Ƀf�[�^�^������Ǘ��}�X�^�͌Ăь��œǍ��ݍς݂̎�
'****************************************************
Dim sts                 As Integer
Dim Sumi_Zaiko_Qty      As Long
Dim Mi_Zaiko_Qty        As Long

Dim RETRY_CNT           As Integer
Dim MESG_FLG            As Integer
Dim RETRY_SU            As Integer
    
Dim ans                 As Integer
                                            
    IDOREKI_OUTPUT_PROC = True
                                            
    MESG_FLG = CInt(Mid(Format(RETRY, "00"), 1, 1))
    RETRY_SU = CInt(Mid(Format(RETRY, "00"), 2, 1))
                            '�o�ɕ\�^�`�[ID�o�׎��̗v����u��������
    If Left(YOIN, 1) = ACT_DENPYO_ID Or _
        Left(YOIN, 1) = ACT_SYUKA_HYO Or _
        Left(YOIN, 1) = ACT_DENPYO_ID2 Then 'ACT_DENPYO_ID2 �ǉ��@2015.02.21
        YOIN = ACT_SYUKA_KEI & CYU_KBN
    End If
                            
                            '�i�ڃ}�X�^�ǂݍ���
    Call UniCode_Conv(K0_wITEM.JGYOBU, JGYOBU)
    Call UniCode_Conv(K0_wITEM.NAIGAI, NAIGAI)
    Call UniCode_Conv(K0_wITEM.HIN_GAI, HIN_GAI)
    sts = BTRV(BtOpGetEqual, wITEM_POS, wITEMREC, Len(wITEMREC), K0_wITEM, Len(K0_wITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(wITEMREC.HIN_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
            IDOREKI_OUTPUT_PROC = SYS_ERR
            Exit Function
    End Select
    
'2008.08.06    Call UniCode_Conv(IDOREC.JITU_DT, Format(Now, "yyyymmdd"))              '���ѓ��t
'2008.08.06    Call UniCode_Conv(IDOREC.JITU_TM, Format(Now, "HHmmss"))                '���ю���
    
    
    If YOIN = Wel_S_SHOUHI2 Then                                            '�V���ޏ���̗v���u�������@2015.02.21
        YOIN = Wel_S_SHOUHI
    End If
    
    
    
    If Trim(Ins_DateTime) = "" Then     '2008.09.01
    
        Call UniCode_Conv(IDOREC.JITU_DT, Format(Now, "yyyymmdd"))              '���ѓ��t
        Call UniCode_Conv(IDOREC.JITU_TM, Format(Now, "HHmmss"))                '���ю���
    
    Else
    
    
        Call UniCode_Conv(IDOREC.JITU_DT, Left(Ins_DateTime, 8))                '���ѓ��t   2008.08.06
        Call UniCode_Conv(IDOREC.JITU_TM, Right(Ins_DateTime, 6))               '���ю���   2008.08.06
    
    End If
    
    
    
    Call UniCode_Conv(IDOREC.JGYOBU, JGYOBU)                                '���ƕ�
    Call UniCode_Conv(IDOREC.NAIGAI, NAIGAI)                                '�����O
    Call UniCode_Conv(IDOREC.HIN_GAI, HIN_GAI)                              '�i�ԁi�O���j
    
    Call UniCode_Conv(IDOREC.RIRK_ID, YOIN)                                 '�������
    
                                                                            
                                                                            
                                                                            '���ѐ���(���i���ς�)
    Call UniCode_Conv(IDOREC.SUMI_JITU_QTY, Format(SUMI_JITU_QTY, "00000000"))
                                                                            '���ѐ���(�����i)
    Call UniCode_Conv(IDOREC.MI_JITU_QTY, Format(MI_JITU_QTY, "00000000"))
    Call UniCode_Conv(IDOREC.FROM_SOKO, Mid(FROM_LOCATION, 1, 2))           'FROM �q�ɇ�
    Call UniCode_Conv(IDOREC.FROM_RETU, Mid(FROM_LOCATION, 3, 2))           'FROM ��
    Call UniCode_Conv(IDOREC.FROM_REN, Mid(FROM_LOCATION, 5, 2))            'FROM �A
    Call UniCode_Conv(IDOREC.FROM_DAN, Mid(FROM_LOCATION, 7, 2))            'FROM �i
    Call UniCode_Conv(IDOREC.TO_SOKO, Mid(TO_LOCATION, 1, 2))               'TO �q�ɇ�
    Call UniCode_Conv(IDOREC.TO_RETU, Mid(TO_LOCATION, 3, 2))               'TO ��
    Call UniCode_Conv(IDOREC.TO_REN, Mid(TO_LOCATION, 5, 2))                'TO �A
    Call UniCode_Conv(IDOREC.TO_DAN, Mid(TO_LOCATION, 7, 2))                'TO �i
    Call UniCode_Conv(IDOREC.PRG_ID, StrConv(App.EXEName, vbUpperCase))     '�o�͌��v���O����
            
            
            
            
            
'''2011.02.03    If YOIN = YOIN_TANASHOGO Or YOIN = YOIN_TANAHINSHOGO Then
'�i�ԕʏƍ���ǉ�   2011.02.03
    If YOIN = YOIN_TANASHOGO Or YOIN = YOIN_TANAHINSHOGO Or YOIN = YOIN_HIN_SHOGO Then
       '�v�����I�ƍ��̎��͕s��ƂȂ�
                                                    '�i�ԁi�����j
        Call UniCode_Conv(IDOREC.HIN_NAI, StrConv(wITEMREC.HIN_NAI, vbUnicode))
                                                    '���ɓ�
        Call UniCode_Conv(IDOREC.NYUKO_DT, "")
    Else
                                                        '�i�ԁi�����j
        Call UniCode_Conv(IDOREC.HIN_NAI, StrConv(ZAIKOREC.HIN_NAI, vbUnicode))
                                                            '���ɓ�
        Call UniCode_Conv(IDOREC.NYUKO_DT, StrConv(ZAIKOREC.NYUKO_DT, vbUnicode))
   End If
    
    
    
    If Trim(FROM_LOCATION) = "" And Trim(TO_LOCATION) = "" Then                     '2014.03.05
        Call UniCode_Conv(IDOREC.HIN_NAI, StrConv(wITEMREC.HIN_NAI, vbUnicode))     '2014.03.05
                                                    '���ɓ�
        Call UniCode_Conv(IDOREC.NYUKO_DT, "")                                      '2014.03.05
    End If
    
    
    
    Call UniCode_Conv(IDOREC.NYUKA_DT, NYUKA_DT)
    Call UniCode_Conv(IDOREC.WEL_ID, ID)                                    '�[��ID
    
    
    
    Call UniCode_Conv(K0_YOIN.CODE_TYPE, Mid(YOIN, 1, 1))                   '���𖼏�
    Call UniCode_Conv(K0_YOIN.YOIN_CODE, Mid(YOIN, 2, 1))
                                            '�v��Ͻ��Ǎ���
    sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    Select Case sts
        Case BtNoErr
            Call UniCode_Conv(IDOREC.RIRK_NAME, StrConv(YOINREC.YOIN_DNAME, vbUnicode))
            Call UniCode_Conv(IDOREC.SUM_KBN, StrConv(YOINREC.SUM_KBN, vbUnicode))
        Case BtErrKeyNotFound
            Call UniCode_Conv(IDOREC.RIRK_NAME, "")
                                            '�s���̂Ƃ��͍ݒ�����
            Call UniCode_Conv(IDOREC.SUM_KBN, SUM_KBN_ZT)
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�v��Ͻ�")
            IDOREKI_OUTPUT_PROC = SYS_ERR
            Exit Function
    End Select
                                                                            '�i�ږ���
    Call UniCode_Conv(IDOREC.HIN_NAME, StrConv(wITEMREC.HIN_NAME, vbUnicode))
                                                                            '�i�ڕʍ݌ɐ�
    If Zaiko_Syukei_Proc(Sumi_Zaiko_Qty, Mi_Zaiko_Qty, JGYOBU, NAIGAI, HIN_GAI) Then
        IDOREKI_OUTPUT_PROC = SYS_ERR
        Exit Function
    End If
    Call UniCode_Conv(IDOREC.SUMI_HIN_Zaiko_Qty, Format(Sumi_Zaiko_Qty, "00000000"))
    Call UniCode_Conv(IDOREC.MI_HIN_Zaiko_Qty, Format(Mi_Zaiko_Qty, "00000000"))
                                                                            
                                                                            'FROM�I�ʕi�ڕʍ݌ɐ�
    If Len(Trim(FROM_LOCATION)) <> 0 Then
        If Zaiko_Syukei_Proc(Sumi_Zaiko_Qty, Mi_Zaiko_Qty, JGYOBU, NAIGAI, HIN_GAI, FROM_LOCATION) Then
            IDOREKI_OUTPUT_PROC = SYS_ERR
            Exit Function
        End If
        Call UniCode_Conv(IDOREC.SUMI_FROM_TANA_Zaiko_Qty, Format(Sumi_Zaiko_Qty, "00000000"))
        Call UniCode_Conv(IDOREC.MI_FROM_TANA_Zaiko_Qty, Format(Mi_Zaiko_Qty, "00000000"))
    Else
        Call UniCode_Conv(IDOREC.SUMI_FROM_TANA_Zaiko_Qty, "00000000")
        Call UniCode_Conv(IDOREC.MI_FROM_TANA_Zaiko_Qty, "00000000")
    End If
                                                                            'TO�I�ʕi�ڕʍ݌ɐ�
    If Len(Trim(TO_LOCATION)) <> 0 Then
        If Zaiko_Syukei_Proc(Sumi_Zaiko_Qty, Mi_Zaiko_Qty, JGYOBU, NAIGAI, HIN_GAI, TO_LOCATION) Then
            IDOREKI_OUTPUT_PROC = SYS_ERR
            Exit Function
        End If
        Call UniCode_Conv(IDOREC.SUMI_TO_TANA_Zaiko_Qty, Format(Sumi_Zaiko_Qty, "00000000"))
        Call UniCode_Conv(IDOREC.MI_TO_TANA_Zaiko_Qty, Format(Mi_Zaiko_Qty, "00000000"))
    Else
        Call UniCode_Conv(IDOREC.SUMI_TO_TANA_Zaiko_Qty, "00000000")
        Call UniCode_Conv(IDOREC.MI_TO_TANA_Zaiko_Qty, "00000000")
    End If
                                        
    If CYU_KBN = " " Then
        Call UniCode_Conv(IDOREC.DEN_DT, "")
        Call UniCode_Conv(IDOREC.DEN_NO, "")
        Call UniCode_Conv(IDOREC.TOKU_MARK, "")
        Call UniCode_Conv(IDOREC.MUKE_CODE, "")
        Call UniCode_Conv(IDOREC.MUKE_CHG_CD, "")
        Call UniCode_Conv(IDOREC.MUKE_DNAME, "")
        Call UniCode_Conv(IDOREC.SS_CODE, "")
        Call UniCode_Conv(IDOREC.SS_NAME, "")
        Call UniCode_Conv(IDOREC.ID_NO, "")
    
        If StrConv(IDOREC.RIRK_ID, vbUnicode) = YOIN_TU_NYUKA Then
            Call UniCode_Conv(IDOREC.ID_NO, ID_NO2)
        End If
    
    Else
        Call UniCode_Conv(IDOREC.DEN_DT, StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode))
        Call UniCode_Conv(IDOREC.DEN_NO, StrConv(Y_SYUREC.DEN_NO, vbUnicode))
        If CYU_KBN = CYU_KBN_SPO And _
            StrConv(Y_SYUREC.TOK_KBN, vbUnicode) = "1" Then
            Call UniCode_Conv(IDOREC.TOKU_MARK, "*")                        '������}�[�N
        Else
            Call UniCode_Conv(IDOREC.TOKU_MARK, "")
        End If
        Call UniCode_Conv(IDOREC.ID_NO, StrConv(Y_SYUREC.ID_NO, vbUnicode))
    
    
                                                                                '������R�[�h
        Call UniCode_Conv(IDOREC.MUKE_CODE, StrConv(MTSREC.MUKE_CODE, vbUnicode))
                                                                                '�����於��
        Call UniCode_Conv(IDOREC.MUKE_DNAME, StrConv(MTSREC.MUKE_DNAME, vbUnicode))
                                                                                '�r�r�R�[�h
        Call UniCode_Conv(IDOREC.SS_CODE, StrConv(MTSREC.SS_CODE, vbUnicode))
                                                                                '�r�r����
        Call UniCode_Conv(IDOREC.SS_NAME, StrConv(MTSREC.SS_NAME, vbUnicode))
                                                                                '���Ӑ旪��
        Call UniCode_Conv(IDOREC.MUKE_DNAME, StrConv(MTSREC.MUKE_DNAME, vbUnicode))
    
    
    End If
                                                                        
                                                                        
    '2009.03.18
    If Left(YOIN, 1) = ACT_BINNO Then
        Call UniCode_Conv(IDOREC.SS_CODE, SS_CODE)
        Call UniCode_Conv(IDOREC.SS_NAME, SS_CODE)
    End If
                                                                        
                                                                        
    Call UniCode_Conv(IDOREC.MEMO, MEMO)                                    '����
                                                                            
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, TANTO_CODE)                      '�S����
                                            '�S����Ͻ��Ǎ���
    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(TANTOREC.TANTO_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�S����Ͻ�")
            IDOREKI_OUTPUT_PROC = SYS_ERR
            Exit Function
    End Select
                                                                            '�S���҃R�[�h
    Call UniCode_Conv(IDOREC.TANTO_CODE, StrConv(TANTOREC.TANTO_CODE, vbUnicode))
                                                                            '�S���Җ���
    Call UniCode_Conv(IDOREC.TANTO_NAME, StrConv(TANTOREC.TANTO_NAME, vbUnicode))
                                                                            
                                                                            '�}������
    If Len(Trim(Ins_DateTime)) = 0 Then
        Ins_DateTime = StrConv(IDOREC.JITU_DT, vbUnicode) & StrConv(IDOREC.JITU_TM, vbUnicode)
    End If
    
    Call UniCode_Conv(IDOREC.Ins_DateTime, Ins_DateTime)
                                            
                                            
    Call UniCode_Conv(IDOREC.SHIIRE_CODE, SHIIRE_CODE)          '�d������2006.01.06
    Call UniCode_Conv(IDOREC.SHIIRE_TANKA, SHIIRE_TANKA)        '�d���P��2006.01.06
    Call UniCode_Conv(IDOREC.KEIJYO_YM, KEIJYO_YM)              '�v��N��2006.01.06
                                            
    Call UniCode_Conv(IDOREC.BIN_NO, BIN_NO)                    '�և� 2007.05.16
                                            
                                            
    If Trim(DEN_NO) <> "" Then                                  '���PC�@���ד`�[�� 2007.06.07
        Call UniCode_Conv(IDOREC.DEN_NO, DEN_NO)
    End If
    If Trim(DEN_YMD) <> "" Then                                 '���PC�@���ד`�[���t 2007.06.07
        Call UniCode_Conv(IDOREC.DEN_DT, DEN_YMD)
    End If
                                            
                                            
                                            
                                            
                                            
                                            
    '----------------   2010.07.08 ��
    Call UniCode_Conv(IDOREC.GENSANKOKU, GENSANKOKU)            '���Y����
                                                                '���ގd����ܰ�����
    Call UniCode_Conv(IDOREC.SHIIRE_WORK_CENTER, SHIIRE_WORK_CENTER)
    Call UniCode_Conv(IDOREC.ID_NO2, ID_NO2)                    'ID_NO
    Call UniCode_Conv(IDOREC.YOSAN_FROM, YOSAN_FROM)            '�\�Z�P�ʁi���j
    Call UniCode_Conv(IDOREC.YOSAN_TO, YOSAN_TO)                '�\�Z�P�ʁi��j
    '----------------   2010.07.08 ��
                                            
    
    '----------------   2011.04.29 ��
    If Trim(SEK_TEI_LABELID) <> "" Then
        Call UniCode_Conv(IDOREC.ID_NO, SEK_TEI_LABELID)
    End If
    '----------------   2011.04.29 ��
                                            
                                            
                                            
                                            
                                            
                                            
                                            
    Call UniCode_Conv(IDOREC.FILLER, "")
                                            
                                        '�݌Ɉړ����o��
    
    '�v��=���i���� & �I��TO <>"" 2019/12/25 ���i�������o�^�݌Ɍv�㎞ �ړ������ɓ��ɐ���\��
    If YOIN = "M8" And Trim(TO_LOCATION) <> "" Then '2020/03/16 "M8" ���i�������o�^�v�� ���ߑł��ɏC��
         Call UniCode_Conv(IDOREC.SUM_KBN, SUM_KBN_IN)
    End If
    
    RETRY_CNT = 0
    Do
        
        sts = BTRV(BtOpInsert, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        '���g���C�񐔃`�F�b�N
                If RETRY_SU <> 0 Then
                    
                    RETRY_CNT = RETRY_CNT + 1
                    If RETRY_CNT > RETRY_SU Then
                                        '�񐔃I�[�o�[
                        Call File_Error(sts, BtOpInsert, "�݌Ɉړ���", 0)
                        IDOREKI_OUTPUT_PROC = SYS_CANCEL
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
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<IDOREKI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        IDOREKI_OUTPUT_PROC = SYS_CANCEL
                        Exit Function
                    End If
                End If
            
            Case BtErrDEAD_LOCK
                IDOREKI_OUTPUT_PROC = SYS_CANCEL
                Exit Function
            
            Case Else
                Call File_Error(sts, BtOpInsert, "�݌Ɉړ���")
                IDOREKI_OUTPUT_PROC = SYS_ERR
                Exit Function
        End Select
    Loop
                            

    
    
    If LOG_MODE = 0 Then
        If Trim(MENU_NO) = "" Then
        Else
        '���۸ޏo��
            
            If Trim(MUKE_CODE) <> "" Then
                Call UniCode_Conv(IDOREC.MUKE_CODE, MUKE_CODE)
            End If
            
            If Trim(SS_CODE) <> "" Then
                Call UniCode_Conv(IDOREC.SS_CODE, SS_CODE)
            End If
            
            If Trim(ID_NO) <> "" Then
                Call UniCode_Conv(IDOREC.ID_NO, ID_NO)
            End If
            
            
            If App.EXEName = "F102015" Then
                Call UniCode_Conv(IDOREC.ID_NO, ID_NO2)
            End If
            
            If YOIN = RYOHEN Then
                Call UniCode_Conv(IDOREC.MUKE_CODE, Left(StrConv(IDOREC.MEMO, vbUnicode), 4))
            End If
            
            If P_SAGYO_LOG_OUTPUT_PROC(TANTO_CODE, _
                                                ID, _
                                                JGYOBU, _
                                                NAIGAI, _
                                                MENU_NO, _
                                                YOIN, _
                                                HIN_GAI, _
                                                SUMI_JITU_QTY, _
                                                MI_JITU_QTY, _
                                                FROM_LOCATION, _
                                                TO_LOCATION, _
                                                StrConv(IDOREC.ID_NO, vbUnicode), _
                                                StrConv(IDOREC.MUKE_CODE, vbUnicode), _
                                                StrConv(IDOREC.SS_CODE, vbUnicode), _
                                                RETRY, , , , wkMTS, , , , HINBAN_DAMMY) Then
                IDOREKI_OUTPUT_PROC = SYS_ERR
                Exit Function
            End If
        End If
    End If
                            
                            '����I��
    IDOREKI_OUTPUT_PROC = False

End Function
                    
Public Function wITEM_Open(Mode As Integer) As Integer
'****************************************************
'*      �u�ړ����o�͏����v    �i�ڂn�o�d�m����
'*
'*  �i�ڃ}�X�^��ʃ|�C���^�łn�o�d�m����
'*  (�Ăь��ŋN�����ɂP�x�����Ăяo��)
'*
'*  �߂�l: false       :����
'*          true        :�ُ�
'*          SYS_CANCEL  :�X�V��ݾ�
'****************************************************
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String

Dim ans         As Integer
    
    
    wITEM_Open = True
                                '�i�ڃ}�X�^�@�t���p�X�捞��
    sts = GetIni("FILE", ITEM_ID, "SYS", c)
    
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, wITEM_POS, wITEMREC, Len(wITEMREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
'-------------- �n�o�d�m�����ł̎g�p���́A�����グ���ɂP�񂾂��̂͂��Ȃ̂ŁA��ɉ�ʓ��͂Ƃ��A
'               ��ݾق́A�����̋N����ݾقƂ���B
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    wITEM_Open = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpOpen, "�i�ڃ}�X�^")
                Exit Function
        End Select
    Loop

    wITEM_Open = False

End Function
