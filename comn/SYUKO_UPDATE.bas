Attribute VB_Name = "SYUKO_UPDATE"
Option Explicit
'---------------------------------------------- *�X�V�p�o�ח\�胏�[�N
'�|�W�V���j���O
Public wY_SYU_POS   As POSBLK
'�f�[�^�E�o�b�t�@
Public wY_SYUREC    As Y_SYUREC_Tag
'�L�[�E�f�[�^
Public K3_wY_SYU    As KEY3_Y_SYU


Public Function Syuko_Update_Proc(JGYOBU As String, _
                                    NAIGAI As String, _
                                    HIN_GAI As String, _
                                    NYUKA_DT As String, _
                                    FROM_LOCATION As String, _
                                    YOIN As String, _
                                    SUMI_JITU_QTY As Long, _
                                    MI_JITU_QTY As Long, _
                                    SYUKA_QTY As Long, _
                                    ID As String, _
                                    TANTO_CODE As String, _
                                    Optional RETRY As Integer = 10, _
                                    Optional MEMO As String = "          ", _
                                    Optional CYU_KBN As String = " ", _
                                    Optional MUKE_CODE As String = "                ", _
                                    Optional SYUKA_YMD As String = "        ", _
                                    Optional DEN_NO As String = "          ", _
                                    Optional ID_NO As String = "        ", _
                                    Optional MENU_NO As String = "  ", _
                                    Optional LOG_NON As Integer = 0) As Integer
'****************************************************
'*      �u�o�ׁ^�o�ɏ����v�݌Ƀf�[�^�X�V
'*
'*  �݌Ƀf�[�^�̍X�V���s���B
'*  (�����̐ݒ�~�X�͂�����ł̓`�F�b�N���Ȃ�)
'*  ���g�����U�N�V�����������K�v�ȏꍇ�͌Ăь��ōs����
'*  �g�p̧��    :   �݌Ƀf�[�^
'*                  �i�ڃ}�X�^
'*                  �v���}�X�^
'*                  ������}�X�^
'*                  �o�ח\��f�[�^
'*                  �݌Ɉړ���
'*  �����F  ���ƕ��i�ȗ��s�j
'*          �����O�i�ȗ��s�j
'*          �i�ԊO��(�ȗ��s��)
'*          FROM�I�ԁiXXXXXXXX(�q�ɇ�+��+�A+�i)�ȗ��s�j
'*          ���ד�(YYYYMMDD �ȗ��� �ȗ���FIFO)
'*          �v��(�ȗ��s��)
'*          ���i���ςݎ��ѐ��i�����ꂩ�K�{�j
'*          �����i���ѐ�    �i�@�@�@�V�@�@�j
'*          �o�א���        �i            �j
'*          ID(�ȗ��s��)
'*          �S���ҁi�ȗ��s�j
'*          ���g���C(�ȗ��� �P����:1=��ʃ��b�Z�[�W�L 0:���C�Q����:���g���C��(0�`9 0:����))
'*          ����(�ȗ��� �����ɏo�͂��郁�����e)
'*          �����敪�i�o�׎��K�{�j
'*          �`�[�h�c�i�o�׎��K�{�j
'*          �ƭ���ٰ�߁i�����Ǘ����ځj  2006.01.30
'*  �߂�l: false       :����
'*          true        :�p���\�Ȉُ�
'*          SYS_ERR     :�p���ł��Ȃ��ُ�
'*          SYS_CANCEL  :�X�V��ݾ�
'****************************************************
Dim sts         As Integer
Dim com         As Integer
Dim Upd_com     As Integer


Dim RETRY_CNT   As Integer
Dim MESG_FLG    As Integer
Dim RETRY_SU    As Integer
    
Dim ans         As Integer
    
Dim Zan_Qty     As Long
Dim WK_Qty      As Long
    
Dim GET_DEN_NO  As String * 6
Dim GET_ID_NO   As String * 12
    
Dim JITU_QTY    As Long
Dim GOODS_F     As String * 1
    
Dim Wk_SUMI_JITU_QTY As Long
Dim Wk_MI_JITU_QTY As Long
Dim Wk_SYUKA_QTY As Long

Dim Ins_DateTime    As String * 14              '2004.12.09
    
Dim wkYOIN      As String * 2
    
    
'''''   2011.04.04
Dim Total_SUMI_JITU_QTY     As Long
Dim Total_MI_JITU_QTY       As Long
'''''   2011.04.04
    
    
    Syuko_Update_Proc = True
                                                                      
    MESG_FLG = CInt(Mid(Format(RETRY, "00"), 1, 1))
    RETRY_SU = CInt(Mid(Format(RETRY, "00"), 2, 1))
                                        
                                        
    Wk_SUMI_JITU_QTY = SUMI_JITU_QTY
    Wk_MI_JITU_QTY = MI_JITU_QTY
    Wk_SYUKA_QTY = SYUKA_QTY
                                        
    Ins_DateTime = Format(Now, "YYYYMMDDHHMMSS")            '2004.12.09
    '*------------------------------------------------------'�i��Ͻ��̊m��
    Call UniCode_Conv(K0_ITEM.JGYOBU, JGYOBU)               '���ƕ�
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI)               '���O
    Call UniCode_Conv(K0_ITEM.HIN_GAI, HIN_GAI)             '�i�ԁi�O���j
        
    RETRY_CNT = 0
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
                If MESG_FLG = 1 Then
                    Beep
                    MsgBox "�i�ڃR�[�h�����݂��܂���B�X�V�����𒆎~���܂��B", vbOKOnly, "�m�F����"
                End If
                Exit Function
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        '���g���C�񐔃`�F�b�N
                If RETRY_SU <> 0 Then
                    
                    RETRY_CNT = RETRY_CNT + 1
                    If RETRY_CNT > RETRY_SU Then
                                        '�񐔃I�[�o�[
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^", 0)
                        Syuko_Update_Proc = SYS_CANCEL
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
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Syuko_Update_Proc = SYS_CANCEL
                        Exit Function
                    End If
                End If
            Case BtErrDEAD_LOCK
                Syuko_Update_Proc = SYS_CANCEL
                Exit Function
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^")
                Syuko_Update_Proc = SYS_ERR
                Exit Function
        End Select
    Loop
    '*------------------------------------------------------'�o�׎��o�ח\��̊m��
    If CYU_KBN <> " " Then       '�o�׎�������Ͻ��Ǎ���
        Call UniCode_Conv(K0_MTS.MUKE_CODE, Left(MUKE_CODE, 8))
        Call UniCode_Conv(K0_MTS.SS_CODE, Right(MUKE_CODE, 8))
                                    '������}�X�^��ǂݍ��݌����溰�ނ�Ă���
        sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrKeyNotFound           '�L��Ƃ܂������G���[�ɂ��Ȃ�
                Call UniCode_Conv(MTSREC.MUKE_DNAME, "")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "������Ǘ�Ͻ�")
                Syuko_Update_Proc = SYS_ERR
                Exit Function
        End Select
    
    End If
    
    Select Case CYU_KBN
        Case " "
        Case CYU_KBN_KIN                '�ً}���A�o�ח\����N�[����
            Call UniCode_Conv(Y_SYUREC.WEL_ID, "")                      '�g�p�q�@�h�c
            Call UniCode_Conv(Y_SYUREC.PRG_ID, "")                      '�g�p���v���O����
            Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_KBN_UN)             '�����敪
            Call UniCode_Conv(Y_SYUREC.DT_SYU, "R")                     '�f�[�^���
            Call UniCode_Conv(Y_SYUREC.JGYOBU, Last_JGYOBU)             '���ƕ��敪
            Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, CYU_KBN_KIN)        '�����敪
            Call UniCode_Conv(Y_SYUREC.CYU_KBN, CYU_KBN_KIN)
            If Len(Trim(ID_NO)) <> 0 Then                               '�h�c��
                Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, ID_NO)
                Call UniCode_Conv(Y_SYUREC.ID_NO, ID_NO)
            Else
                sts = Den_No_Set_Proc(21, Last_JGYOBU, GET_ID_NO)
                If sts Then
                    Syuko_Update_Proc = SYS_ERR
                    Exit Function
                End If
                Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, GET_ID_NO)
                Call UniCode_Conv(Y_SYUREC.ID_NO, GET_ID_NO)
            End If
                                                                            
            Call UniCode_Conv(Y_SYUREC.NAIGAI, NAIGAI)                  '�����O
                                                                    
            Call UniCode_Conv(Y_SYUREC.KEY_HIN_NO, HIN_GAI)             '�i�ڔԍ�
            Call UniCode_Conv(Y_SYUREC.HIN_NO, HIN_GAI)                 '�i�ڔԍ�
                                                                        '���Ӑ�R�[�h
            Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, Left(MUKE_CODE, 8))
            Call UniCode_Conv(Y_SYUREC.MUKE_CODE, Left(MUKE_CODE, 8))
                                                                        '������R�[�h
            Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, Right(MUKE_CODE, 8))
            Call UniCode_Conv(Y_SYUREC.SS_CODE, Right(MUKE_CODE, 8))
                                                                        '�o�ד�
            Call UniCode_Conv(Y_SYUREC.KEY_SYUKA_YMD, SYUKA_YMD)
            Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, SYUKA_YMD)
            Call UniCode_Conv(Y_SYUREC.SYUKO_YMD, SYUKA_YMD)
    
    
    
            Call UniCode_Conv(Y_SYUREC.JGYOBA, "")                      '���Ə꺰��
            Call UniCode_Conv(Y_SYUREC.DATA_KBN, "")                    '�f�[�^�敪
            Call UniCode_Conv(Y_SYUREC.TORI_KBN, "")                    '����敪

                
            Call UniCode_Conv(Y_SYUREC.KAIKEI_JGYOBA, "")               '��v�p���Ə꺰��
            Call UniCode_Conv(Y_SYUREC.SHISAN_JGYOBA, "")               '���Y�Ǘ��p���Ə꺰��

            If Len(Trim(DEN_NO)) <> 0 Then
                Call UniCode_Conv(Y_SYUREC.DEN_NO, DEN_NO)
            Else
                sts = Den_No_Set_Proc(20, Last_JGYOBU, GET_DEN_NO)
                If sts Then
                    Syuko_Update_Proc = SYS_ERR
                    Exit Function
                End If
                Call UniCode_Conv(Y_SYUREC.DEN_NO, GET_DEN_NO)
        
            End If
    
                                                                        '�o�ɐ���
            Call UniCode_Conv(Y_SYUREC.SURYO, Format(SUMI_JITU_QTY + MI_JITU_QTY, "0000000"))
            Call UniCode_Conv(Y_SYUREC.SYUKO_SYUSI, "")                 '�݌Ɏ��x
            Call UniCode_Conv(Y_SYUREC.SHISAN_SYUSI, "")                '���Y�Ǘ��p�݌Ɏ��x����
            Call UniCode_Conv(Y_SYUREC.HOJYO_SYUSI, "")                 '�⏕�݌Ɏ��x����
            Call UniCode_Conv(Y_SYUREC.TANKA, "")                       '�P��
    
            Call UniCode_Conv(Y_SYUREC.ODER_NO, "")                     '�I�[�_�[�ԍ�
            Call UniCode_Conv(Y_SYUREC.ITEM_NO, "")                     '�A�C�e���ԍ�
            Call UniCode_Conv(Y_SYUREC.ODER_NO_R, "")                   '�����Ǘ��ԍ�����

            
            Call UniCode_Conv(Y_SYUREC.KOSO_KEITAI, "")                 '���`�Ժ���
            Call UniCode_Conv(Y_SYUREC.TANABAN1, "")                    '�I�ԂP
            Call UniCode_Conv(Y_SYUREC.TANABAN2, "")                    '�I�ԂQ
            Call UniCode_Conv(Y_SYUREC.TANABAN3, "")                    '�I�ԂR
                                                                        '���Ӑ於��
            Call UniCode_Conv(Y_SYUREC.MUKE_NAME, StrConv(MTSREC.MUKE_NAME, vbUnicode))
    
            Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, CYU_KBN_T)         '�����敪����
    
            Call UniCode_Conv(Y_SYUREC.ORIGIN1, "")                     '���Y��1
            Call UniCode_Conv(Y_SYUREC.ORIGIN2, "")                     '���Y��2
            Call UniCode_Conv(Y_SYUREC.BIKOU2, "")                      '���l2
  
    
            Call UniCode_Conv(Y_SYUREC.HAN_KBN, "")                     '�̔��敪
            Call UniCode_Conv(Y_SYUREC.CHOKU_KBN, "")                   '�����w���敪
            Call UniCode_Conv(Y_SYUREC.UNIT_ID_NO, "")                  '�ƯďC���Ǘ��ԍ�
            Call UniCode_Conv(Y_SYUREC.ZAIKO_HIKIATE, "")               '�݌Ɉ�������
            Call UniCode_Conv(Y_SYUREC.GOKON_KANRI_NO, "")              '�����Ǘ��ԍ�
            Call UniCode_Conv(Y_SYUREC.JYUCHU_ZAN, "")                  '�󒍎c����
            Call UniCode_Conv(Y_SYUREC.KYOKYU_KBN, "")                  '�����敪
            Call UniCode_Conv(Y_SYUREC.SHOHIN_SYUSI, "")                '���i���[�i�݌Ɏ��x����
            Call UniCode_Conv(Y_SYUREC.S_SHISAN_SYUSI, "")              '���i���[�i���Y�Ǘ����x����
            Call UniCode_Conv(Y_SYUREC.S_HOJYO_SYUSI, "")               '���i���[�i�⏕���x����
            Call UniCode_Conv(Y_SYUREC.BIKOU1, "")                      '���l1
            Call UniCode_Conv(Y_SYUREC.CHOHA_KBN, "")                   '���[�敪
            Call UniCode_Conv(Y_SYUREC.JYU_HIN_NO, "")                  '��t�i�ڔԍ�
                                                                        '�i��
            Call UniCode_Conv(Y_SYUREC.HIN_NAME, StrConv(ITEMREC.HIN_NAME, vbUnicode))
            Call UniCode_Conv(Y_SYUREC.HIN_CHANGE_KBN, "")              '�i�ڔԍ��ύX�敪
            Call UniCode_Conv(Y_SYUREC.MODULE_EXCHANGE, "")             'Ӽޭ�ٌ����敪
            Call UniCode_Conv(Y_SYUREC.ZAIKO_SYUSI, "")                 '�c�݌ɂ܂Ƃߍ݌Ɏ��x����
            Call UniCode_Conv(Y_SYUREC.ZAN_SHISAN_SYUSI, "")            '�c�݌ɂ܂Ƃߎ��Y�Ǘ����x����
            Call UniCode_Conv(Y_SYUREC.ZAN_HOJYO_SYUSI, "")             '�c�݌ɂ܂Ƃߕ⏕���x����
            Call UniCode_Conv(Y_SYUREC.NOUKI_YMD, "")                   '�w��[��
            Call UniCode_Conv(Y_SYUREC.SERVICE_KANRI_NO, "")            '���޽��ЊǗ��ԍ�
            Call UniCode_Conv(Y_SYUREC.KISHU_CODE, "")                  '�@��i�ں���
            Call UniCode_Conv(Y_SYUREC.ENVIRONMENT_KBN, "")             '����敔�i�敪
            Call UniCode_Conv(Y_SYUREC.KEPIN_KAIJYO, "")                '���i�����敪
                                                                        '�i�ԓ���
            Call UniCode_Conv(Y_SYUREC.HIN_NAI, StrConv(ITEMREC.HIN_NAI, vbUnicode))
                                                                                
                                                                        '�z�X�g�I��
            Call UniCode_Conv(Y_SYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode))
    
            Call UniCode_Conv(Y_SYUREC.PRINT_YMD, "")                   '������t
            Call UniCode_Conv(Y_SYUREC.KAN_YMD, "")                     '�������t
            Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, "")                  '���i���t
            Call UniCode_Conv(Y_SYUREC.TOK_KBN, "")                     '������敪
    
            Call UniCode_Conv(Y_SYUREC.JITU_SURYO, "00000000")          '���ѐ���
                                                                        '�捞�ݓ���
            Call UniCode_Conv(Y_SYUREC.INS_NOW, Format(Now, "YYYY/MM/DD HH:MM:SS"))
            
            
            Call UniCode_Conv(Y_SYUREC.FILLER, "")
        
'            SYUKA_QTY = SUMI_JITU_QTY + MI_JITU_QTY
'            SUMI_JITU_QTY = 0
'            MI_JITU_QTY = 0
        
        Case Else
                                    '�o�ח\��Ǎ���
            Call UniCode_Conv(K0_Y_SYU.JGYOBU, JGYOBU)                  '���ƕ�
'            Call UniCode_Conv(K0_Y_SYU.KEY_CYU_KBN, CYU_KBN)            '�����敪2004.04.08
            Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, ID_NO)                'IDNo

            RETRY_CNT = 0
            Do
                sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                Select Case sts
                    Case BtNoErr
                        If CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)) = CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) Then
                            If MESG_FLG = 1 Then
                                Beep
                                MsgBox "���[���Ńf�[�^���ύX����Ă��܂��B�X�V�����𒆎~���܂��B", vbOKOnly, "�m�F����"
                            End If
                            Exit Function
                        End If
                        Exit Do
                    Case BtErrKeyNotFound
                        If MESG_FLG = 1 Then
                            Beep
                            MsgBox "�o�ח\�肪���݂��܂���B�X�V�����𒆎~���܂��B", vbOKOnly, "�m�F����"
                        End If
                        Exit Function
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        '���g���C�񐔃`�F�b�N
                        If RETRY_SU <> 0 Then

                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                        '�񐔃I�[�o�[
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��", 0)
                                Syuko_Update_Proc = SYS_CANCEL
                                Exit Function

                            End If

                        End If

                        If MESG_FLG = 0 Then
'                            DoEvents
                            If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                                DoEvents                                                    '2016.01.26
                            End If                                                          '2016.01.26
                        Else
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                Syuko_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case BtErrDEAD_LOCK
                        Syuko_Update_Proc = SYS_CANCEL
                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��", MESG_FLG)
                        Syuko_Update_Proc = SYS_ERR
                        Exit Function
                End Select
            Loop
    End Select
'============================================================
    If Len(Trim(NYUKA_DT)) = 0 Then
    '*------------------------------------------------------'���ד��w�薳�� �݌Ƀf�[�^�Ǎ��݁i�X�L���i�����j
    '
    
    
    
        '���۸ޏo��    '2008.08.06
        
        wkYOIN = YOIN   '2011.08.12
        
        If LOG_NON = 1 Then
        Else
            If SUMI_JITU_QTY = 0 And MI_JITU_QTY = 0 Then
                
                
                If Left(YOIN, 1) = ACT_DENPYO_ID Or _
                    Left(YOIN, 1) = ACT_DENPYO_ID2 Or _
                    Left(YOIN, 1) = ACT_SYUKA_HYO Then      'ACT_DENPYO_ID2�ǉ��@2015.02.21
                    wkYOIN = ACT_SYUKA_KEI & CYU_KBN
                End If
                
                
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.04
'                If P_SAGYO_LOG_OUTPUT_PROC(TANTO_CODE, _
'                                            ID, _
'                                            JGYOBU, _
'                                            NAIGAI, _
'                                            MENU_NO, _
'                                            wkYOIN, _
'                                            HIN_GAI, _
'                                            SYUKA_QTY, _
'                                            0, _
'                                            FROM_LOCATION, _
'                                            "", _
'                                            ID_NO, _
'                                            MUKE_CODE) Then
'                    Exit Function
'                End If
                Total_SUMI_JITU_QTY = 0
                Total_MI_JITU_QTY = 0
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.04
            
            
            Else
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
                                            "") Then
                    Exit Function
                End If
            End If
        End If
    
    
    
    
    
    '---------------------------------------'���i���ςݏ���
        If SUMI_JITU_QTY <> 0 Then
            Zan_Qty = SUMI_JITU_QTY
            Do
                Call UniCode_Conv(K0_ZAIKO.SOKO_NO, Mid(FROM_LOCATION, 1, 2))   '�q�ɇ�
                Call UniCode_Conv(K0_ZAIKO.Retu, Mid(FROM_LOCATION, 3, 2))      '��
                Call UniCode_Conv(K0_ZAIKO.Ren, Mid(FROM_LOCATION, 5, 2))       '�A
                Call UniCode_Conv(K0_ZAIKO.Dan, Mid(FROM_LOCATION, 7, 2))       '�i
                Call UniCode_Conv(K0_ZAIKO.JGYOBU, JGYOBU)                      '���ƕ�
                Call UniCode_Conv(K0_ZAIKO.NAIGAI, NAIGAI)                      '���O
                Call UniCode_Conv(K0_ZAIKO.HIN_GAI, HIN_GAI)                    '�i�ԁi�O���j
                Call UniCode_Conv(K0_ZAIKO.GOODS_ON, "0")                       '���i�^�����i
                Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, "")                        '���ד�
                
                RETRY_CNT = 0

                Do
                    sts = BTRV(BtOpGetGreater + BtSNoWait, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
                    Select Case sts
                        Case BtNoErr
                                                '�I�{�i�{���i�^�����i�u���[�N
                            If FROM_LOCATION <> (StrConv(ZAIKOREC.SOKO_NO, vbUnicode) & _
                                                StrConv(ZAIKOREC.Retu, vbUnicode) & _
                                                StrConv(ZAIKOREC.Ren, vbUnicode) & _
                                                StrConv(ZAIKOREC.Dan, vbUnicode)) Or _
                                JGYOBU <> StrConv(ZAIKOREC.JGYOBU, vbUnicode) Or _
                                NAIGAI <> StrConv(ZAIKOREC.NAIGAI, vbUnicode) Or _
                                Trim(HIN_GAI) <> Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Or _
                                StrConv(ZAIKOREC.GOODS_ON, vbUnicode) <> "0" Then
                                
                                If MESG_FLG = 1 Then
                                    Beep
                                    MsgBox "�݌ɐ����s�����Ă��܂��B�X�V�����𒆎~���܂��B", vbOKOnly, "�m�F����"
                                End If
                                Exit Function
                            
                            End If


                            If Zan_Qty < CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) Then
                                Upd_com = BtOpUpdate
                                WK_Qty = Zan_Qty
                            Else
                                Upd_com = BtOpDelete
                                 WK_Qty = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                            End If

                            Exit Do
                        Case BtErrEOF

                            If MESG_FLG = 1 Then
                                Beep
                                MsgBox "�݌ɐ����s�����Ă��܂��B�X�V�����𒆎~���܂��B", vbOKOnly, "�m�F����"
                            End If
                            Exit Function

                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        '���g���C�񐔃`�F�b�N
                            If RETRY_SU <> 0 Then

                                RETRY_CNT = RETRY_CNT + 1
                                If RETRY_CNT > RETRY_SU Then
                                        '�񐔃I�[�o�[
                                    Call File_Error(sts, com + BtSNoWait, "�݌Ƀf�[�^", 1)
                                    Syuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If

                            End If

                            If MESG_FLG = 0 Then
'                                DoEvents
                                If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                                    DoEvents                                                    '2016.01.26
                                End If                                                          '2016.01.26
                            Else
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Syuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If
                            End If
                        Case BtErrDEAD_LOCK
                            Syuko_Update_Proc = SYS_CANCEL
                            Exit Function
                        Case Else
                            Call File_Error(sts, com + BtSNoWait, "�݌Ƀf�[�^")
                            Syuko_Update_Proc = SYS_ERR
                            Exit Function
                    End Select
                Loop

                If Upd_com = BtOpUpdate Then
                    Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) - Zan_Qty, "00000000"))
                    Call UniCode_Conv(ZAIKOREC.LOCK_F, LOCK_OFF)                '�r���t���O
                    Call UniCode_Conv(ZAIKOREC.WEL_ID, "")                      '�g�p���q�@ID
                    Call UniCode_Conv(ZAIKOREC.PRG_ID, "")                      '�g�p���v���O����
                End If

                RETRY_CNT = 0
                '*------------------------------------------------------'�݌Ƀf�[�^�o��
                Do
                    sts = BTRV(Upd_com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                            '���g���C�񐔃`�F�b�N
                            If RETRY_SU <> 0 Then

                                RETRY_CNT = RETRY_CNT + 1
                                If RETRY_CNT > RETRY_SU Then
                                            '�񐔃I�[�o�[
                                    Call File_Error(sts, Upd_com, "�݌Ƀf�[�^", 0)
                                    Syuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If
                            End If
                        
                            If MESG_FLG = 0 Then
'                                DoEvents
                                If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                                    DoEvents                                                    '2016.01.26
                                End If                                                          '2016.01.26
                            Else
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Syuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If
                            End If
                        Case BtErrDEAD_LOCK
                            Syuko_Update_Proc = SYS_CANCEL
                            Exit Function
                        Case Else
                            Call File_Error(sts, Upd_com, "�݌Ƀf�[�^")
                            Syuko_Update_Proc = SYS_ERR
                            Exit Function
                    End Select
                Loop
            '============================================================
                '*------------------------------------------------------'�݌Ɉړ����o��
                sts = IDOREKI_OUTPUT_PROC(FROM_LOCATION, _
                                            Space(8), _
                                            JGYOBU, _
                                            NAIGAI, _
                                            HIN_GAI, _
                                            StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), _
                                            YOIN, _
                                            WK_Qty, _
                                            0, _
                                            ID, _
                                            TANTO_CODE, _
                                            RETRY, _
                                            CYU_KBN, _
                                            MEMO, _
                                            Ins_DateTime, _
                                            StrConv(ZAIKOREC.SHIIRE_CODE, vbUnicode), _
                                            StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode), _
                                            StrConv(ZAIKOREC.KEIJYO_YM, vbUnicode), MENU_NO, , , , , , , 1)

                If sts Then
                    Syuko_Update_Proc = sts
                    Exit Function
                End If
                
                Zan_Qty = Zan_Qty - WK_Qty
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.04
                Total_SUMI_JITU_QTY = Total_SUMI_JITU_QTY + WK_Qty
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.04
                
                
                
                If Zan_Qty <= 0 Then
                    Exit Do                     '�������Ƃ��I��
                End If
            Loop
        End If
                    
'************************************************************
    '
    '---------------------------------------'�����i����
        If MI_JITU_QTY <> 0 Then
            Zan_Qty = MI_JITU_QTY
            Do
                Call UniCode_Conv(K0_ZAIKO.SOKO_NO, Mid(FROM_LOCATION, 1, 2))   '�q�ɇ�
                Call UniCode_Conv(K0_ZAIKO.Retu, Mid(FROM_LOCATION, 3, 2))      '��
                Call UniCode_Conv(K0_ZAIKO.Ren, Mid(FROM_LOCATION, 5, 2))       '�A
                Call UniCode_Conv(K0_ZAIKO.Dan, Mid(FROM_LOCATION, 7, 2))       '�i
                Call UniCode_Conv(K0_ZAIKO.JGYOBU, JGYOBU)                      '���ƕ�
                Call UniCode_Conv(K0_ZAIKO.NAIGAI, NAIGAI)                      '���O
                Call UniCode_Conv(K0_ZAIKO.HIN_GAI, HIN_GAI)                    '�i�ԁi�O���j
                Call UniCode_Conv(K0_ZAIKO.GOODS_ON, "1")                       '���i�^�����i
                Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, "")                        '���ד�
                
                RETRY_CNT = 0

                Do
                    sts = BTRV(BtOpGetGreater + BtSNoWait, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
                    Select Case sts
                        Case BtNoErr
                                                '�I�{�i�{���i�^�����i�u���[�N
                            If FROM_LOCATION <> (StrConv(ZAIKOREC.SOKO_NO, vbUnicode) & _
                                                StrConv(ZAIKOREC.Retu, vbUnicode) & _
                                                StrConv(ZAIKOREC.Ren, vbUnicode) & _
                                                StrConv(ZAIKOREC.Dan, vbUnicode)) Or _
                                JGYOBU <> StrConv(ZAIKOREC.JGYOBU, vbUnicode) Or _
                                NAIGAI <> StrConv(ZAIKOREC.NAIGAI, vbUnicode) Or _
                                Trim(HIN_GAI) <> Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Or _
                                StrConv(ZAIKOREC.GOODS_ON, vbUnicode) <> "1" Then
                                
                                If MESG_FLG = 1 Then
                                    Beep
                                    MsgBox "�݌ɐ����s�����Ă��܂��B�X�V�����𒆎~���܂��B", vbOKOnly, "�m�F����"
                                End If
                                Exit Function
                            
                            End If


                            If Zan_Qty < CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) Then
                                Upd_com = BtOpUpdate
                                WK_Qty = Zan_Qty
                            Else
                                Upd_com = BtOpDelete
                                WK_Qty = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                            End If

                            Exit Do
                        Case BtErrEOF

                            If MESG_FLG = 1 Then
                                Beep
                                MsgBox "�݌ɐ����s�����Ă��܂��B�X�V�����𒆎~���܂��B", vbOKOnly, "�m�F����"
                            End If
                            Exit Function

                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        '���g���C�񐔃`�F�b�N
                            If RETRY_SU <> 0 Then

                                RETRY_CNT = RETRY_CNT + 1
                                If RETRY_CNT > RETRY_SU Then
                                        '�񐔃I�[�o�[
                                    Call File_Error(sts, com + BtSNoWait, "�݌Ƀf�[�^", 1)
                                    Syuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If

                            End If

                            If MESG_FLG = 0 Then
'                                DoEvents
                                If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                                    DoEvents                                                    '2016.01.26
                                End If                                                          '2016.01.26
                            Else
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Syuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If
                            End If
                        Case BtErrDEAD_LOCK
                            Syuko_Update_Proc = SYS_CANCEL
                            Exit Function
                        Case Else
                            Call File_Error(sts, com + BtSNoWait, "�݌Ƀf�[�^")
                            Syuko_Update_Proc = SYS_ERR
                            Exit Function
                    End Select
                Loop

                If Upd_com = BtOpUpdate Then
                    Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) - Zan_Qty, "00000000"))
                    Call UniCode_Conv(ZAIKOREC.LOCK_F, LOCK_OFF)                '�r���t���O
                    Call UniCode_Conv(ZAIKOREC.WEL_ID, "")                      '�g�p���q�@ID
                    Call UniCode_Conv(ZAIKOREC.PRG_ID, "")                      '�g�p���v���O����
                End If

                RETRY_CNT = 0
                '*------------------------------------------------------'�݌Ƀf�[�^�o��
                Do
                    sts = BTRV(Upd_com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                            '���g���C�񐔃`�F�b�N
                            If RETRY_SU <> 0 Then

                                RETRY_CNT = RETRY_CNT + 1
                                If RETRY_CNT > RETRY_SU Then
                                            '�񐔃I�[�o�[
                                    Call File_Error(sts, Upd_com, "�݌Ƀf�[�^", 0)
                                    Syuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If
                            End If
                        
                            If MESG_FLG = 0 Then
'                                DoEvents
                                If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                                    DoEvents                                                    '2016.01.26
                                End If                                                          '2016.01.26
                            Else
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Syuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If
                            End If
                        Case BtErrDEAD_LOCK
                            Syuko_Update_Proc = SYS_CANCEL
                            Exit Function
                        Case Else
                            Call File_Error(sts, Upd_com, "�݌Ƀf�[�^")
                            Syuko_Update_Proc = SYS_ERR
                            Exit Function
                    End Select
                Loop
            '============================================================
                '*------------------------------------------------------'�݌Ɉړ����o��
                sts = IDOREKI_OUTPUT_PROC(FROM_LOCATION, _
                                            Space(8), _
                                            JGYOBU, _
                                            NAIGAI, _
                                            HIN_GAI, _
                                            StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), _
                                            YOIN, _
                                            0, _
                                            WK_Qty, _
                                            ID, _
                                            TANTO_CODE, _
                                            RETRY, _
                                            CYU_KBN, _
                                            MEMO, _
                                            Ins_DateTime, _
                                            StrConv(ZAIKOREC.SHIIRE_CODE, vbUnicode), _
                                            StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode), _
                                            StrConv(ZAIKOREC.KEIJYO_YM, vbUnicode), MENU_NO, , , , , , , 1)
                If sts Then
                    Syuko_Update_Proc = sts
                    Exit Function
                End If
                
                Zan_Qty = Zan_Qty - WK_Qty
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.04
                Total_MI_JITU_QTY = Total_MI_JITU_QTY + WK_Qty
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.04

                If Zan_Qty <= 0 Then
                    Exit Do                     '�������Ƃ��I��
                End If
            Loop
                    
        End If
'************************************************************
    '
    '---------------------------------------'���i���ς݁`�����i�������������Ă�
        If SYUKA_QTY <> 0 Then
    
            Zan_Qty = SYUKA_QTY
            Do
                Call UniCode_Conv(K0_ZAIKO.SOKO_NO, Mid(FROM_LOCATION, 1, 2))   '�q�ɇ�
                Call UniCode_Conv(K0_ZAIKO.Retu, Mid(FROM_LOCATION, 3, 2))      '��
                Call UniCode_Conv(K0_ZAIKO.Ren, Mid(FROM_LOCATION, 5, 2))       '�A
                Call UniCode_Conv(K0_ZAIKO.Dan, Mid(FROM_LOCATION, 7, 2))       '�i
                Call UniCode_Conv(K0_ZAIKO.JGYOBU, JGYOBU)                      '���ƕ�
                Call UniCode_Conv(K0_ZAIKO.NAIGAI, NAIGAI)                      '���O
                Call UniCode_Conv(K0_ZAIKO.HIN_GAI, HIN_GAI)                    '�i�ԁi�O���j
                Call UniCode_Conv(K0_ZAIKO.GOODS_ON, "")                        '���i�^�����i
                Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, "")                        '���ד�
                
                
                SUMI_JITU_QTY = 0
                MI_JITU_QTY = 0
                
                RETRY_CNT = 0

                Do
                    sts = BTRV(BtOpGetGreater + BtSNoWait, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
                    Select Case sts
                        Case BtNoErr
                                                '�I�{�i�{���i�^�����i�u���[�N
                            If FROM_LOCATION <> (StrConv(ZAIKOREC.SOKO_NO, vbUnicode) & _
                                                StrConv(ZAIKOREC.Retu, vbUnicode) & _
                                                StrConv(ZAIKOREC.Ren, vbUnicode) & _
                                                StrConv(ZAIKOREC.Dan, vbUnicode)) Or _
                                JGYOBU <> StrConv(ZAIKOREC.JGYOBU, vbUnicode) Or _
                                NAIGAI <> StrConv(ZAIKOREC.NAIGAI, vbUnicode) Or _
                                Trim(HIN_GAI) <> Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Then
                                
                                If MESG_FLG = 1 Then
                                    Beep
                                    MsgBox "�݌ɐ����s�����Ă��܂��B�X�V�����𒆎~���܂��B", vbOKOnly, "�m�F����"
                                End If
                                Exit Function
                            
                            End If


                            If Zan_Qty < CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) Then
                                Upd_com = BtOpUpdate
                                WK_Qty = Zan_Qty
                            Else
                                Upd_com = BtOpDelete
                                 WK_Qty = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                            End If

                            Exit Do
                        Case BtErrEOF

                            If MESG_FLG = 1 Then
                                Beep
                                MsgBox "�݌ɐ����s�����Ă��܂��B�X�V�����𒆎~���܂��B", vbOKOnly, "�m�F����"
                            End If
                            Exit Function

                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        '���g���C�񐔃`�F�b�N
                            If RETRY_SU <> 0 Then

                                RETRY_CNT = RETRY_CNT + 1
                                If RETRY_CNT > RETRY_SU Then
                                        '�񐔃I�[�o�[
                                    Call File_Error(sts, com + BtSNoWait, "�݌Ƀf�[�^", 1)
                                    Syuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If

                            End If

                            If MESG_FLG = 0 Then
'                                DoEvents
                                If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                                    DoEvents                                                    '2016.01.26
                                End If                                                          '2016.01.26
                            Else
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Syuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If
                            End If
                        Case BtErrDEAD_LOCK
                            Syuko_Update_Proc = SYS_CANCEL
                            Exit Function
                        Case Else
                            Call File_Error(sts, com + BtSNoWait, "�݌Ƀf�[�^")
                            Syuko_Update_Proc = SYS_ERR
                            Exit Function
                    End Select
                Loop

                If Upd_com = BtOpUpdate Then
                    Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) - Zan_Qty, "00000000"))
                    Call UniCode_Conv(ZAIKOREC.LOCK_F, LOCK_OFF)                '�r���t���O
                    Call UniCode_Conv(ZAIKOREC.WEL_ID, "")                      '�g�p���q�@ID
                    Call UniCode_Conv(ZAIKOREC.PRG_ID, "")                      '�g�p���v���O����
                End If

                If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = GOODS_ON Then
                    SUMI_JITU_QTY = SUMI_JITU_QTY + WK_Qty
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.04
                    Total_SUMI_JITU_QTY = Total_SUMI_JITU_QTY + WK_Qty
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.04
                Else
                    MI_JITU_QTY = MI_JITU_QTY + WK_Qty
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.04
                    Total_MI_JITU_QTY = Total_MI_JITU_QTY + WK_Qty
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.04
                End If

                RETRY_CNT = 0
                '*------------------------------------------------------'�݌Ƀf�[�^�o��
                Do
                    sts = BTRV(Upd_com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                            '���g���C�񐔃`�F�b�N
                            If RETRY_SU <> 0 Then

                                RETRY_CNT = RETRY_CNT + 1
                                If RETRY_CNT > RETRY_SU Then
                                            '�񐔃I�[�o�[
                                    Call File_Error(sts, Upd_com, "�݌Ƀf�[�^", 0)
                                    Syuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If
                            End If
                        
                            If MESG_FLG = 0 Then
'                                DoEvents
                                If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                                    DoEvents                                                    '2016.01.26
                                End If                                                          '2016.01.26
                            Else
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Syuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If
                            End If
                        Case BtErrDEAD_LOCK
                            Syuko_Update_Proc = SYS_CANCEL
                            Exit Function
                        Case Else
                            Call File_Error(sts, Upd_com, "�݌Ƀf�[�^")
                            Syuko_Update_Proc = SYS_ERR
                            Exit Function
                    End Select
                Loop
            '============================================================
                '*------------------------------------------------------'�݌Ɉړ����o��
                sts = IDOREKI_OUTPUT_PROC(FROM_LOCATION, _
                                            Space(8), _
                                            JGYOBU, _
                                            NAIGAI, _
                                            HIN_GAI, _
                                            StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), _
                                            YOIN, _
                                            SUMI_JITU_QTY, _
                                            MI_JITU_QTY, _
                                            ID, _
                                            TANTO_CODE, _
                                            RETRY, _
                                            CYU_KBN, _
                                            MEMO, _
                                            Ins_DateTime, _
                                            StrConv(ZAIKOREC.SHIIRE_CODE, vbUnicode), _
                                            StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode), _
                                            StrConv(ZAIKOREC.KEIJYO_YM, vbUnicode), MENU_NO, , , , , , , 1)

                If sts Then
                    Syuko_Update_Proc = sts
                    Exit Function
                End If
                
                Zan_Qty = Zan_Qty - WK_Qty

                If Zan_Qty <= 0 Then
                    Exit Do                     '�������Ƃ��I��
                End If
            Loop
        End If



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.04
        If LOG_NON = 1 Then
        Else
            If SYUKA_QTY <> 0 Then                              '2011.08.18
                
                If P_SAGYO_LOG_OUTPUT_PROC(TANTO_CODE, _
                                            ID, _
                                            JGYOBU, _
                                            NAIGAI, _
                                            MENU_NO, _
                                            wkYOIN, _
                                            HIN_GAI, _
                                            Total_SUMI_JITU_QTY, _
                                            Total_MI_JITU_QTY, _
                                            FROM_LOCATION, _
                                            "", _
                                            ID_NO, _
                                            MUKE_CODE) Then
                    Exit Function
                End If
            End If                                              '2011.08.18
        End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.04



    Else
    '*------------------------------------------------------'���ד��w��L�� �݌Ƀf�[�^�Ǎ��݁i��ʏ����j
    '
    '---------------------------------------'���i���ςݏ���
        If SUMI_JITU_QTY <> 0 Then
        
            Call UniCode_Conv(K0_ZAIKO.SOKO_NO, Mid(FROM_LOCATION, 1, 2))   '�q�ɇ�
            Call UniCode_Conv(K0_ZAIKO.Retu, Mid(FROM_LOCATION, 3, 2))      '��
            Call UniCode_Conv(K0_ZAIKO.Ren, Mid(FROM_LOCATION, 5, 2))       '�A
            Call UniCode_Conv(K0_ZAIKO.Dan, Mid(FROM_LOCATION, 7, 2))       '�i
            Call UniCode_Conv(K0_ZAIKO.JGYOBU, JGYOBU)                      '���ƕ�
            Call UniCode_Conv(K0_ZAIKO.NAIGAI, NAIGAI)                      '���O
            Call UniCode_Conv(K0_ZAIKO.HIN_GAI, HIN_GAI)                    '�i�ԁi�O���j
            Call UniCode_Conv(K0_ZAIKO.GOODS_ON, "0")                       '���i�^�����i
            Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, NYUKA_DT)                  '���ד�
                                                                    
            RETRY_CNT = 0
            Do
                sts = BTRV(BtOpGetEqual + BtSNoWait, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
                Select Case sts
                    Case BtNoErr
                        If SUMI_JITU_QTY > CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) Then
                            If MESG_FLG = 1 Then
                                Beep
                                MsgBox "�݌ɐ����s�����Ă��܂��B�X�V�����𒆎~���܂��B", vbOKOnly, "�m�F����"
                            End If
                            Exit Function
                        Else
                            If SUMI_JITU_QTY = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) Then
                                Upd_com = BtOpDelete
                            Else
                                Upd_com = BtOpUpdate
                            End If
                        End If
                    
                        Exit Do
                    Case BtErrKeyNotFound
                        If MESG_FLG = 1 Then
                            Beep
                            MsgBox "�݌Ƀf�[�^�����݂��܂���B�X�V�����𒆎~���܂��B", vbOKOnly, "�m�F����"
                        End If
                        Exit Function
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        '���g���C�񐔃`�F�b�N
                        If RETRY_SU <> 0 Then
                    
                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                        '�񐔃I�[�o�[
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�݌Ƀf�[�^", 0)
                                Syuko_Update_Proc = SYS_CANCEL
                                Exit Function
                    
                           End If
                
                        End If
                
                        If MESG_FLG = 0 Then
'                            DoEvents
                            If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                                DoEvents                                                    '2016.01.26
                            End If                                                          '2016.01.26
                        Else
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                Syuko_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case BtErrDEAD_LOCK
                        Syuko_Update_Proc = SYS_CANCEL
                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�݌Ƀf�[�^")
                        Syuko_Update_Proc = SYS_ERR
                        Exit Function
                End Select
            
            Loop
                                        '�݌ɐ�
            If Upd_com = BtOpUpdate Then
                Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) - SUMI_JITU_QTY, "00000000"))
                Call UniCode_Conv(ZAIKOREC.LOCK_F, LOCK_OFF)            '�r���t���O
                Call UniCode_Conv(ZAIKOREC.WEL_ID, "")                  '�g�p���q�@ID
                Call UniCode_Conv(ZAIKOREC.PRG_ID, "")                  '�g�p���v���O����
            End If
        
            RETRY_CNT = 0
    '*------------------------------------------------------'�݌Ƀf�[�^�o��
            Do
                sts = BTRV(Upd_com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        '���g���C�񐔃`�F�b�N
                        If RETRY_SU <> 0 Then
                    
                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                        '�񐔃I�[�o�[
                                Call File_Error(sts, Upd_com, "�݌Ƀf�[�^", 0)
                                Syuko_Update_Proc = SYS_CANCEL
                                Exit Function
                    
                            End If
                    
                        End If
                
                        If MESG_FLG = 0 Then
'                            DoEvents
                            If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                                DoEvents                                                    '2016.01.26
                            End If                                                          '2016.01.26
                        Else
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                Syuko_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case BtErrDEAD_LOCK
                        Syuko_Update_Proc = SYS_CANCEL
                        Exit Function
                    Case Else
                        Call File_Error(sts, Upd_com, "�݌Ƀf�[�^")
                        Syuko_Update_Proc = SYS_ERR
                        Exit Function
                        
                End Select
            Loop
'============================================================
    '*------------------------------------------------------'�݌Ɉړ����o��
            sts = IDOREKI_OUTPUT_PROC(FROM_LOCATION, _
                                        Space(8), _
                                        JGYOBU, _
                                        NAIGAI, _
                                        HIN_GAI, _
                                        NYUKA_DT, _
                                        YOIN, _
                                        SUMI_JITU_QTY, _
                                        0, _
                                        ID, _
                                        TANTO_CODE, _
                                        RETRY, _
                                        CYU_KBN, _
                                        MEMO, _
                                        Ins_DateTime, _
                                        StrConv(ZAIKOREC.SHIIRE_CODE, vbUnicode), _
                                        StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode), _
                                        StrConv(ZAIKOREC.KEIJYO_YM, vbUnicode), MENU_NO)
            If sts Then
                Syuko_Update_Proc = sts
                Exit Function
            End If
        End If
'************************************************************
    '
    '---------------------------------------'���i���ςݏ���
    
        If MI_JITU_QTY <> 0 Then
        
            Call UniCode_Conv(K0_ZAIKO.SOKO_NO, Mid(FROM_LOCATION, 1, 2))   '�q�ɇ�
            Call UniCode_Conv(K0_ZAIKO.Retu, Mid(FROM_LOCATION, 3, 2))      '��
            Call UniCode_Conv(K0_ZAIKO.Ren, Mid(FROM_LOCATION, 5, 2))       '�A
            Call UniCode_Conv(K0_ZAIKO.Dan, Mid(FROM_LOCATION, 7, 2))       '�i
            Call UniCode_Conv(K0_ZAIKO.JGYOBU, JGYOBU)                      '���ƕ�
            Call UniCode_Conv(K0_ZAIKO.NAIGAI, NAIGAI)                      '���O
            Call UniCode_Conv(K0_ZAIKO.HIN_GAI, HIN_GAI)                    '�i�ԁi�O���j
            Call UniCode_Conv(K0_ZAIKO.GOODS_ON, "1")                       '���i�^�����i
            Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, NYUKA_DT)                  '���ד�
                                                                    
            RETRY_CNT = 0
            Do
                sts = BTRV(BtOpGetEqual + BtSNoWait, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
                Select Case sts
                    Case BtNoErr
                        If SUMI_JITU_QTY > CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) Then
                            If MESG_FLG = 1 Then
                                Beep
                                MsgBox "�݌ɐ����s�����Ă��܂��B�X�V�����𒆎~���܂��B", vbOKOnly, "�m�F����"
                            End If
                            Exit Function
                        Else
                            If MI_JITU_QTY = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) Then
                                Upd_com = BtOpDelete
                            Else
                                Upd_com = BtOpUpdate
                            End If
                        End If
                    
                        Exit Do
                    Case BtErrKeyNotFound
                        If MESG_FLG = 1 Then
                            Beep
                            MsgBox "�݌Ƀf�[�^�����݂��܂���B�X�V�����𒆎~���܂��B", vbOKOnly, "�m�F����"
                        End If
                        Exit Function
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        '���g���C�񐔃`�F�b�N
                        If RETRY_SU <> 0 Then
                    
                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                        '�񐔃I�[�o�[
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�݌Ƀf�[�^", 0)
                                Syuko_Update_Proc = SYS_CANCEL
                                Exit Function
                    
                           End If
                
                        End If
                
                        If MESG_FLG = 0 Then
'                            DoEvents
                            If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                                DoEvents                                                    '2016.01.26
                            End If                                                          '2016.01.26
                        Else
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                Syuko_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case BtErrDEAD_LOCK
                        Syuko_Update_Proc = SYS_CANCEL
                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�݌Ƀf�[�^")
                        Syuko_Update_Proc = SYS_ERR
                        Exit Function
                End Select
            
            Loop
                                        '�݌ɐ�
            If Upd_com = BtOpUpdate Then
                Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) - MI_JITU_QTY, "00000000"))
                Call UniCode_Conv(ZAIKOREC.LOCK_F, LOCK_OFF)            '�r���t���O
                Call UniCode_Conv(ZAIKOREC.WEL_ID, "")                  '�g�p���q�@ID
                Call UniCode_Conv(ZAIKOREC.PRG_ID, "")                  '�g�p���v���O����
            End If
        
            RETRY_CNT = 0
    '*------------------------------------------------------'�݌Ƀf�[�^�o��
            Do
                sts = BTRV(Upd_com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        '���g���C�񐔃`�F�b�N
                        If RETRY_SU <> 0 Then
                    
                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                        '�񐔃I�[�o�[
                                Call File_Error(sts, Upd_com, "�݌Ƀf�[�^", 0)
                                Syuko_Update_Proc = SYS_CANCEL
                                Exit Function
                    
                            End If
                    
                        End If
                
                        If MESG_FLG = 0 Then
'                            DoEvents
                            If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                                DoEvents                                                    '2016.01.26
                            End If                                                          '2016.01.26
                        Else
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                Syuko_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case BtErrDEAD_LOCK
                        Syuko_Update_Proc = SYS_CANCEL
                        Exit Function
                    Case Else
                        Call File_Error(sts, Upd_com, "�݌Ƀf�[�^")
                        Syuko_Update_Proc = SYS_ERR
                        Exit Function
                        
                End Select
            Loop
'============================================================
    '*------------------------------------------------------'�݌Ɉړ����o��
            sts = IDOREKI_OUTPUT_PROC(FROM_LOCATION, _
                                        Space(8), _
                                        JGYOBU, _
                                        NAIGAI, _
                                        HIN_GAI, _
                                        NYUKA_DT, _
                                        YOIN, _
                                        0, _
                                        MI_JITU_QTY, _
                                        ID, _
                                        TANTO_CODE, _
                                        RETRY, _
                                        CYU_KBN, _
                                        MEMO, _
                                        Ins_DateTime, _
                                        StrConv(ZAIKOREC.SHIIRE_CODE, vbUnicode), _
                                        StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode), _
                                        StrConv(ZAIKOREC.KEIJYO_YM, vbUnicode), MENU_NO)
            If sts Then
                Syuko_Update_Proc = sts
                Exit Function
            End If
    
        End If
    End If
'============================================================
    If CYU_KBN = " " Then
    Else
        
        Wk_SYUKA_QTY = Wk_SYUKA_QTY + Wk_SUMI_JITU_QTY + Wk_MI_JITU_QTY
        If CYU_KBN <> CYU_KBN_KIN Then
        
        '*------------------------------------------------------'�o�ח\��X�V
            If (CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)) - CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode))) = Wk_SYUKA_QTY Then
                Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_KBN_FIN)
                                                                '�������t
                Call UniCode_Conv(Y_SYUREC.KAN_YMD, Format(Date, "yyyymmdd"))
                                                                
                '2011.03.30 �X�V���t�̏������ݒǉ�
                Call UniCode_Conv(Y_SYUREC.KAN_HMS, Format(Now, "HHMMSS"))
                                                                
                                                                '�m�萔��
                Call UniCode_Conv(Y_SYUREC.JITU_SURYO, StrConv(Y_SYUREC.SURYO, vbUnicode))
            Else
                                                                '�m�萔��
                Call UniCode_Conv(Y_SYUREC.JITU_SURYO, Format(CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) + Wk_SYUKA_QTY, "0000000"))
            End If
            
            Call UniCode_Conv(Y_SYUREC.WEL_ID, "")              '�g�p�[��ID�i���󔒁j
            Call UniCode_Conv(Y_SYUREC.PRG_ID, "")              '�g�p���v���O�����i���󔒁j
            com = BtOpUpdate
        Else
        
            Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_KBN_FIN)
                                                                '�������t
            Call UniCode_Conv(Y_SYUREC.KAN_YMD, Format(Date, "yyyymmdd"))
        
            
            '2011.03.30 �X�V���t�̏������ݒǉ�
            Call UniCode_Conv(Y_SYUREC.KAN_HMS, Format(Now, "HHMMSS"))
            
            Call UniCode_Conv(Y_SYUREC.JITU_SURYO, Format(Wk_SYUKA_QTY, "0000000"))
            Call UniCode_Conv(Y_SYUREC.WEL_ID, "")              '�g�p�[��ID�i���󔒁j
            Call UniCode_Conv(Y_SYUREC.PRG_ID, "")              '�g�p���v���O�����i���󔒁j
            
            com = BtOpInsert
        End If


        '*------------------------------------------------------'�o�ח\��o��
        RETRY_CNT = 0
        Do
            sts = BTRV(com, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                    '���g���C�񐔃`�F�b�N
                    If RETRY_SU <> 0 Then

                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > RETRY_SU Then
                                '�񐔃I�[�o�[
                            Call File_Error(sts, com, "�o�ח\��", 0)
                            Syuko_Update_Proc = SYS_CANCEL
                            Exit Function

                        End If

                    End If

                    If MESG_FLG = 0 Then
'                        DoEvents
                        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                            DoEvents                                                    '2016.01.26
                        End If                                                          '2016.01.26
                    Else
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            Syuko_Update_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                Case BtErrDuplicates
                    If com = BtOpUpdate Then            '�X�V���ُ͈�
                        Call File_Error(sts, com, "�o�ח\��", MESG_FLG)
                        Syuko_Update_Proc = SYS_ERR
                        Exit Function
                    End If
                    If Len(Trim(ID_NO)) <> 0 Then
                        Call File_Error(sts, com, "�o�ח\��", MESG_FLG)
                        Syuko_Update_Proc = SYS_ERR
                        Exit Function
                    Else
                                                        'ID���Ď捞�݂�LOOP
                        sts = Den_No_Set_Proc(21, JGYOBU, GET_ID_NO)
                        If sts Then
                            Syuko_Update_Proc = sts
                            Exit Function
                        End If
                                                        'ID��
                        Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, GET_ID_NO)
                        Call UniCode_Conv(Y_SYUREC.ID_NO, GET_ID_NO)
'                        DoEvents
                        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                            DoEvents                                                    '2016.01.26
                        End If                                                          '2016.01.26
                    End If
                Case BtErrDEAD_LOCK
                    Syuko_Update_Proc = SYS_CANCEL
                    Exit Function
                Case Else
                    Call File_Error(sts, com, "�o�ח\��", MESG_FLG)
                    Syuko_Update_Proc = SYS_ERR
                    Exit Function

            End Select
        Loop
    End If
'============================================================
                                        '�ŏI�o�ɓ�
    Call UniCode_Conv(ITEMREC.LAST_SYU_DT, Format(Date, "yyyymmdd"))
    
    Call UniCode_Conv(ITEMREC.G_LAST_SYUKA_QTY, Format(SUMI_JITU_QTY + MI_JITU_QTY, "00000000"))
    
    '*------------------------------------------------------'�i�ڃ}�X�^�o��
    RETRY_CNT = 0
    Do
        sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                    '���g���C�񐔃`�F�b�N
                If RETRY_SU <> 0 Then
                        
                    RETRY_CNT = RETRY_CNT + 1
                    If RETRY_CNT > RETRY_SU Then
                                '�񐔃I�[�o�[
                        Call File_Error(sts, BtOpUpdate, "�i�ڃ}�X�^", 0)
                        Syuko_Update_Proc = SYS_CANCEL
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
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Syuko_Update_Proc = SYS_CANCEL
                        Exit Function
                    End If
                End If
            Case BtErrDEAD_LOCK
                Syuko_Update_Proc = SYS_CANCEL
                Exit Function
            Case Else
                Call File_Error(sts, BtOpUpdate, "�i�ڃ}�X�^")
                Syuko_Update_Proc = SYS_ERR
                Exit Function
                        
        End Select
    Loop
'============================================================
    
    Syuko_Update_Proc = False
    
End Function

Public Function wY_SYU_Open(Mode As Integer) As Integer
'****************************************************
'*      �u�o�ׁ^�o�ɏ����v    �o�ח\��n�o�d�m����
'*
'*  �o�ח\��t�@�C����ʃ|�C���^�łn�o�d�m����
'*  (�Ăь��ŋN�����ɂP�x�����Ăяo��)

'*  �߂�l: false       :����
'*          true        :�ُ�
'*          SYS_CANCEL  :�X�V��ݾ�
'****************************************************
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String

Dim ans         As Integer
    
    
    wY_SYU_Open = True
                                '�݌Ƀf�[�^�@�t���p�X�捞��
    sts = GetIni("FILE", Y_SYU_ID, "SYS", c)
    
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, wY_SYU_POS, wY_SYUREC, Len(wY_SYUREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
'-------------- �n�o�d�m�����ł̎g�p���́A�����グ���ɂP�񂾂��̂͂��Ȃ̂ŁA��ɉ�ʓ��͂Ƃ��A
'               ��ݾق́A�����̋N����ݾقƂ���B
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    wY_SYU_Open = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpOpen, "�o�ח\��")
                Exit Function
        End Select
    Loop

    wY_SYU_Open = False

End Function

Public Function wY_SYU_CLOSE() As Integer

'****************************************************
'*      �u�o�ׁ^�o�ɏ����v    �o�ח\��b�k�n�r�d����
'*
'*  �o�ח\��t�@�C����ʃ|�C���^�łb�k�n�r�d����
'*  (�Ăь��ŏI�����ɂP�x�����Ăяo��)
'*  �߂�l: false       :����
'*          true        :�ُ�
'****************************************************
Dim sts As Integer
    
    wY_SYU_CLOSE = True
    
    sts = BTRV(BtOpClose, wY_SYU_POS, wY_SYUREC, Len(wY_SYUREC), K3_wY_SYU, Len(K3_wY_SYU), 3)
    
    Select Case sts
        Case BtNoErr, BtErrNoOpen
        Case Else
            Call File_Error(sts, BtOpClose, "�o�ח\��")
            Exit Function
    End Select

    wY_SYU_CLOSE = False

End Function

