VERSION 5.00
Begin VB.Form F1100701 
   BackColor       =   &H00C0C0C0&
   Caption         =   "�o�ח\�苭������ "
   ClientHeight    =   4710
   ClientLeft      =   1920
   ClientTop       =   2430
   ClientWidth     =   7320
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MousePointer    =   11  '�����v
   ScaleHeight     =   4710
   ScaleWidth      =   7320
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "�{���ȑO�̏o�ח\��f�[�^"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   5760
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "�S���u�폜�v�X�V���I"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   1
      Left            =   1320
      TabIndex        =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   4800
   End
End
Attribute VB_Name = "F1100701"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Limit_day   As String           '���������P�\����
Private Shori_Mode  As Integer          '���s���[�h 0:�蓮 1:����

Private JYOGAI_MTS  As Variant          '(ini) ���O�q�� �z��
Private OSAKA_PC    As Boolean          '2006.12.06

'2011.07.27
Private SEK_MTS_TBL As Variant          '�ϐ�������
Private SEK_MTS_FLG As Boolean          '�ϐ�������L��
'2011.07.27


Private KENPIN_CHECK    As Integer      '���i������ 2012.11.19


'Private Const LAST_UPDATE_Day$ = "[F110070]2016.07.06 15:30"
Private Const LAST_UPDATE_Day$ = "[F110070]2018.08.30 11:15"




Private Sub Form_Activate()
Dim ans As Integer

    
    
    If Shori_Mode = 0 Then
                        '�蓮���s
        Beep
        ans = MsgBox("�u�o�ח\�苭���폜�i�J�z���X�V�j�v�����@���s���܂����H", vbYesNo + vbDefaultButton2, "�m�F����")
        If ans = vbYes Then
            Call Y_SYU_DEL_PROC
        End If
    
    Else
                        '�������s
        Call Y_SYU_DEL_PROC
    End If

    Unload Me



End Sub

Private Sub Form_DblClick()
    PrintForm
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub
Private Sub Form_Load()
Dim c As String * 128

    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If
                                
                                
   F1100701.Caption = F1100701.Caption & LAST_UPDATE_Day
                                
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)
                                
                                '�o�׃��O�t�@�C������荞��
    If SYUKA_LOGF_GET_PROC() Then
        Beep
        MsgBox "�o�׃��O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
                                
                                '�폜�Ώۓ��t�Z�o
    If GetIni(App.EXEName, "CMPLT_DAY", App.EXEName, c) Then
        Beep
        MsgBox "���t�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
'        Call LOG_OUT(LOG_F, "[SYS.INI] [SYSTEM] [CMPLT_DAY] READ ERROR")           '2016.06.30
        Call LOG_OUT(LOG_F, "[F110070.INI] [F110070] [CMPLT_DAY] READ ERROR")        '2016.06.30
        End
    End If
    
    If Not IsNumeric(Trim(c)) Then
                                '��荞�ُ݈�͍����Ƃ���
        Limit_day = Format(Date, "yyyymmdd")
    Else
        Limit_day = Format(DateAdd("d", -CInt(Trim(c)), Date), "yyyymmdd")
    End If
                                '�������[�h��荞��
    If GetIni(App.EXEName, "AUTO", App.EXEName, c) Then
        Shori_Mode = 0          '��荞�ُ݈�͎蓮����
    Else
        If IsNumeric(Trim(c)) Then
            Shori_Mode = CInt(Trim(c))
        Else
            Shori_Mode = 0      '��荞�ُ݈�͎蓮����
        End If
    End If
                                '��ɏ������ݑΏۂ̂l�s�r ��荞��  2005/06/15
    If GetIni(App.EXEName, "MTS", App.EXEName, c) Then
        c = " "
    End If
    JYOGAI_MTS = Split(Trim(c), ",", -1)


                                '���o�b�H         2006.12.06
    OSAKA_PC = False
    If GetIni(App.EXEName, "OSAKA_PC", App.EXEName, c) Then
    Else
        If Trim(c) = "1" Then
            OSAKA_PC = True
        End If
    End If


'2011.07.27     �ϐ��󂯐�
    
    SEK_MTS_FLG = False
    If GetIni(App.EXEName, "SEK_MTS", App.EXEName, c) Then
        If OSAKA_PC Then                                                                '2016.07.04
            Call LOG_OUT(LOG_F, "[F110070.INI] [F110070] [SEK_MTS] READ ERROR")         '2016.07.04
        End If                                                                          '2016.07.04
    
    
    Else
        SEK_MTS_FLG = True
        SEK_MTS_TBL = Split(Trim(c), ",", -1)
    End If
'2011.07.27

    '-------------------------  �����i������    2012.11.19
    KENPIN_CHECK = 0
    If GetIni(App.EXEName, "KENPIN_CHECK", App.EXEName, c) Then
    Else
        If Trim(c) = "1" Then
            KENPIN_CHECK = 1
        End If
    End If
    '-------------------------  �����i������    2012.11.19

    Show
End Sub
Private Sub Form_Unload(CANCEL As Integer)
    
    Set F1100701 = Nothing
        
    End
End Sub

Private Sub Y_SYU_DEL_PROC()

Dim sts         As Integer
Dim com         As Integer
        
Dim ans         As Integer
        
Dim Undo        As Boolean
Dim i           As Integer
        
        
Dim DEN_NO      As String   '2006.12.06
Dim SEQ_NO      As String   '2006.12.06
        
        
Dim Next_Flg    As Boolean  '2009.02.17
        
        
    MsgLab(0).Visible = True
    MsgLab(1).Visible = True
        
    DoEvents
    If Y_SYU_Open(BtOpenNomal) Then                 '�o�ח\��f�[�^
        Exit Sub
    End If
        
    If DEL_SYU_Open(BtOpenNomal) Then               '�폜�Ϗo�ח\��f�[�^
        Exit Sub
    End If
        
'----------------
    If OSAKA_PC Then
        If Y_SYU_H_Open(BtOpenNomal) Then                 '�o�ח\��(νĲҰ��)�f�[�^
            Exit Sub
        End If
            
        If DEL_SYU_H_Open(BtOpenNomal) Then               '�폜�Ϗo�ח\��(νĲҰ��)�f�[�^
            Exit Sub
        End If
    End If

'----------------
        
        
    com = BtOpGetFirst

    Do
        Do
            DoEvents
            sts = BTRV(com + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
'���̎��_�ł̃t�@�C���g�p���͖������[�v�Ƃ���B�L�����Z���ňُ�I��
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Do
                    End If
                
                    com = BtOpGetEqual
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "�o�ח\��f�[�^")
                    Exit Do
            End Select
        Loop
    
        If sts = BtNoErr Then
            
'            If Limit_day >= StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode) Then            '2018.08.30
            If Limit_day >= StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode) Then         '2018.08.30
                
                
                
                Undo = False
                
                
                                '2009.08.24
                If StrConv(Y_SYUREC.JGYOBU, vbUnicode) = SETSUBI Then
                '-------------------------  �ϐ��Ή�    2011.06.29
                    If KENPIN_CHECK_PROC(Undo) Then
                        Exit Do
                    End If
                '-------------------------  �ϐ��Ή�    2011.06.29
                
                
                '-------------------------  �����i������    2012.11.19
                    If KENPIN_CHECK = 1 Then
                        If Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode)) = "" Then
                            Undo = True
                        End If
                    End If
                '-------------------------  �����i������    2012.11.19
                
                Else
                
                
                    '�f�Ղ�����
                    
                                    
                    If StrConv(Y_SYUREC.CYU_KBN, vbUnicode) = CYU_KBN_BOU Then
                        If CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) <> CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)) Then
                            Undo = True
                        End If
                    Else
                        For i = 0 To UBound(JYOGAI_MTS)
                            If Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)) = JYOGAI_MTS(i) Then
                                Exit For
                            End If
                        Next i
                    
                        If UBound(JYOGAI_MTS) >= 0 Then
                            If i > UBound(JYOGAI_MTS) Then
                                If CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) <> CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)) Then
                                    Undo = True
                                End If
                            End If
                        End If
                    End If
                    
                    
                    If StrConv(Y_SYUREC.JGYOBA, vbUnicode) = "00036003" Then        '2008.01.10
                    
                        '�o�׎����ް��ɑΉ� 2006.08.11
                        If (StrConv(Y_SYUREC.TORI_KBN, vbUnicode) = "25" Or StrConv(Y_SYUREC.TORI_KBN, vbUnicode) = "29") Then
                            If StrConv(Y_SYUREC.CYU_KBN, vbUnicode) <> CYU_KBN_BOU Then
                                
                                If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" Or StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "3" Then
                                
                                    If Trim(StrConv(Y_SYUREC.LK_SEQ_NO, vbUnicode)) = "" Then
                                        Undo = True
                                    End If
                                End If
                            End If
                        End If
                    
                    
                    '2008.02.22
                    Else
                        
                        If CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) <> CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)) Then
                            Undo = True
                        End If
                    
                    
                    End If
                                    
                    '2009.02.09
                    If StrConv(Y_SYUREC.KAN_KBN, vbUnicode) <> KAN_KBN_FIN Then
                        Undo = True
                    End If
                
                
                '-------------------------  �����i������    2012.11.19
                    If KENPIN_CHECK = 1 Then
                        If Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode)) = "" Then
                            Undo = True
                        End If
                    End If
                '-------------------------  �����i������    2012.11.19
                
                
                End If
                                    
                
                
                
                If Undo Then
                Else
                    '���t�����؂�
                    Do
                        DoEvents
                        sts = BTRV(BtOpInsert, DEL_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_DEL_SYU, Len(K0_DEL_SYU), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<DEL_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Exit Do
                                End If
                            Case Else
                                Call File_Error(sts, BtOpInsert, "�폜�Ϗo�ח\��")
                                Exit Do
                        End Select
                    Loop
                    
                    Do
                        DoEvents
                        sts = BTRV(BtOpDelete, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Exit Do
                                End If
                            Case Else
                                Call File_Error(sts, BtOpDelete, "�o�ח\��")
                                Exit Do
                        End Select
                    Loop
                
                
                    '-------------------------------------  νĲҰ���ް��̏���  2006.12.06
                    If OSAKA_PC Then
'                        DEN_NO = Left(Trim(StrConv(Y_SYUREC.KEY_ID_NO, vbUnicode)), Len(Trim(StrConv(Y_SYUREC.KEY_ID_NO, vbUnicode))) - 1)
'                        SEQ_NO = Right(Trim(StrConv(Y_SYUREC.KEY_ID_NO, vbUnicode)), 1)
                    
                    
'                        Call UniCode_Conv(K0_Y_SYU_H.DEN_NO, DEN_NO)
'                        Call UniCode_Conv(K0_Y_SYU_H.SEQ_NO, SEQ_NO)
                    
                    
                    
                        Call UniCode_Conv(K4_Y_SYU_H.ID_NO, StrConv(Y_SYUREC.KEY_ID_NO, vbUnicode))
                    
                        Do
                            DoEvents
'                            sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
                            sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                    Beep
                                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA_H.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                    If ans = vbCancel Then
                                        Exit Do
                                    End If
                                
                                
                                Case BtErrKeyNotFound
                                    
                                    Next_Flg = True
                                    
                                    Exit Do
                                Case Else
                                    Call File_Error(BtOpGetEqual, BtOpGetEqual + BtSNoWait, "�o�ח\��(νĲҰ��)�f�[�^")
                                    Exit Do
                            End Select
                        Loop
                    
                    
                        If sts = BtNoErr Then
                    
                    
                            Do
                                DoEvents
                                sts = BTRV(BtOpInsert, DEL_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_DEL_SYU_H, Len(K0_DEL_SYU_H), 0)
                                Select Case sts
                                    Case BtNoErr
                                        Exit Do
                                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        Beep
                                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<DEL_SYUKA_H.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                        If ans = vbCancel Then
                                            Exit Do
                                        End If
                                    Case Else
                                        Call File_Error(sts, BtOpInsert, "�폜�Ϗo�ח\��(νĲҰ��)")
                                        Exit Do
                                End Select
                            Loop
                            
                            Do
                                DoEvents
                                sts = BTRV(BtOpDelete, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
                                Select Case sts
                                    Case BtNoErr
                                        Exit Do
                                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        Beep
                                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA_H.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                        If ans = vbCancel Then
                                            Exit Do
                                        End If
                                    Case Else
                                        Call File_Error(sts, BtOpDelete, "�o�ח\��(νĲҰ��)")
                                        Exit Do
                                End Select
                            Loop
                                                
                        Else
                            If sts = BtErrKeyNotFound Then
                                sts = BtNoErr
                            End If
                        End If
                    
                    
                    End If
                
                
                
                
                
                    '-------------------------------------  νĲҰ���ް��̏���  2006.12.06
                
                
                
                    If SYUKA_LOG_ON Then
                        
                        If OSAKA_PC Then
                            If Not Next_Flg Then
                        
                                Call SYUKA_LOG_OUT_PROC("DEL", "AFT")
                            End If
                    
                        Else
                            Call SYUKA_LOG_OUT_PROC("DEL", "AFT")
                    
                        End If
                    
                    End If
                End If
            End If
        End If
            
        If sts <> BtNoErr Then
            Exit Do
        End If
            
        com = BtOpGetNext
    
    Loop
        
                                                    '�o�ח\��f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        Call File_Error(sts, BtOpClose, "�o�ח\��f�[�^")
    End If
                                                    '�폜�Ϗo�ח\��f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, DEL_SYU_POS, DEL_SYUREC, Len(DEL_SYUREC), K0_DEL_SYU, Len(K0_DEL_SYU), 0)
    If sts Then
        Call File_Error(sts, BtOpClose, "�폜�Ϗo�ח\��f�[�^")
    End If
        
    
    If OSAKA_PC Then                                '2006.12.06
                                                        '�o�ח\��f�[�^�b�k�n�r�d
        sts = BTRV(BtOpClose, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
        If sts Then
            Call File_Error(sts, BtOpClose, "�o�ח\��(νĲҰ��)�f�[�^")
        End If
                                                        '�폜�Ϗo�ח\��f�[�^�b�k�n�r�d
        sts = BTRV(BtOpClose, DEL_SYU_H_POS, DEL_SYU_HREC, Len(DEL_SYU_HREC), K0_DEL_SYU_H, Len(K0_DEL_SYU_H), 0)
        If sts Then
            Call File_Error(sts, BtOpClose, "�폜�Ϗo�ח\��(νĲҰ��)�f�[�^")
        End If
    End If
    
    sts = BTRV(BtOpReset, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

End Sub


Private Function KENPIN_CHECK_PROC(Undo As Boolean) As Integer
'-----------------------------------------------------------------------------------
'
'   ���o�b�����@���i�ς݁@���@�L�����Z���̃`�F�b�N
'
'               2011.06.29
'               2011.07.27 �ϐ��ȊO�͓��t�ō폜����
'
'-----------------------------------------------------------------------------------
Dim sts As Integer
Dim i   As Integer      '2011.07.27
    
    
    
    KENPIN_CHECK_PROC = True
    
    
    If Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)) = "Z0001" Then
        Debug.Print
    End If
    '------------------------------------------ 2011.07.27
    If SEK_MTS_FLG Then
        For i = 0 To UBound(SEK_MTS_TBL)
            If Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)) = SEK_MTS_TBL(i) Then
                Exit For
            End If
        Next i
        
        If i > UBound(SEK_MTS_TBL) Then     '2016.07.06
            KENPIN_CHECK_PROC = False       '2016.07.06
            Exit Function                   '2016.07.06
        End If                              '2016.07.06
    
    
    
    End If
        
'    If i > UBound(SEK_MTS_TBL) Then        '2016.07.06
'        KENPIN_CHECK_PROC = False          '2016.07.06
'        Exit Function                      '2016.07.06
'    End If                                 '2016.07.06
    
    
    
    
    '------------------------------------------ 2011.07.28
    
    If Trim(StrConv(Y_SYUREC.KENPIN_TANTO_CODE, vbUnicode)) = "" Then
        
        Call UniCode_Conv(K4_Y_SYU_H.ID_NO, StrConv(Y_SYUREC.ID_NO, vbUnicode))
        
        sts = BTRV(BtOpGetEqual, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Call UniCode_Conv(Y_SYU_HREC.CANCEL_F, "1")
            Case Else
                Call File_Error(BtOpGetEqual, BtOpGetEqual, "�o�ח\��(νĲҰ��)�f�[�^")
                Exit Function
        End Select
        
        
        If StrConv(Y_SYU_HREC.CANCEL_F, vbUnicode) = "1" Then
            Debug.Print
        Else
            Undo = True
        End If
    End If
    
    KENPIN_CHECK_PROC = False

End Function
