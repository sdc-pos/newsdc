VERSION 5.00
Begin VB.Form F1060701 
   BackColor       =   &H00C0C0C0&
   Caption         =   "���i�h�~�x������"
   ClientHeight    =   4710
   ClientLeft      =   2025
   ClientTop       =   2265
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
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   7320
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "���s��"
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
      Left            =   3120
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "���i�h�~�x������"
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
      Left            =   1800
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.Label LabJIGYO 
      Appearance      =   0  '�ׯ�
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   6480
      Width           =   180
   End
End
Attribute VB_Name = "F1060701"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private KeppinArarm_DATA    As String       '���i�A���[���f�[�^�t���p�X
Private Kakeritu            As Integer      '���i�h�~�|����

Private Function OUTPUT_Proc() As Integer
'----------------------------------------------------------------------------
'                  �b�r�u�f�[�^�o�͏���
'----------------------------------------------------------------------------
    
Dim sts             As Integer
Dim com             As Integer
Dim ans             As Integer

Dim FileNo          As Integer
Dim fileName        As String


Dim AVE_SYUKA       As Long

Dim Alarm_Flg       As Boolean

Dim c               As String * 128
Dim Soko_No         As String * 2


    OUTPUT_Proc = True
'���s���̓C�x���g�擾�s��
    Call Input_Lock         '��ʍ��ڃ��b�N

    Label1(0).Visible = True
    Label1(1).Visible = True
    

    '-------------------------------------------    ���i�h�~���O���瑝�����L�����i�ڂ��폜����
    com = BtOpGetFirst
    Do
        DoEvents
        '�݌ɏW�v�f�[�^�ǂݍ���
        sts = BTRV(com, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)

        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�݌ɏW�v�f�[�^")
                Exit Function
        End Select
            
        If CLng(StrConv(SUMZREC.HS_ZAIQTY, vbUnicode)) > CLng(StrConv(SUMZREC.ZEN_HS_ZAIQTY, vbUnicode)) Then
            '�݌ɑ������L�����猎���Ϗo�א��ǂݍ���
            Call UniCode_Conv(K0_AVE_SYUKA.JGYOBU, StrConv(SUMZREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_AVE_SYUKA.NAIGAI, StrConv(SUMZREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_AVE_SYUKA.HIN_GAI, StrConv(SUMZREC.HIN_GAI, vbUnicode))
    
            sts = BTRV(BtOpGetEqual, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
            Select Case sts
                Case BtNoErr
                    AVE_SYUKA = CLng(StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode))
                Case BtErrKeyNotFound
                    AVE_SYUKA = 0
                Case Else
                    Call File_Error(sts, com, "�݌ɏW�v�f�[�^")
                    Exit Function
            End Select
                '�O���݌ɐ��@���@�����Ϗo�א��@*�@����
            If CLng(StrConv(SUMZREC.HS_ZAIQTY, vbUnicode)) > (AVE_SYUKA * (Kakeritu / 100)) Then
                '���i�h�~���O�����������
                Call UniCode_Conv(K0_KEPPINLOG.JGYOBU, StrConv(SUMZREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_KEPPINLOG.NAIGAI, StrConv(SUMZREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_KEPPINLOG.HIN_GAI, StrConv(SUMZREC.HIN_GAI, vbUnicode))
               
                Do
                    sts = BTRV(BtOpGetEqual + BtSNoWait, KEPPINLOG_POS, KEPPINLOGREC, Len(KEPPINLOGREC), K0_KEPPINLOG, Len(K0_KEPPINLOG), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrKeyNotFound
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            '���̏����ł͖{�����肦�Ȃ��I�I
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<KEPPINLOG.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                Exit Function
                            End If
                        Case Else
                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "���i�h�~�x�����O")
                            Exit Function
                    End Select
                Loop
        
                If sts = BtNoErr Then
        
                    Do
                        sts = BTRV(BtOpDelete, KEPPINLOG_POS, KEPPINLOGREC, Len(KEPPINLOGREC), K0_KEPPINLOG, Len(K0_KEPPINLOG), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                '���̏����ł͖{�����肦�Ȃ��I�I
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<KEPPINLOG.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Exit Function
                                End If
                           Case Else
                                Call File_Error(sts, BtOpDelete, "���i�h�~�x�����O")
                                Exit Function
                        End Select
                    Loop
        
                End If
            End If
        End If
    
        com = BtOpGetNext
    
    Loop
    '-------------------------------------------
    
    FileNo = FreeFile
    fileName = KeppinArarm_DATA
    
    
    On Error GoTo Error_Proc
    
    Open (fileName) For Output As FileNo


    Write #FileNo, "���i�h�~�x�����X�g(" & Format(Now, "YYYY/MM/DD") & "�쐬�j"
    Write #FileNo, "���ƕ�", "�����O", "�i�ԁi�O���j", "���_�݌�", "�����Ϗo�א�", "�W���I��"
    
    
    '-------------------------------------------    �O�X������݌ɂ����������̌��i���`�F�b�N����
    Alarm_Flg = False
    com = BtOpGetFirst
    Do
        DoEvents
        '�݌ɏW�v�f�[�^�ǂݍ���
        sts = BTRV(com, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)

        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�݌ɏW�v�f�[�^")
                Exit Function
        End Select
    
        If CLng(StrConv(SUMZREC.HS_ZAIQTY, vbUnicode)) < CLng(StrConv(SUMZREC.ZEN_HS_ZAIQTY, vbUnicode)) Then
        
            '�݌Ɍ������L�����猎���Ϗo�א��ǂݍ���
            Call UniCode_Conv(K0_AVE_SYUKA.JGYOBU, StrConv(SUMZREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_AVE_SYUKA.NAIGAI, StrConv(SUMZREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_AVE_SYUKA.HIN_GAI, StrConv(SUMZREC.HIN_GAI, vbUnicode))
    
            sts = BTRV(BtOpGetEqual, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
            Select Case sts
                Case BtNoErr
                    AVE_SYUKA = CLng(StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode))
                Case BtErrKeyNotFound
                    AVE_SYUKA = 0
                Case Else
                    Call File_Error(sts, com, "�݌ɏW�v�f�[�^")
                    Exit Function
            End Select
        
            If CLng(StrConv(SUMZREC.HS_ZAIQTY, vbUnicode)) < (AVE_SYUKA * (Kakeritu / 100)) Then
                '�݌ɐ������Ȃ��Ȃ����猇�i�h�~���O���`�F�b�N����
                Call UniCode_Conv(K0_KEPPINLOG.JGYOBU, StrConv(SUMZREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_KEPPINLOG.NAIGAI, StrConv(SUMZREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_KEPPINLOG.HIN_GAI, StrConv(SUMZREC.HIN_GAI, vbUnicode))
               
                sts = BTRV(BtOpGetEqual, KEPPINLOG_POS, KEPPINLOGREC, Len(KEPPINLOGREC), K0_KEPPINLOG, Len(K0_KEPPINLOG), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "���i�h�~�x�����O")
                        Exit Function
                End Select
                            
            
            
                If sts = BtErrKeyNotFound Then
                    '���i�h�~���O���o�^�Ȃ�
                    Alarm_Flg = True
                
                                                '���ƕ�
                    Write #FileNo, StrConv(SUMZREC.JGYOBU, vbUnicode),
                                                '�����O
                    Write #FileNo, StrConv(SUMZREC.NAIGAI, vbUnicode),
                                                '�i�ځi�O���j
                    Write #FileNo, StrConv(SUMZREC.HIN_GAI, vbUnicode),
                                                '�O�����_�݌�
                    Write #FileNo, Format(CLng(StrConv(SUMZREC.HS_ZAIQTY, vbUnicode))),
                                                '�O�����_�݌�
                    Write #FileNo, Format(CLng(StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode))),
                                                '�W���I��
                    If GetIni("SOKO_NO", StrConv(SUMZREC.ST_SOKO, vbUnicode), "SYS", c) Then
                        Soko_No = StrConv(SUMZREC.ST_SOKO, vbUnicode)
                    Else
                        Soko_No = Trim(c)
                    End If
                    
                    
                    
                    Write #FileNo, Soko_No & "-" _
                                     & StrConv(SUMZREC.ST_RETU, vbUnicode) & "-" _
                                     & StrConv(SUMZREC.ST_REN, vbUnicode) & "-" _
                                     & StrConv(SUMZREC.ST_DAN, vbUnicode)
                
                    '���i���O�o��
                                    
                    Call UniCode_Conv(KEPPINLOGREC.JGYOBU, StrConv(SUMZREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(KEPPINLOGREC.NAIGAI, StrConv(SUMZREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(KEPPINLOGREC.HIN_GAI, StrConv(SUMZREC.HIN_GAI, vbUnicode))
                    Call UniCode_Conv(KEPPINLOGREC.CREATE_DT, Format(Now, "YYYYMMDD"))
                    Do
                        sts = BTRV(BtOpInsert, KEPPINLOG_POS, KEPPINLOGREC, Len(KEPPINLOGREC), K0_KEPPINLOG, Len(K0_KEPPINLOG), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                '���̏����ł͖{�����肦�Ȃ��I�I
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<KEPPINLOG.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Exit Function
                                End If
                           Case Else
                                Call File_Error(sts, BtOpInsert, "���i�h�~�x�����O")
                                Exit Function
                        End Select
                    Loop

                
                
                
                End If
            
            
            End If
        End If
    
        com = BtOpGetNext
    Loop




    Close #FileNo
    
    Call Input_UnLock         '��ʍ��ڃ��b�N����
    
    If Alarm_Flg Then
        Beep
        MsgBox "���i�h�~�̑Ώەi�ڂ��L��܂����B�u" & fileName & "�v���o�͂���܂����B"
    Else
        Beep
        MsgBox "���i�h�~�̑Ώەi�ڂ͗L��܂���ł����B"
    End If
    
    OUTPUT_Proc = False
    
    Exit Function


Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox fileName & "���g�p���ł��B"
        Call Input_UnLock         '��ʍ��ڃ��b�N����
        OUTPUT_Proc = False
    Else
        MsgBox "Err.Number" & Err.Number
        OUTPUT_Proc = True
    End If
End Function
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    F1060701.MousePointer = vbHourglass

    Call Ctrl_Lock(F1060701)


End Sub
Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1060701)


    F1060701.MousePointer = vbDefault

End Sub

Private Sub Form_Activate()

Dim ans     As Integer


    ans = MsgBox("�u���i�h�~�x�������v���s���܂����H", vbYesNo, "�m�F����")
    
    If ans = vbYes Then
        If OUTPUT_Proc() Then
            Unload Me
        End If
    End If

    Unload Me

End Sub

Private Sub Form_DblClick()
    PrintForm
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   �j���� �c������ �O����
'----------------------------------------------------------------------------
    Select Case KeyCode
        Case vbKeyF1 To vbKeyF12
            Command(KeyCode - vbKeyF1).Value = True
    End Select

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
Dim i   As Integer
Dim c   As String * 128
Dim sts As Integer

    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If
    
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = Trim(c)
                                '���i�h�~�x���f�[�^�t�@�C������荞��
    If GetIni("FILE", "KeppinArarm_DATA", "SYS", c) Then
        Beep
        MsgBox "���i�h�~�x���f�[�^�쐬�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    KeppinArarm_DATA = Trim(c)
                                '���i�h�~�|����
    If GetIni(App.EXEName, "KAKERITU", "SYS", c) Then
        Kakeritu = 100
    Else
        If IsNumeric(Trim(c)) Then
            Kakeritu = CInt(Trim(c))
        Else
            Kakeritu = 100
        End If
    End If
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌ɏW�v�f�[�^�n�o�d�m
    If SUMZ_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�����Ϗo�א��n�o�d�m
    If AVE_SYUKA_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���i�h�~�x�����O�n�o�d�m
    If KEPPINLOG_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    Show
                                
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '�݌ɏW�v�f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌ɏW�v�f�[�^")
        End If
    End If
                                            '�����Ϗo�א��b�k�n�r�d
    sts = BTRV(BtOpClose, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�����Ϗo�א�")
        End If
    End If
                                            '���i�h�~�x�����O�b�k�n�r�d
    sts = BTRV(BtOpClose, KEPPINLOG_POS, KEPPINLOGREC, Len(KEPPINLOGREC), K0_KEPPINLOG, Len(K0_KEPPINLOG), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���i�h�~�x�����O")
        End If
    End If
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1060701 = Nothing

    End
End Sub
