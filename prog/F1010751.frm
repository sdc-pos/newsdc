VERSION 5.00
Begin VB.Form F1010751 
   BackColor       =   &H00C0C0C0&
   Caption         =   "���j���[�Ǘ��}�X�^�Z�b�g�A�b�v"
   ClientHeight    =   4710
   ClientLeft      =   2325
   ClientTop       =   2625
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
   ScaleHeight     =   4710
   ScaleWidth      =   7320
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "���j���[�}�X�^�X�V���I"
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
      Left            =   1440
      TabIndex        =   1
      Top             =   2160
      Visible         =   0   'False
      Width           =   5280
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "���j���[�Ǘ��}�X�^�Z�b�g�A�b�v"
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
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   7200
   End
End
Attribute VB_Name = "F1010751"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private GLB_MENU_NO         As String * 2       '���ʃ��j���[�ԍ�
Private GLB_MENU_NAME       As String           '���ʃ��j���[����


Private NAIGAI_CODE()       As String * 1       '���O�e�[�u��


Private Type YOIN_TBL_Tag                       '�v���e�[�u���i���j���[�̐擪�j
    CODE_TYPE               As String * 1       '��o�[�R�[�h�^�C�v
    CODE_NAME               As String * 5       '��o�[�R�[�h����
End Type

Private YOIN_TBL()          As YOIN_TBL_Tag
Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                  �ŐV���ʃ��j���[�쐬����
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com_MENU    As Integer
Dim com_YOIN    As Integer
Dim com_MTS     As Integer
Dim ans         As Integer

Dim LEVEL_NO1   As Integer
Dim LEVEL_NO2   As Integer
Dim LEVEL_NO3   As Integer
        
Dim i           As Integer
Dim j           As Integer
    
    Update_Proc = True
    
    MsgLab(0).Visible = True
    MsgLab(1).Visible = True
    Me.MousePointer = vbHourglass
    
    '----------------------------   �Ώۃ��j���[�S���폜
    Call UniCode_Conv(K0_MENU.MENU_GRP_NO, GLB_MENU_NO)
    Call UniCode_Conv(K0_MENU.JGYOBU, "")
    Call UniCode_Conv(K0_MENU.NAIGAI, "")
    Call UniCode_Conv(K0_MENU.MENU_LV1, "")
    Call UniCode_Conv(K0_MENU.MENU_LV2, "")
    Call UniCode_Conv(K0_MENU.MENU_LV3, "")
    
    com_MENU = BtOpGetGreaterEqual
    
    Do
        DoEvents
        Do
            sts = BTRV(com_MENU + BtSNoWait, MENU_POS, MENUREC, Len(MENUREC), K0_MENU, Len(K0_MENU), 0)
            Select Case sts
                Case BtNoErr
                    If StrConv(MENUREC.MENU_GRP_NO, vbUnicode) <> GLB_MENU_NO Then
                        sts = BtErrEOF
                    End If
                    Exit Do
                Case BtErrEOF
                    Exit Do
                
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<MENU.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                
                Case Else
                    Call File_Error(sts, com_MENU + BtSNoWait, "���j���[�Ǘ�")
                    Exit Function
            End Select
        
        Loop
        
        If sts = BtErrEOF Then
            Exit Do
        End If
        
        Do
        
            sts = BTRV(BtOpDelete, MENU_POS, MENUREC, Len(MENUREC), K0_MENU, Len(K0_MENU), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<MENU.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                
                Case Else
                    Call File_Error(sts, BtOpDelete, "���j���[�Ǘ�")
                    Exit Function
            End Select
        
        
        Loop
        
        com_MENU = BtOpGetNext
    
    Loop

    '----------------------------   ���j���[�쐬�J�n
    For i = 0 To UBound(JGYOBU_T)                   '���ƕ��̃��[�v
        
        For j = 0 To UBound(NAIGAI_CODE)                '�����O�̃��[�v
            
            For LEVEL_NO1 = 0 To UBound(YOIN_TBL)           '�v���̃��[�v

                Call UniCode_Conv(MENUREC.MENU_GRP_NO, GLB_MENU_NO)
                Call UniCode_Conv(MENUREC.JGYOBU, JGYOBU_T(i).Code)
                Call UniCode_Conv(MENUREC.NAIGAI, NAIGAI_CODE(j))
                Call UniCode_Conv(MENUREC.MENU_LV1, Format(LEVEL_NO1, "000"))
                Call UniCode_Conv(MENUREC.MENU_LV2, "")
                Call UniCode_Conv(MENUREC.MENU_LV3, "")
                
                Call UniCode_Conv(MENUREC.MENU_KBN, "0")
                Call UniCode_Conv(MENUREC.MENU_GRP, GLB_MENU_NAME)
                Call UniCode_Conv(MENUREC.DISPLAY_ITEM, YOIN_TBL(LEVEL_NO1).CODE_NAME)

                Call UniCode_Conv(MENUREC.CODE_TYPE, YOIN_TBL(LEVEL_NO1).CODE_TYPE)
                Call UniCode_Conv(MENUREC.YOIN_CODE, "")
                Call UniCode_Conv(MENUREC.PARAM, "")
                Call UniCode_Conv(MENUREC.FILLER, "")
            
                Do
                    sts = BTRV(BtOpInsert, MENU_POS, MENUREC, Len(MENUREC), K0_MENU, Len(K0_MENU), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<MENU.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                Exit Function
                            End If
                
                        Case Else
                            Call File_Error(sts, BtOpDelete, "���j���[�Ǘ�")
                            Exit Function
                    End Select
                Loop
                '------------------------   �Y���v���}�X�^�r�s�`�q�s
                Call UniCode_Conv(K0_YOIN.CODE_TYPE, YOIN_TBL(LEVEL_NO1).CODE_TYPE)
                Call UniCode_Conv(K0_YOIN.YOIN_CODE, "")
                        
                LEVEL_NO2 = 0
                com_YOIN = BtOpGetGreater
                Do
                    DoEvents
                    sts = BTRV(com_YOIN, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
                    Select Case sts
                        Case BtNoErr
                            If StrConv(YOINREC.CODE_TYPE, vbUnicode) <> YOIN_TBL(LEVEL_NO1).CODE_TYPE Then
                                Exit Do
                            End If
                        Case BtErrEOF
                            Exit Do
                
                        Case Else
                            Call File_Error(sts, com_YOIN, "�v���}�X�^")
                            Exit Function
                    End Select
                
                
                    If StrConv(YOINREC.REGI_F, vbUnicode) = "0" Or StrConv(YOINREC.REGI_F, vbUnicode) = "1" Then
                
                        Call UniCode_Conv(MENUREC.MENU_GRP_NO, GLB_MENU_NO)
                        Call UniCode_Conv(MENUREC.JGYOBU, JGYOBU_T(i).Code)
                        Call UniCode_Conv(MENUREC.NAIGAI, NAIGAI_CODE(j))
                        Call UniCode_Conv(MENUREC.MENU_LV1, Format(LEVEL_NO1, "000"))
                        Call UniCode_Conv(MENUREC.MENU_LV2, Format(LEVEL_NO2, "000"))
                        Call UniCode_Conv(MENUREC.MENU_LV3, "")
                        
                        Call UniCode_Conv(MENUREC.MENU_KBN, "1")
                        Call UniCode_Conv(MENUREC.MENU_GRP, GLB_MENU_NAME)
                        Call UniCode_Conv(MENUREC.DISPLAY_ITEM, StrConv(YOINREC.YOIN_DNAME, vbUnicode))
    
                        Call UniCode_Conv(MENUREC.CODE_TYPE, YOIN_TBL(LEVEL_NO1).CODE_TYPE)
                        Call UniCode_Conv(MENUREC.YOIN_CODE, StrConv(YOINREC.YOIN_CODE, vbUnicode))
                        Call UniCode_Conv(MENUREC.PARAM, "")
                        Call UniCode_Conv(MENUREC.FILLER, "")
                
                        Do
                            sts = BTRV(BtOpInsert, MENU_POS, MENUREC, Len(MENUREC), K0_MENU, Len(K0_MENU), 0)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                                    Beep
                                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<MENU.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                    If ans = vbCancel Then
                                        Exit Function
                                    End If
                    
                                Case Else
                                    Call File_Error(sts, BtOpDelete, "���j���[�Ǘ�")
                                    Exit Function
                            End Select
                        Loop
                    
'                        If StrConv(YOINREC.PARAM_F, vbUnicode) = "1" Then
'
'                        '------------------------------ ������r�s�`�q�s
'                            com_MTS = BtOpGetFirst
'
'                            LEVEL_NO3 = 0
'                            Do
'                                DoEvents
'                                sts = BTRV(com_MTS, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
'                                Select Case sts
'                                    Case BtNoErr
'                                    Case BtErrEOF
'                                        Exit Do
'
'                                    Case Else
'                                        Call File_Error(sts, com_MTS, "������Ǘ��}�X�^")
'                                        Exit Function
'                                End Select
'
'
'                                Call UniCode_Conv(MENUREC.MENU_GRP_NO, GLB_MENU_NO)
'                                Call UniCode_Conv(MENUREC.JGYOBU, JGYOBU_T(i).Code)
'                                Call UniCode_Conv(MENUREC.NAIGAI, NAIGAI_CODE(j))
'                                Call UniCode_Conv(MENUREC.MENU_LV1, Format(LEVEL_NO1, "000"))
'                                Call UniCode_Conv(MENUREC.MENU_LV2, Format(LEVEL_NO2, "000"))
'                                Call UniCode_Conv(MENUREC.MENU_LV3, Format(LEVEL_NO3, "000"))
'
'                                Call UniCode_Conv(MENUREC.MENU_KBN, "1")
'                                Call UniCode_Conv(MENUREC.MENU_GRP, GLB_MENU_NAME)
'                                Call UniCode_Conv(MENUREC.DISPLAY_ITEM, StrConv(MTSREC.MUKE_DNAME, vbUnicode))
'
'                                Call UniCode_Conv(MENUREC.CODE_TYPE, YOIN_TBL(LEVEL_NO1).CODE_TYPE)
'                                Call UniCode_Conv(MENUREC.YOIN_CODE, StrConv(YOINREC.YOIN_CODE, vbUnicode))
'                                Call UniCode_Conv(MENUREC.PARAM, StrConv(MTSREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.SS_CODE, vbUnicode))
'                                Call UniCode_Conv(MENUREC.FILLER, "")
'
'                                Do
'                                    sts = BTRV(BtOpInsert, MENU_POS, MENUREC, Len(MENUREC), K0_MENU, Len(K0_MENU), 0)
'                                    Select Case sts
'                                        Case BtNoErr
'                                            Exit Do
'                                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
'                                            Beep
'                                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<MENU.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'                                            If ans = vbCancel Then
'                                                Exit Function
'                                            End If
'
'                                        Case Else
'                                            Call File_Error(sts, BtOpDelete, "���j���[�Ǘ�")
'                                        Exit Function
'                                    End Select
'                                Loop
'
'
'
'                                LEVEL_NO3 = LEVEL_NO3 + 1
'
'                                com_MTS = BtOpGetNext
'
'
'                            Loop
'
'
'
'                        End If
                    
                    
                        LEVEL_NO2 = LEVEL_NO2 + 1
                    
                    End If
                    
                    com_YOIN = BtOpGetNext
                
                
                
                Loop
            
            
            
            
            Next LEVEL_NO1                                  '�v���̃��[�v

        Next j                                      '�����O�̃��[�v
    
    Next i                                      '���ƕ��̃��[�v

    Update_Proc = False

End Function

Private Sub Form_Activate()

Dim ans As Integer
                                
                                '�����I��
    Beep
    ans = MsgBox("���s���܂����H", vbYesNo + vbQuestion, "�m�F����")
    If ans = vbYes Then
        If Update_Proc() Then
            MsgBox "�ُ�I�����܂����B"
            Unload Me
        End If
    End If
    MsgBox "����I�����܂����B"
    Unload Me

End Sub

Private Sub Form_DblClick()
    PrintForm
End Sub
Private Sub Form_Load()

Dim i           As Integer
Dim j           As Integer

Dim c           As String * 128
Dim sts         As Integer
Dim CODE_TYPE   As String * 1
    
    
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
    LOG_F = RTrim(c)
                                
                                '���ƕ��̊l��
    If JGYOB_TB_Set() Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B"
        End
    End If
                                '�����O�Ǘ��̊l��
    i = 0
    Do
        i = i + 1
        If GetIni(App.EXEName, "NAIGAI" & Format(i, "0"), "SYS", c) Then
            Exit Do
        End If
        ReDim Preserve NAIGAI_CODE(i - 1)
        NAIGAI_CODE(i - 1) = Trim(c)
    
    Loop
    If i = 1 Then
        Beep
        MsgBox "�����O�̊l���Ɏ��s���܂����B"
        End
    End If
                                '���ʃ��j���[�ԍ���荞��
    If GetIni(App.EXEName, "GLB_MENU_NO", "SYS", c) Then
        Beep
        MsgBox "���ʃ��j���[�ԍ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    GLB_MENU_NO = RTrim(c)
                                
                                '���ʃ��j���[���̎�荞��
    If GetIni(App.EXEName, "GLB_MENU_NM", "SYS", c) Then
        Beep
        MsgBox "���ʃ��j���[�ԍ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    GLB_MENU_NAME = RTrim(c)
                                
                                '�v������荞��
    i = -1
    j = 1
    Do
        If GetIni("ACTION", "ACTION_CD" & Format(j, "00"), "SYS", c) Then
            Beep
            MsgBox "�v�����[ACTION_CD]�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
        End If
                
        If Trim(c) = "NON" Then
            Exit Do
        End If
        CODE_TYPE = Trim(c)
           
    
        If GetIni("ACTION", "ACTION_TYPE" & Format(j, "00"), "SYS", c) Then
            Beep
            MsgBox "�v�����[ACTION_TYPE]�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
        End If
    
        If Trim(c) = "1" Then
            '���j���[�o�^�s��
        Else
            '���j���[�o�^��
            
            i = i + 1
            
            ReDim Preserve YOIN_TBL(i)
            YOIN_TBL(i).CODE_TYPE = CODE_TYPE
        
            If GetIni("ACTION", "ACTION_NM" & Format(j, "00"), "SYS", c) Then
                Beep
                MsgBox "�v�����[ACTION_NM]�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
                End
            End If
            YOIN_TBL(i).CODE_NAME = Trim(c)
        End If
    
        j = j + 1
    Loop
                                '������Ǘ��}�X�^�n�o�d�m
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�v���}�X�^�n�o�d�m
    If YOIN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���j���[�Ǘ��}�X�^�n�o�d�m
    If MENU_Open(BtOpenNomal) Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '������Ǘ��}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "������Ǘ��}�X�^")
        End If
    End If
                                            '�v���}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�v���}�X�^")
        End If
    End If
                                            '���j���[�Ǘ��}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, MENU_POS, MENUREC, Len(MENUREC), K0_MENU, Len(K0_MENU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���j���[�Ǘ��}�X�^")
        End If
    End If
    
    sts = BTRV(BtOpReset, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1010751 = Nothing

    End
End Sub

