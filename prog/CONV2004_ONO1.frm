VERSION 5.00
Begin VB.Form CONV2004_ONO1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "�f�[�^�R���o�[�g�����i���с˃A�C�����j"
   ClientHeight    =   7230
   ClientLeft      =   2325
   ClientTop       =   2625
   ClientWidth     =   10095
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
   ScaleHeight     =   7230
   ScaleWidth      =   10095
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Label OUT_Cnt 
      Alignment       =   2  '��������
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   4560
      TabIndex        =   13
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label OUT_Cnt 
      Alignment       =   2  '��������
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   4560
      TabIndex        =   12
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label OUT_Cnt 
      Alignment       =   2  '��������
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   4560
      TabIndex        =   11
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label OUT_Cnt 
      Alignment       =   2  '��������
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   4560
      TabIndex        =   10
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label IN_Cnt 
      Alignment       =   2  '��������
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   3000
      TabIndex        =   9
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�@�o�ח\�聁"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   8
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�݌Ɉړ�����"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   7
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label IN_Cnt 
      Alignment       =   2  '��������
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   6
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label IN_Cnt 
      Alignment       =   2  '��������
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   3000
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label IN_Cnt 
      Alignment       =   2  '��������
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   3000
      TabIndex        =   4
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�݌Ƀf�[�^��"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   1440
      TabIndex        =   5
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�i�ڃ}�X�^��"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
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
      Width           =   240
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "�f�[�^�R���o�[�g�����i���с˃A�C�����j"
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
      Width           =   9120
   End
End
Attribute VB_Name = "CONV2004_ONO1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function Update_Proc() As Integer

Dim sts             As Integer
Dim Upd_com         As Integer
Dim com             As Integer
Dim ans             As Integer
Dim IN_Count        As Long
Dim OUT_Count       As Long

Dim DISP_INTERVAL   As Long


Dim MTS_CODE        As String * 8
Dim c               As String * 128

    Update_Proc = True

'---------------------------------------------  �i�ڃ}�X�^�̃R���o�[�g
    
    Call Log_Out(LOG_F, "�i�ڃ}�X�^�R���o�[�g�J�n=" & Format(Now, "HH:MM:SS"))
    
    
    MsgLab(1) = "�i�ڃ}�X�^�R���o�[�g�������I�I"
    Me.MousePointer = vbHourglass
    IN_Count = 0
    OUT_Count = 0

    IN_Cnt(0).Caption = Format(IN_Count, "#0")
    OUT_Cnt(0).Caption = Format(OUT_Count, "#0")


    com = BtOpGetFirst
    Do

        DoEvents

        sts = BTRV(com, OLD_ITEM2_POS, OLD_ITEM2REC, Len(OLD_ITEM2REC), K0_OLD_ITEM2, Len(K0_OLD_ITEM2), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�i���Q�j�i�ڃ}�X�^")
                Exit Function
        End Select


        IN_Count = IN_Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            IN_Cnt(0).Caption = Format(IN_Count, "#0")
            DISP_INTERVAL = 0
        End If




        If StrConv(OLD_ITEM2REC.JGYOBU, vbUnicode) = SUIHAN Then
            sts = GetIni(App.EXEName, Trim(StrConv(OLD_ITEM2REC.HIN_GAI, vbUnicode)), "SETUP", c)
            If Not sts Then
                Call UniCode_Conv(OLD_ITEM2REC.JGYOBU, SENTAKU)
                OUT_Count = OUT_Count + 1
            End If
        End If
        Do
            sts = BTRV(BtOpInsert, ITEM_POS, OLD_ITEM2REC, Len(OLD_ITEM2REC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "�i�ڃ}�X�^" & "[" & StrConv(OLD_ITEM2REC.JGYOBU, vbUnicode) & "-" & StrConv(OLD_ITEM2REC.NAIGAI, vbUnicode) & "-" & StrConv(OLD_ITEM2REC.HIN_GAI, vbUnicode))
                    Exit Function
            End Select
        Loop



        com = BtOpGetNext

    Loop

    OUT_Cnt(0).Caption = Format(OUT_Count, "#0")    '���с˃A�C�����X�V����


'---------------------------------------------  �݌Ƀf�[�^�̃R���o�[�g
    Call Log_Out(LOG_F, "�݌Ƀf�[�^�R���o�[�g�J�n=" & Format(Now, "HH:MM:SS"))
    
    
    MsgLab(1) = "�݌Ƀf�[�^�R���o�[�g�������I�I"
    Me.MousePointer = vbHourglass
    IN_Count = 0
    OUT_Count = 0
    DISP_INTERVAL = 0


    IN_Cnt(1).Caption = Format(IN_Count, "#0")
    OUT_Cnt(1).Caption = Format(OUT_Count, "#0")


    com = BtOpGetFirst
    Do

        DoEvents

        sts = BTRV(com, OLD_ZAIKO2_POS, OLD_ZAIKO2REC, Len(OLD_ZAIKO2REC), K0_OLD_ZAIKO2, Len(K0_OLD_ZAIKO2), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�i���Q�j�݌Ƀf�[�^")
                Exit Function
        End Select

        IN_Count = IN_Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            IN_Cnt(1).Caption = Format(IN_Count, "#0")
            DISP_INTERVAL = 0
        End If


        If StrConv(OLD_ZAIKO2REC.JGYOBU, vbUnicode) = SUIHAN Then
            sts = GetIni(App.EXEName, Trim(StrConv(OLD_ZAIKO2REC.HIN_GAI, vbUnicode)), "SETUP", c)
            If Not sts Then
                Call UniCode_Conv(OLD_ZAIKO2REC.JGYOBU, SENTAKU)

                OUT_Count = OUT_Count + 1

            End If
        End If

        Do
            sts = BTRV(BtOpInsert, ZAIKO_POS, OLD_ZAIKO2REC, Len(OLD_ZAIKO2REC), K0_ZAIKO, Len(K0_ZAIKO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "�݌Ƀf�[�^")
                    Exit Function
            End Select
        Loop



        com = BtOpGetNext

    Loop

    OUT_Cnt(1).Caption = Format(OUT_Count, "#0")    '���с˃A�C�����X�V����

'---------------------------------------------  �݌Ɉړ����̃R���o�[�g
    Call Log_Out(LOG_F, "�݌Ɉړ����R���o�[�g�J�n=" & Format(Now, "HH:MM:SS"))
    
    MsgLab(1) = "�݌Ɉړ����R���o�[�g�������I�I"
    Me.MousePointer = vbHourglass
    IN_Count = 0
    OUT_Count = 0
    DISP_INTERVAL = 0

    IN_Cnt(2).Caption = Format(IN_Count, "#0")
    OUT_Cnt(2).Caption = Format(OUT_Count, "#0")


    com = BtOpGetFirst
    Do

        DoEvents

        sts = BTRV(com, OLD_IDO2_POS, OLD_IDO2REC, Len(OLD_IDO2REC), K0_OLD_IDO2, Len(K0_OLD_IDO2), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�i���Q�j�݌Ɉړ���")
                Exit Function
        End Select


        IN_Count = IN_Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            IN_Cnt(2).Caption = Format(IN_Count, "#0")
            DISP_INTERVAL = 0
        End If


        If StrConv(OLD_IDO2REC.JGYOBU, vbUnicode) = SUIHAN Then
            sts = GetIni(App.EXEName, Trim(StrConv(OLD_IDO2REC.HIN_GAI, vbUnicode)), "SETUP", c)
            If Not sts Then

                Call UniCode_Conv(OLD_IDO2REC.JGYOBU, SENTAKU)

                OUT_Count = OUT_Count + 1

            End If

        End If
        Do
            sts = BTRV(BtOpInsert, IDO_POS, OLD_IDO2REC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<IDOREKI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "�݌Ɉړ���")
                    Exit Function
            End Select
        Loop



        com = BtOpGetNext

    Loop
    
    OUT_Cnt(2).Caption = Format(OUT_Count, "#0")    '���с˃A�C�����X�V����

'---------------------------------------------  �o�ח\��̃R���o�[�g
    Call Log_Out(LOG_F, "�o�ח\��R���o�[�g�J�n=" & Format(Now, "HH:MM:SS"))
    
    MsgLab(1) = "�o�ח\��f�[�^�R���o�[�g�������I�I"
    Me.MousePointer = vbHourglass
    IN_Count = 0
    OUT_Count = 0
    DISP_INTERVAL = 0
    
    IN_Cnt(3).Caption = Format(IN_Count, "#0")
    OUT_Cnt(3).Caption = Format(OUT_Count, "#0")
                                        
                                        
    com = BtOpGetFirst
    Do
        
        DoEvents
        
        sts = BTRV(com, OLD_Y_SYU2_POS, OLD_Y_SYU2REC, Len(OLD_Y_SYU2REC), K0_OLD_Y_SYU2, Len(K0_OLD_Y_SYU2), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�i���Q�j�o�ח\��f�[�^")
                Exit Function
        End Select
        
        
        IN_Count = IN_Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            IN_Cnt(3).Caption = Format(IN_Count, "#0")
            DISP_INTERVAL = 0
        End If

        
        If StrConv(OLD_Y_SYU2REC.JGYOBU, vbUnicode) = SUIHAN Then
            sts = GetIni(App.EXEName, Trim(StrConv(OLD_Y_SYU2REC.HIN_NO, vbUnicode)), "SETUP", c)
            If Not sts Then
                Call UniCode_Conv(OLD_Y_SYU2REC.JGYOBU, SENTAKU)
                OUT_Count = OUT_Count + 1
            End If
        End If
            
        Do
            sts = BTRV(BtOpInsert, Y_SYU_POS, OLD_Y_SYU2REC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "�o�ח\��")
                    Exit Function
            End Select
        Loop
        

        com = BtOpGetNext
    
    Loop

    OUT_Cnt(3).Caption = Format(OUT_Count, "#0")    '���с˃A�C�����X�V����


'---------------------------------------------  �I��
    Call Log_Out(LOG_F, "�R���o�[�g�I��=" & Format(Now, "HH:MM:SS"))
    
    
    Update_Proc = False

End Function

Private Sub Form_Activate()

Dim ans As Integer
                                
                                '�����I��
    Beep
    ans = MsgBox("���s���܂����H", vbYesNo + vbQuestion, "�m�F����")
    If ans = vbYes Then
        If Update_Proc() Then
            Unload Me
        End If
    End If
    MsgBox "�I�����܂����B"
    Unload Me

End Sub

Private Sub Form_DblClick()
    PrintForm
End Sub
Private Sub Form_Load()
Dim i As Integer
Dim c As String * 128
Dim sts As Integer
    
    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If
   
    Show
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i���j�i�ڃ}�X�^�n�o�d�m
    If OLD_ITEM2_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                '�݌Ƀf�[�^�n�o�d�m
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i���j�݌Ƀf�[�^�n�o�d�m
    
    If OLD_ZAIKO2_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ɉړ����f�[�^�n�o�d�m
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i���j�݌Ɉړ����f�[�^�n�o�d�m
    If OLD_IDO2_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�o�ח\��f�[�^�n�o�d�m
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i���j�o�ח\��f�[�^�n�o�d�m
    If OLD_Y_SYU2_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                    
    
    
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
                                            '(��)�i�ڃ}�X�^CLOSE
    sts = BTRV(BtOpClose, OLD_ITEM2_POS, OLD_ITEM2REC, Len(OLD_ITEM2REC), K0_OLD_ITEM2, Len(K0_OLD_ITEM2), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i���j�i�ڃ}�X�^")
        End If
    End If
                                            '�݌Ƀf�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
        End If
    End If
                                            '(��)�݌Ƀf�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, OLD_ZAIKO2_POS, OLD_ZAIKO2REC, Len(OLD_ZAIKO2REC), K0_OLD_ZAIKO2, Len(K0_OLD_ZAIKO2), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(��)�݌Ƀf�[�^")
        End If
    End If
    
                                            '�݌Ɉړ����b�k�n�r�d
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ɉړ���")
        End If
    End If
                                            '(��)�݌Ɉړ����b�k�n�r�d
    sts = BTRV(BtOpClose, OLD_IDO2_POS, OLD_IDO2REC, Len(OLD_IDO2REC), K0_OLD_IDO2, Len(K0_OLD_IDO2), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(��)�݌Ɉړ���")
        End If
    End If
                                            '�o�ח\��f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�o�ח\��f�[�^")
        End If
    End If
                                            '(��)�o�ח\��f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, OLD_Y_SYU2_POS, OLD_Y_SYU2REC, Len(OLD_Y_SYU2REC), K0_OLD_Y_SYU2, Len(K0_OLD_Y_SYU2), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(��)�o�ח\��f�[�^")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set CONV2004_ONO1 = Nothing

    End
End Sub

