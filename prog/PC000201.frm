VERSION 5.00
Begin VB.Form PC000201 
   BackColor       =   &H00C0C0C0&
   Caption         =   "�\���}�X�^�R���o�[�g����"
   ClientHeight    =   7230
   ClientLeft      =   2325
   ClientTop       =   2625
   ClientWidth     =   9120
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
   ScaleWidth      =   9120
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   5640
      Width           =   8055
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   4920
      Width           =   8055
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   3360
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '�E����
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�\���}�X�^��"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Top             =   3000
      Width           =   1695
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
      Caption         =   "�f�[�^�R���o�[�g����"
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
      Width           =   4800
   End
End
Attribute VB_Name = "PC000201"
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
Dim Count           As Long

Dim DISP_INTERVAL   As Long


Dim FileNo          As Long
Dim fileName        As String


Dim COMPO_REC       As Variant
Dim RecordBuf       As String

Dim c               As String * 128

Dim i               As Integer

Dim SEQNO           As Integer

Dim Err_FLg         As Boolean



    Update_Proc = True

    FileNo = FreeFile
    
                                '���O�t�@�C������荞��
    If GetIni("FILE", "COMPO_TXT", "CONV2006", c) Then
        Beep
        MsgBox "[COMPO_TXT]�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        Unload Me
    End If
    fileName = RTrim(c)
    
        
    Open fileName For Input As FileNo
    
    
    
    
    
    MsgLab(1) = "�\���}�X�^�R���o�[�g�������I�I"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(0).Caption = Format(Count, "#0")
                                        
                                        
                                        
                                        
    Do Until EOF(FileNo)
        
        DoEvents
        
        Line Input #FileNo, RecordBuf
        
        
        Err_FLg = False
        
        COMPO_REC = Split(RecordBuf, vbTab, -1)
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(0).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        
        '---------------------------------------------------------- ͯ�ްں���
        Call UniCode_Conv(P_COMPO_O_REC.SHIMUKE_CODE, CStr(COMPO_REC(1)))           '�d�����溰��
                                                                                            
        '�R�[�h�}�X�^�ǂݍ���
        Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN04_CD)
        Call UniCode_Conv(K0_P_CODE.C_Code, CStr(COMPO_REC(1)))
        
        sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
        Select Case sts
            Case BtNoErr
                                                                                    '���ƕ�
                Call UniCode_Conv(P_COMPO_O_REC.JGYOBU, Trim(StrConv(P_CODEREC.OPTION1, vbUnicode)))
                                                                                    '�����O
                Call UniCode_Conv(P_COMPO_O_REC.NAIGAI, Trim(StrConv(P_CODEREC.OPTION2, vbUnicode)))
            
            Case BtErrKeyNotFound
                Err_FLg = True
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�R�[�h�}�X�^")
                Exit Function
        End Select
        
        
        If Not Err_FLg Then
            Call UniCode_Conv(K0_ITEM.JGYOBU, Trim(StrConv(P_CODEREC.OPTION1, vbUnicode)))
            Call UniCode_Conv(K0_ITEM.NAIGAI, Trim(StrConv(P_CODEREC.OPTION2, vbUnicode)))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, CStr(COMPO_REC(0)))
                            
        
        
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                
                Case BtErrKeyNotFound
                    Err_FLg = True
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                    Exit Function
            End Select
        
            If Not Err_FLg Then
        
        
                Call UniCode_Conv(P_COMPO_O_REC.HIN_GAI, CStr(COMPO_REC(0)))    '�i��
                                                                                            
                Call UniCode_Conv(P_COMPO_O_REC.DATA_KBN, P_HEAD)               '�f�[�^�敪�iͯ�ް�j
                Call UniCode_Conv(P_COMPO_O_REC.SEQNO, "000")                   '�ǔ�
                                                                                            
                Call UniCode_Conv(P_COMPO_O_REC.CLASS_CODE, CStr(COMPO_REC(3))) '��{�N���X
                Call UniCode_Conv(P_COMPO_O_REC.BIKOU, Trim(CStr(COMPO_REC(68))) & "/" & Trim(CStr(COMPO_REC(69))))    '���l
Text3.Text = Trim(CStr(COMPO_REC(68)))
Text4.Text = Trim(CStr(COMPO_REC(69)))
                                                                                            
                Call UniCode_Conv(P_COMPO_O_REC.FILLER, "")                     'Filler
                
                
                Call UniCode_Conv(P_COMPO_O_REC.UPD_TANTO, "CONV")              '�X�V�S����
                                                                                '�X�V����
                Call UniCode_Conv(P_COMPO_O_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
        
                Do
                    sts = BTRV(BtOpInsert, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_COMPOREC.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                Exit Function
                            End If
                        
                        Case BtErrDuplicates
                            Call Log_Out(LOG_F, "DUP HEAD " & CStr(COMPO_REC(1)) & "-" & CStr(COMPO_REC(0)))
                            
                            Exit Do
                        
                        Case Else
                            
                            
                            
                            Call File_Error(sts, BtOpInsert, "�\���}�X�^")
                            Exit Function
                    End Select
                Loop
        
            End If
        End If
            
If Trim(COMPO_REC(0)) = "AMC00P-EW09" Then
    Debug.Print
End If
            
        If Err_FLg Then
            '�G���[����
            Call Log_Out(LOG_F, "HEAD " & CStr(COMPO_REC(1)) & "-" & CStr(COMPO_REC(0)))
    
        Else
        '---------------------------------------------------------- ������ں���
            SEQNO = 0
    
            For i = 7 To 13
            
                If COMPO_REC(i) = "" Then
                Else
                                                                                                
                    '-----------------------------------------------    �ŏ��Ͷ��Ă̎��ƕ�/�����O��
                    Call UniCode_Conv(K0_ITEM.JGYOBU, Trim(StrConv(P_CODEREC.OPTION1, vbUnicode)))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, Trim(StrConv(P_CODEREC.OPTION2, vbUnicode)))
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, CStr(COMPO_REC(i)))
        
                    Err_FLg = False
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        
                        Case BtErrKeyNotFound
                        
                            '�����O�𔽓]����
                            If Trim(StrConv(P_CODEREC.OPTION2, vbUnicode)) = NAIGAI_NAI Then
                                Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_GAI)
                            Else
                                Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                            End If
                        
                            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                            Select Case sts
                                Case BtNoErr
                                
                                Case BtErrKeyNotFound
                                
                                    '���ނƂ���
                                    Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                                    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                                
                                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                    Select Case sts
                                        Case BtNoErr
                                        
                                        Case BtErrKeyNotFound
                                        
'                                            Err_FLg = True
                                        
                                            Call UniCode_Conv(ITEMREC.JGYOBU, SHIZAI)
                                            Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)
                                        
                                            Call UniCode_Conv(ITEMREC.HIN_GAI, CStr(COMPO_REC(i)))

                                        
                                        Case Else
                                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                            Exit Function
                                    End Select
                                
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                    Exit Function
                            End Select
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                            Exit Function
                    End Select
        
                    If Not Err_FLg Then
        
                        Call UniCode_Conv(P_COMPO_K_REC.SHIMUKE_CODE, CStr(COMPO_REC(1)))   '�d�����溰��
                                                                                    '���ƕ�
                        Call UniCode_Conv(P_COMPO_K_REC.JGYOBU, Trim(StrConv(P_CODEREC.OPTION1, vbUnicode)))
                                                                                    '�����O
                        Call UniCode_Conv(P_COMPO_K_REC.NAIGAI, Trim(StrConv(P_CODEREC.OPTION2, vbUnicode)))
                        
                        Call UniCode_Conv(P_COMPO_K_REC.HIN_GAI, CStr(COMPO_REC(0)))        '�i��
                        
                        
                        Call UniCode_Conv(P_COMPO_K_REC.DATA_KBN, P_KOSOU)                                  '�f�[�^�敪�i�����ށj
                        
                        SEQNO = SEQNO + 10
                        Call UniCode_Conv(P_COMPO_K_REC.SEQNO, Format(SEQNO, "000"))                        '�ǔ�
                        
                        Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, "")                                    '�q�@���
                        Call UniCode_Conv(P_COMPO_K_REC.KO_JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))      '�q�@���ƕ�
                        Call UniCode_Conv(P_COMPO_K_REC.KO_NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))      '�q�@�����O
                        Call UniCode_Conv(P_COMPO_K_REC.KO_HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))    '�q�@�i��
                                
                        If IsNumeric(COMPO_REC(i + 7)) Then                                 '�q�@����
                            Call UniCode_Conv(P_COMPO_K_REC.KO_QTY, Format(CDbl(COMPO_REC(i + 7)), "000.00"))
                        Else
                            Call UniCode_Conv(P_COMPO_K_REC.KO_QTY, "001.00")
                        End If
                        
                        Call UniCode_Conv(P_COMPO_K_REC.KO_BIKOU, "")                                       '�q�@���l
                        Call UniCode_Conv(P_COMPO_K_REC.UPD_TANTO, "CONV")                                  '�X�V�S����
                                                                                                            '�X�V����
                        Call UniCode_Conv(P_COMPO_K_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                                                                                            
                                                                                            
                                                                                            
                        Do
                            sts = BTRV(BtOpInsert, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                    Beep
                                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_COMPOREC.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                    If ans = vbCancel Then
                                        Exit Function
                                    End If
                                
                                Case BtErrDuplicates
                                    Call Log_Out(LOG_F, "DUP ������ " & CStr(COMPO_REC(1)) & "-" & CStr(COMPO_REC(0)) & "-" & CStr(COMPO_REC(i)))
                                    
                                    Exit Do
                                
                                Case Else
                                    
                                    
                                    
                                    Call File_Error(sts, BtOpInsert, "�\���}�X�^")
                                    Exit Function
                            End Select
                        Loop
                                                                                            
                                                                                            
                    Else
                        '�G���[����
                        Call Log_Out(LOG_F, "������ " & CStr(COMPO_REC(1)) & "-" & CStr(COMPO_REC(0)) & "-" & CStr(COMPO_REC(i)))
                    End If
                                                                                    
                End If
            Next i
    
        '---------------------------------------------------------- �O������ں���
            SEQNO = 0
    
            For i = 21 To 23
            
                If COMPO_REC(i) = "" Then
                Else
                                                                                                
                    '-----------------------------------------------    �ŏ��Ͷ��Ă̎��ƕ�/�����O��
                    Call UniCode_Conv(K0_ITEM.JGYOBU, Trim(StrConv(P_CODEREC.OPTION1, vbUnicode)))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, Trim(StrConv(P_CODEREC.OPTION2, vbUnicode)))
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, CStr(COMPO_REC(i)))
        
                    Err_FLg = False
                    
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        
                        Case BtErrKeyNotFound
                        
                            '�����O�𔽓]����
                            Call UniCode_Conv(K0_ITEM.JGYOBU, Trim(StrConv(P_CODEREC.OPTION1, vbUnicode)))
                            
                            If Trim(StrConv(P_CODEREC.OPTION2, vbUnicode)) = NAIGAI_NAI Then
                                Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_GAI)
                            Else
                                Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                            End If
                        
                            Call UniCode_Conv(K0_ITEM.HIN_GAI, CStr(COMPO_REC(i)))
                            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                            Select Case sts
                                Case BtNoErr
                                
                                Case BtErrKeyNotFound
                                
                                    '���ނƂ���
                                    Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                                    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                                    Call UniCode_Conv(K0_ITEM.HIN_GAI, CStr(COMPO_REC(i)))
                                
                                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                    Select Case sts
                                        Case BtNoErr
                                        
                                        Case BtErrKeyNotFound
                                        
'                                            Err_FLg = True
                                        
                                            Call UniCode_Conv(ITEMREC.JGYOBU, SHIZAI)
                                            Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)
                                        
                                            Call UniCode_Conv(ITEMREC.HIN_GAI, CStr(COMPO_REC(i)))
                                        
                                        
                                        Case Else
                                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                            Exit Function
                                    End Select
                                
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                    Exit Function
                            End Select
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                            Exit Function
                    End Select
                    
                    If Not Err_FLg Then
                    
                        Call UniCode_Conv(P_COMPO_K_REC.SHIMUKE_CODE, CStr(COMPO_REC(1)))   '�d�����溰��
                                                                                            '���ƕ�
                        Call UniCode_Conv(P_COMPO_K_REC.JGYOBU, Trim(StrConv(P_CODEREC.OPTION1, vbUnicode)))
                                                                                            '�����O
                        Call UniCode_Conv(P_COMPO_K_REC.NAIGAI, Trim(StrConv(P_CODEREC.OPTION2, vbUnicode)))
                        
                        Call UniCode_Conv(P_COMPO_K_REC.HIN_GAI, CStr(COMPO_REC(0)))        '�i��
        
            
                            
Text1.Text = CStr(COMPO_REC(0))
Text1.Text = StrConv(ITEMREC.HIN_GAI, vbUnicode)
                            
                            
                        Call UniCode_Conv(P_COMPO_K_REC.DATA_KBN, P_GAISOU)                                 '�f�[�^�敪�i�O�����ށj
                        SEQNO = SEQNO + 10
                        Call UniCode_Conv(P_COMPO_K_REC.SEQNO, Format(SEQNO, "000"))                        '�ǔ�
                        
                        Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, "")                                    '�q�@���
                        Call UniCode_Conv(P_COMPO_K_REC.KO_JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))      '�q�@���ƕ�
                        Call UniCode_Conv(P_COMPO_K_REC.KO_NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))      '�q�@�����O
                        Call UniCode_Conv(P_COMPO_K_REC.KO_HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))    '�q�@�i��
                                
                        If IsNumeric(COMPO_REC(i + 3)) Then                                                 '�q�@����
                            Call UniCode_Conv(P_COMPO_K_REC.KO_QTY, Format(CDbl(COMPO_REC(i + 3)), "000.00"))
                        Else
                            Call UniCode_Conv(P_COMPO_K_REC.KO_QTY, "001.00")
                        End If
                        
                        Call UniCode_Conv(P_COMPO_K_REC.KO_BIKOU, "")                                       '�q�@���l
                        Call UniCode_Conv(P_COMPO_K_REC.UPD_TANTO, "CONV")                                  '�X�V�S����
                                                                                                            '�X�V����
                        Call UniCode_Conv(P_COMPO_K_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                                                                                            
                        Do
                            sts = BTRV(BtOpInsert, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                    Beep
                                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_COMPOREC.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                    If ans = vbCancel Then
                                        Exit Function
                                    End If
                                
                                Case BtErrDuplicates
                                    
                                    Call Log_Out(LOG_F, "DUP �O������ " & CStr(COMPO_REC(1)) & "-" & CStr(COMPO_REC(0)) & "-" & CStr(COMPO_REC(i)))
                                    Exit Do
                                
                                Case Else
                                    
                                    
                                    
                                    Call File_Error(sts, BtOpInsert, "�\���}�X�^")
                                    Exit Function
                            End Select
                        Loop
                                                                                            
                    Else
                        '�G���[����
                        Call Log_Out(LOG_F, "�O������ " & CStr(COMPO_REC(1)) & "-" & CStr(COMPO_REC(0)) & "-" & CStr(COMPO_REC(i)))
                    End If
                                                                                    
                End If
            Next i
    
        '---------------------------------------------------------- �����^�\�����iں���
            SEQNO = 0
    
            For i = 27 To 46
            
                If COMPO_REC(i) = "" Then
                Else
                                                                                                
                    '-----------------------------------------------    �ŏ��Ͷ��Ă̎��ƕ�/�����O��
                    Call UniCode_Conv(K0_ITEM.JGYOBU, Trim(StrConv(P_CODEREC.OPTION1, vbUnicode)))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, Trim(StrConv(P_CODEREC.OPTION2, vbUnicode)))
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, CStr(COMPO_REC(i)))
                    
                    Err_FLg = False
        
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        
                        Case BtErrKeyNotFound
                        
                            '�����O�𔽓]����
                            If Trim(StrConv(P_CODEREC.OPTION2, vbUnicode)) = NAIGAI_NAI Then
                                Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                            Else
                                Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_GAI)
                            End If
                        
                            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                            Select Case sts
                                Case BtNoErr
                                
                                Case BtErrKeyNotFound
                                
                                    '���ނƂ���
                                    Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                                    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                                
                                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                    Select Case sts
                                        Case BtNoErr
                                        
                                        Case BtErrKeyNotFound
                                        
'                                            Err_FLg = True
                                            Call UniCode_Conv(ITEMREC.JGYOBU, SHIZAI)
                                            Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)
                                        
                                            Call UniCode_Conv(ITEMREC.HIN_GAI, CStr(COMPO_REC(i)))
                                        
                                        
                                        Case Else
                                            
                                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                            Exit Function
                                    End Select
                                
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                    Exit Function
                            End Select
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                            Exit Function
                    End Select
        
                    If Not Err_FLg Then
                        
                        
                        Call UniCode_Conv(P_COMPO_K_REC.SHIMUKE_CODE, CStr(COMPO_REC(1)))   '�d�����溰��
                        
                                                                                            '���ƕ�
                        Call UniCode_Conv(P_COMPO_K_REC.JGYOBU, Trim(StrConv(P_CODEREC.OPTION1, vbUnicode)))
                                                                                            '�����O
                        Call UniCode_Conv(P_COMPO_K_REC.NAIGAI, Trim(StrConv(P_CODEREC.OPTION2, vbUnicode)))
                        
                        Call UniCode_Conv(P_COMPO_K_REC.HIN_GAI, CStr(COMPO_REC(0)))        '�i��
                        
                        
                        Call UniCode_Conv(P_COMPO_K_REC.DATA_KBN, P_DOUKON)                                 '�f�[�^�敪�i�����^�\���j
                        SEQNO = SEQNO + 10
                        Call UniCode_Conv(P_COMPO_K_REC.SEQNO, Format(SEQNO, "000"))                        '�ǔ�
                        
                        Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, "")                                    '�q�@���
                        Call UniCode_Conv(P_COMPO_K_REC.KO_JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))      '�q�@���ƕ�
                        Call UniCode_Conv(P_COMPO_K_REC.KO_NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))      '�q�@�����O
                        Call UniCode_Conv(P_COMPO_K_REC.KO_HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))    '�q�@�i��
                                
                        If IsNumeric(COMPO_REC(i + 20)) Then                                                '�q�@����
                            Call UniCode_Conv(P_COMPO_K_REC.KO_QTY, Format(CDbl(COMPO_REC(i + 20)), "000.00"))
                        Else
                            Call UniCode_Conv(P_COMPO_K_REC.KO_QTY, "001.00")
                        End If
                        
                        Call UniCode_Conv(P_COMPO_K_REC.KO_BIKOU, "")                                       '�q�@���l
                        Call UniCode_Conv(P_COMPO_K_REC.UPD_TANTO, "CONV")                                  '�X�V�S����
                                                                                                            '�X�V����
                        Call UniCode_Conv(P_COMPO_K_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                                                                                            
                        Do
                            sts = BTRV(BtOpInsert, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                    Beep
                                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_COMPOREC.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                    If ans = vbCancel Then
                                        Exit Function
                                    End If
                                
                                Case BtErrDuplicates
                                    Call Log_Out(LOG_F, "DUP �����^�\�� " & CStr(COMPO_REC(1)) & "-" & CStr(COMPO_REC(0)) & "-" & CStr(COMPO_REC(i)))
                                    
                                    Exit Do
                                
                                Case Else
                                    
                                    
                                    
                                    Call File_Error(sts, BtOpInsert, "�\���}�X�^")
                                    Exit Function
                            End Select
                        Loop
                    Else
                        '�G���[����
                        Call Log_Out(LOG_F, "�����^�\�� " & CStr(COMPO_REC(1)) & "-" & CStr(COMPO_REC(0)) & "-" & CStr(COMPO_REC(i)))
                    End If
                                                                                    
                End If
            Next i
    
    
    
        End If
    
    
    Loop

    Cnt(0).Caption = Format(Count, "#0")


    MsgBox "���ްďI���I�I"

    Close #FileNo



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
                                '�\���}�X�^�n�o�d�m
    If P_COMPO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�R�[�h�}�X�^�n�o�d�m
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '�\���}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�\���}�X�^")
        End If
    End If
                                            '�R�[�h�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�R�[�h�}�X�^")
        End If
    End If
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set PC000201 = Nothing

    End
End Sub

