VERSION 5.00
Begin VB.Form PC000101 
   BackColor       =   &H00C0C0C0&
   Caption         =   "�i�ڃ}�X�^�R���o�[�g����"
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
   Begin VB.CommandButton Command1 
      Caption         =   "�I��"
      Height          =   495
      Index           =   3
      Left            =   5400
      TabIndex        =   11
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   2
      Left            =   1560
      TabIndex        =   10
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   9
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   8
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "�i�ڃ}�X�^(���i)��"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   7
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '�E����
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   4680
      TabIndex        =   6
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '�E����
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   5
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�i�ڃ}�X�^(����)��"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   4
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '�E����
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   4680
      TabIndex        =   3
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�i�ڃ}�X�^(����)��"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   2
      Top             =   2880
      Width           =   2295
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
Attribute VB_Name = "PC000101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type KAISHA_Tbl_Tag
    C_Code          As String * 2
    C_NAME          As String
    JGYOBU          As String * 1
    NAIGAI          As String * 1
End Type


Private KAISHA_Tbl()    As KAISHA_Tbl_Tag



Private Function Update_Proc(SHORI_MODE) As Integer

Dim sts             As Integer
Dim Upd_com         As Integer
Dim com             As Integer
Dim ans             As Integer
Dim Count           As Long

Dim DISP_INTERVAL   As Long


Dim FileNo          As Long
Dim fileName        As String


Dim ITEM_REC        As Variant
Dim RecordBuf       As String
Dim wk              As String

Dim c               As String * 128

Dim i               As Integer
Dim j               As Integer

Dim Err_Flg         As Integer

    Update_Proc = True


    Select Case SHORI_MODE
        Case 0

        FileNo = FreeFile
        
                                    '���O�t�@�C������荞��
        If GetIni("FILE", "SHIZAI_TXT", "CONV2006", c) Then
            Beep
            MsgBox "[SHIZAI_TXT]�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            Unload Me
        End If
        fileName = RTrim(c)
        
            
        Open fileName For Input As FileNo
        
        
        
    '-----------------------------------------------------------------------------  �Ɩ��Ǘ��|�|���o�n�r
        
        
        MsgLab(1) = "���ށ@���i�}�X�^�R���o�[�g�������I�I"
        Me.MousePointer = vbHourglass
        Count = 0
        DISP_INTERVAL = 0
        Cnt(0).Caption = Format(Count, "#0")
                                            
                                            
                                            
                                            
        Do Until EOF(FileNo)
            
            DoEvents
            
            Line Input #FileNo, RecordBuf
            
            ITEM_REC = Split(RecordBuf, vbTab, -1)
            
            
            Count = Count + 1
            DISP_INTERVAL = DISP_INTERVAL + 1
            If DISP_INTERVAL = 100 Then
                Cnt(0).Caption = Format(Count, "#0")
                DISP_INTERVAL = 0
            End If
            
            
            For i = 0 To UBound(ITEM_REC)
                For j = 0 To Len(ITEM_REC(i))
    
                    If Mid(CStr(ITEM_REC(i)), j + 1, 1) = """" Then
                        Mid(ITEM_REC(i), j + 1, 1) = " "
                    End If
    
                Next j
    
                ITEM_REC(i) = Trim(ITEM_REC(i))
            Next i
            
            Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)               '���ƕ�(=����)
            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)           '�����O(=����)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, CStr(ITEM_REC(1)))   '���ޕi��
            
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                
                    Upd_com = BtOpUpdate
                
                
                Case BtErrKeyNotFound
                    
                    Upd_com = BtOpInsert
                
                Case Else
                    
                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                    Exit Function
            End Select
            
                    
            If Upd_com = BtOpInsert Then
                    
                Call UniCode_Conv(ITEMREC.JGYOBU, SHIZAI)       '���ƕ�=����
                Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)   '�����O
                Call UniCode_Conv(ITEMREC.HIN_GAI, CStr(ITEM_REC(1)))       '�i�ں���
                Call UniCode_Conv(ITEMREC.ST_SET_DT, "")                    '�W���I�Ԑݒ���t
                Call UniCode_Conv(ITEMREC.ST_SOKO, "")                      '�W�����Ɂ@�q��
                Call UniCode_Conv(ITEMREC.ST_RETU, "")                      '�W�����Ɂ@��
                Call UniCode_Conv(ITEMREC.ST_REN, "")                       '�W�����Ɂ@�A
                Call UniCode_Conv(ITEMREC.ST_DAN, "")                       '�W�����Ɂ@�i
                Call UniCode_Conv(ITEMREC.BEF_SOKO, "")                     '�O����Ɂ@�q��
                Call UniCode_Conv(ITEMREC.BEF_RETU, "")                     '�O����Ɂ@��
                Call UniCode_Conv(ITEMREC.BEF_REN, "")                      '�O����Ɂ@�A
                Call UniCode_Conv(ITEMREC.BEF_DAN, "")                      '�O����Ɂ@�i
                Call UniCode_Conv(ITEMREC.LAST_NYU_DT, "")                  '�ŏI���ɓ�
                Call UniCode_Conv(ITEMREC.LAST_SYU_DT, "")                  '�ŏI�o�ɓ�
                Call UniCode_Conv(ITEMREC.HIN_NAI, "")                      '�i�ԁi���j
                Call UniCode_Conv(ITEMREC.BIKOU_SOKO, "")                   'νđq��
                Call UniCode_Conv(ITEMREC.BIKOU_TANA, "")                   'νĒI��
                Call UniCode_Conv(ITEMREC.HOJYU_P, "00000000")              '��[�_
                Call UniCode_Conv(ITEMREC.AVE_SYUKA, "00000000")            '�����Ϗo�א�
                Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "0")                  '����ِ�
                Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "0")                  '����ِ�
                Call UniCode_Conv(ITEMREC.LAST_INP_DT, "")                  '�ŏI���ד��t
                Call UniCode_Conv(ITEMREC.LAST_CHK_DT, "")                  '�ŏI�ƍ����t
                Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, "00000000")         '�ƍ����݌ɐ�
                Call UniCode_Conv(ITEMREC.BIKOU, "")                        '������l
                Call UniCode_Conv(ITEMREC.IRI_QTY, "")                      '������萔
                Call UniCode_Conv(ITEMREC.JAN_CODE, "")                     'JAN����
                Call UniCode_Conv(ITEMREC.HIN_CHANGE, "")                   '�i�ԓǂݑւ�����
                Call UniCode_Conv(ITEMREC.GOODS_KBN, "1")                   '���i���L��
                Call UniCode_Conv(ITEMREC.PACKING_NO, "")                   '������
                Call UniCode_Conv(ITEMREC.RANK, "")                         '�����ݸ
                Call UniCode_Conv(ITEMREC.NEW_RANK, "")                     '�V�ݸ
                Call UniCode_Conv(ITEMREC.GLICS1_TANA, "")                  '��د���I��1
                Call UniCode_Conv(ITEMREC.GLICS2_TANA, "")                  '��د���I��2
                Call UniCode_Conv(ITEMREC.GLICS3_TANA, "")                  '��د���I��3
            
                Call UniCode_Conv(ITEMREC.L_HIN_NAME_E, "")                 '�i��E
                Call UniCode_Conv(ITEMREC.L_BIKOU, "")                      '���l
                Call UniCode_Conv(ITEMREC.L_KAISHA_CODE, "")                '��Ж�
                Call UniCode_Conv(ITEMREC.L_KISHU1, "")                     '�@��(1)
                Call UniCode_Conv(ITEMREC.L_KISHU2, "")                     '�@��(2)
                Call UniCode_Conv(ITEMREC.L_KISHU3, "")                     '�@��(3)
                Call UniCode_Conv(ITEMREC.L_PAPER, "")                      '��
                Call UniCode_Conv(ITEMREC.L_PLASTIC, "")                    '��׽���
                Call UniCode_Conv(ITEMREC.L_URIKIN1, "")                    '���i(1)
                Call UniCode_Conv(ITEMREC.L_URIKIN2, "")                    '���i(2)
                Call UniCode_Conv(ITEMREC.L_URIKIN3, "")                    '���i(3)
                Call UniCode_Conv(ITEMREC.L_LABEL, "")                      '�K�p�@������
                Call UniCode_Conv(ITEMREC.L_MAISU, "")                      '���ٖ���
                Call UniCode_Conv(ITEMREC.L_KISHU_BIKOU, "")                '�K�p�@����l
                Call UniCode_Conv(ITEMREC.L_SAGYO_SHIJI, "")                '��Ǝw��
                Call UniCode_Conv(ITEMREC.L_BIKOU3, "")                     '���l(3)
                Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, "")                '���ƕ���
                Call UniCode_Conv(ITEMREC.L_IRI_QTY, "")                    '���萔
                Call UniCode_Conv(ITEMREC.L_TANA1, "")                      '�I��(1)
                Call UniCode_Conv(ITEMREC.L_TANA2, "")                      '�I��(2)
                
                Call UniCode_Conv(ITEMREC.S_TANTO, "")                      '���P�^�S����
                
                Call UniCode_Conv(ITEMREC.ZAIKO_F, P_ZAIKO_F_OFF)            '�݌ɊǗ��Ώ�
                
                Call UniCode_Conv(ITEMREC.FILLER, "")                       'Filler
        
            End If
            
            If IsNumeric(ITEM_REC(33)) Then                                 '�댯�݌�
                Call UniCode_Conv(ITEMREC.HOJYU_P, Format(CDbl(ITEM_REC(33)), "00000000000"))
            Else
                Call UniCode_Conv(ITEMREC.HOJYU_P, "00000000000")
            End If
            
            
            Call UniCode_Conv(ITEMREC.HIN_NAME, CStr(ITEM_REC(2)))          '�i��
            Call UniCode_Conv(ITEMREC.G_SHIIRE_KBN, CStr(ITEM_REC(3)))      '�d���敪
            Call UniCode_Conv(ITEMREC.G_HANBAI_KBN, CStr(ITEM_REC(4)))      '�̔��敪
            Call UniCode_Conv(ITEMREC.G_SYUSHI, CStr(ITEM_REC(5)))          '���x�P��
            If CStr(ITEM_REC(6)) = "" Then                                  '�g�����i
                Call UniCode_Conv(ITEMREC.G_KUMITATE, P_ASSEMBLY_OFF)
            Else
                Call UniCode_Conv(ITEMREC.G_KUMITATE, P_ASSEMBLY_ON)
            End If
            If IsNumeric(CStr(ITEM_REC(7))) Then
                                                                            '�����P��
                Call UniCode_Conv(ITEMREC.G_ST_URITAN, Format(CDbl(ITEM_REC(7)), "00000000.00"))
                                                                            '�����ݒ��
                
                If IsDate(ITEM_REC(8)) Then
                    Call UniCode_Conv(ITEMREC.G_ST_URITAN_DT, Format(CStr(ITEM_REC(8)), "YYYYMMDD"))
                Else
                    Call UniCode_Conv(ITEMREC.G_ST_URITAN_DT, Format(Now, "YYYYMMDD"))
                End If
            Else
                                                                            '�����P��
                Call UniCode_Conv(ITEMREC.G_ST_URITAN, "")
                                                                            '�����ݒ��
                Call UniCode_Conv(ITEMREC.G_ST_URITAN_DT, "")
            End If
            
            If IsNumeric(ITEM_REC(9)) Then
                                                                            '����
                Call UniCode_Conv(ITEMREC.G_ST_SHITAN, Format(CDbl(ITEM_REC(9)), "00000000.00"))
                                                                            '�����ݒ��
                If IsDate(ITEM_REC(10)) Then
                    Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, Format(CStr(ITEM_REC(10)), "YYYYMMDD"))
                Else
                    Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, Format(Now, "YYYYMMDD"))
                End If
            Else
                                                                            '����
                Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "")
                                                                            '�����ݒ��
                Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, "")
            End If
            
            
            j = -1
            For i = 13 To 19 Step 3
            
                j = j + 1
                
                If j = 0 Then
                
                    If IsNumeric(ITEM_REC(34)) Then
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LOT, CLng(ITEM_REC(34)))                  'ۯĐ�
                    Else
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LOT, "")                  'ۯĐ�
                    End If
                    
                    Select Case Trim(CStr(ITEM_REC(35)))
                        Case "D"
                                        
                            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LEAD_TIME, "007")            'ذ�����
                    
                        Case "F"
                                        
                            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LEAD_TIME, "010")            'ذ�����
                        Case "K"
                                        
                            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LEAD_TIME, "014")            'ذ�����
                    
                        Case "L"
                                        
                            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LEAD_TIME, "007")            'ذ�����
                    
                    
                        Case "P"
                                        
                            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LEAD_TIME, "010")            'ذ�����
                    
                        Case "Q"
                                        
                            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LEAD_TIME, "010")            'ذ�����
                    
                        Case "S"
                                        
                            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LEAD_TIME, "000")            'ذ�����
                    
                        Case Else
                            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LEAD_TIME, "")            'ذ�����
                    End Select
                
                Else
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LOT, "")                  'ۯĐ�
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LEAD_TIME, "")            'ذ�����
                End If
                Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LAST_ORDER_DT, "")        '�O�񒍕���
                Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LAST_ORDER_QTY, "")       '�O�񒍕���
                
                
                
                If ITEM_REC(i) = "" Then
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).CODE, "")             '�d����
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).TANKA, "")            '�P��
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).TANKA_DT, "")         '�P���ݒ��
                Else
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).CODE, CStr(ITEM_REC(i)))      '�d����
                    If IsNumeric(ITEM_REC(i + 1)) Then                                      '�P��
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).TANKA, _
                                                Format(CDbl(ITEM_REC(i + 1)), "00000000.00"))
                    Else
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).TANKA, "")
                    End If
                    If IsDate(ITEM_REC(i + 2)) Then                                         '�P���ݒ��
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).TANKA_DT, _
                                                Format((ITEM_REC(i + 2)), "YYYYMMDD"))
                    Else
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).TANKA_DT, "")
                    End If
            
                End If
            
            
            
            Next i
            
            
            If IsNumeric(CStr(ITEM_REC(24))) Then                                         '�O���݌ɋ��z
                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, Format(ITEM_REC(24), "00000000000"))
            Else
                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, "")
            End If
            
            Call UniCode_Conv(ITEMREC.G_SHIZAI_KBN, CStr(ITEM_REC(35)))             '���ދ敪
            
            Call UniCode_Conv(ITEMREC.G_LABEL_NON, P_G_LABEL_OFF)                   '���ٓ\��t��
            
            
            
            Call UniCode_Conv(ITEMREC.UPD_TANTO, "CONV")                    '�X�V�S����
                                                                            '�X�V����
            
            
            Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
            
            
            Do
                sts = BTRV(Upd_com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
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
                        
                        Call File_Error(sts, Upd_com, "�i�ڃ}�X�^")
                        Exit Function
                End Select
            Loop
            
        
        Loop
    '---------------------------------------------  �I��
    
        Cnt(0).Caption = Format(Count, "#0")
        
        Close #FileNo

    Case 1

'-----------------------------------------------------------------------------  ���i���x���|�|���o�n�r


        '�R�[�h�}�X�^����Ж��^���ƕ����̃Z�b�g
        Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN07_CD)
        Call UniCode_Conv(K0_P_CODE.C_Code, "")
            
        com = BtOpGetGreater
    
        i = -1
        Erase KAISHA_Tbl
    
        Do
            
            sts = BTRV(com, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            Select Case sts
                Case BtNoErr
                
                    If StrConv(P_CODEREC.DATA_KBN, vbUnicode) <> P_KBN07_CD Then
                        Exit Do
                    End If
                
                Case BtErrEOF
                    
                    Exit Do
                
                Case Else
                    
                    Call File_Error(sts, com, "�R�[�h�}�X�^")
                    Exit Function
            End Select
        
            i = i + 1
            ReDim Preserve KAISHA_Tbl(i)
        
            KAISHA_Tbl(i).C_Code = StrConv(P_CODEREC.C_Code, vbUnicode)
            KAISHA_Tbl(i).C_NAME = Trim(StrConv(P_CODEREC.C_NAME, vbUnicode))
            KAISHA_Tbl(i).JGYOBU = Trim(StrConv(P_CODEREC.OPTION1, vbUnicode))
            KAISHA_Tbl(i).NAIGAI = Trim(StrConv(P_CODEREC.OPTION2, vbUnicode))
        
        
        
        
            com = BtOpGetNext
        
        
        Loop
    
    
    
    
        FileNo = FreeFile
        
                                    '���O�t�@�C������荞��
        If GetIni("FILE", "LABEL_TXT", "CONV2006", c) Then
            Beep
            MsgBox "[LABEL_TXT]�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            Unload Me
        End If
        fileName = RTrim(c)
        
            
        Open fileName For Input As FileNo
        
        
        
        
        
        MsgLab(1) = "���i���x���@�@��}�X�^�R���o�[�g�������I�I"
        Me.MousePointer = vbHourglass
        Count = 0
        DISP_INTERVAL = 0
        Cnt(1).Caption = Format(Count, "#0")
                                            
                                            
                                            
                                            
        Do Until EOF(FileNo)
            
            DoEvents
            
            RecordBuf = ""
            
            Do
                wk = Input(1, FileNo)
                If wk = "!" Then
                    wk = Input(2, FileNo)
                    Exit Do
                End If
                RecordBuf = RecordBuf & wk
            Loop
            
            ITEM_REC = Split(RecordBuf, vbTab, -1)
            
            
            Count = Count + 1
            DISP_INTERVAL = DISP_INTERVAL + 1
            If DISP_INTERVAL = 100 Then
                Cnt(1).Caption = Format(Count, "#0")
                DISP_INTERVAL = 0
            End If
            
                    
            For i = 0 To UBound(ITEM_REC)
                For j = 0 To Len(ITEM_REC(i))
    
                    If Mid(CStr(ITEM_REC(i)), j + 1, 1) = """" Then
                        Mid(ITEM_REC(i), j + 1, 1) = " "
                    End If
    
                Next j
    
                ITEM_REC(i) = Trim(ITEM_REC(i))
            Next i
                    
            For i = 0 To UBound(KAISHA_Tbl)
            
            
                If Trim(KAISHA_Tbl(i).C_NAME) = Trim(CStr(ITEM_REC(8))) Then
                    Exit For
                End If
            
            Next i
            
            
            If i > UBound(KAISHA_Tbl) Then
                '�G���[����
                Call Log_Out(LOG_F, CStr(ITEM_REC(0)) & "-" & CStr(ITEM_REC(8)))
                
            Else
            
                Call UniCode_Conv(K0_ITEM.JGYOBU, KAISHA_Tbl(i).JGYOBU)     '���ƕ�
                Call UniCode_Conv(K0_ITEM.NAIGAI, KAISHA_Tbl(i).NAIGAI)     '�����O
                Call UniCode_Conv(K0_ITEM.HIN_GAI, CStr(ITEM_REC(0)))       '�i��
                
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                    
                        Upd_com = BtOpUpdate
                    
                    
                    Case BtErrKeyNotFound
                        
                        Upd_com = BtOpInsert
                    
                    Case Else
                        
                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                        Exit Function
                End Select
            
            
                If Upd_com = BtOpInsert Then
                
                    Call UniCode_Conv(ITEMREC.JGYOBU, KAISHA_Tbl(i).JGYOBU)                     '���ƕ�
                    Call UniCode_Conv(ITEMREC.NAIGAI, KAISHA_Tbl(i).NAIGAI)                     '�����O
                    Call UniCode_Conv(ITEMREC.HIN_GAI, CStr(ITEM_REC(0)))                       '�i�ں���
                    Call UniCode_Conv(ITEMREC.HIN_NAME, CStr(ITEM_REC(1)))                      '�i�ږ���
                    
                    Call UniCode_Conv(ITEMREC.ST_SET_DT, "")                                    '�W���I�Ԑݒ���t
                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")                                      '�W�����Ɂ@�q��
                    Call UniCode_Conv(ITEMREC.ST_RETU, "")                                      '�W�����Ɂ@��
                    Call UniCode_Conv(ITEMREC.ST_REN, "")                                       '�W�����Ɂ@�A
                    Call UniCode_Conv(ITEMREC.ST_DAN, "")                                       '�W�����Ɂ@�i
                    Call UniCode_Conv(ITEMREC.BEF_SOKO, "")                                     '�O����Ɂ@�q��
                    Call UniCode_Conv(ITEMREC.BEF_RETU, "")                                     '�O����Ɂ@��
                    Call UniCode_Conv(ITEMREC.BEF_REN, "")                                      '�O����Ɂ@�A
                    Call UniCode_Conv(ITEMREC.BEF_DAN, "")                                      '�O����Ɂ@�i
                    Call UniCode_Conv(ITEMREC.LAST_NYU_DT, "")                                  '�ŏI���ɓ�
                    Call UniCode_Conv(ITEMREC.LAST_SYU_DT, "")                                  '�ŏI�o�ɓ�
                    Call UniCode_Conv(ITEMREC.HIN_NAI, "")                                      '�i�ԁi���j
                    Call UniCode_Conv(ITEMREC.BIKOU_SOKO, "")                                   'νđq��
                    Call UniCode_Conv(ITEMREC.BIKOU_TANA, "")                                   'νĒI��
                    Call UniCode_Conv(ITEMREC.HOJYU_P, "00000000")                              '��[�_
                    Call UniCode_Conv(ITEMREC.AVE_SYUKA, "00000000")                            '�����Ϗo�א�
                    Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "0")                                  '����ِ�
                    Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "0")                                  '����ِ�
                    Call UniCode_Conv(ITEMREC.LAST_INP_DT, "")                                  '�ŏI���ד��t
                    Call UniCode_Conv(ITEMREC.LAST_CHK_DT, "")                                  '�ŏI�ƍ����t
                    Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, "00000000")                         '�ƍ����݌ɐ�
                    Call UniCode_Conv(ITEMREC.BIKOU, "")                                        '������l
                    Call UniCode_Conv(ITEMREC.IRI_QTY, "")                                      '������萔
                    Call UniCode_Conv(ITEMREC.JAN_CODE, "")                                     'JAN����
                    Call UniCode_Conv(ITEMREC.HIN_CHANGE, "")                                   '�i�ԓǂݑւ�����
                    Call UniCode_Conv(ITEMREC.GOODS_KBN, "1")                                   '���i���L��
                    Call UniCode_Conv(ITEMREC.PACKING_NO, "")                                   '������
                    Call UniCode_Conv(ITEMREC.RANK, "")                                         '�����ݸ
                    Call UniCode_Conv(ITEMREC.NEW_RANK, "")                                     '�V�ݸ
                    Call UniCode_Conv(ITEMREC.GLICS1_TANA, "")                                  '��د���I��1
                    Call UniCode_Conv(ITEMREC.GLICS2_TANA, "")                                  '��د���I��2
                    Call UniCode_Conv(ITEMREC.GLICS3_TANA, "")                                  '��د���I��3
                
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_KBN, "")                                 '�Ɩ��Ǘ��@ �d���敪
                    Call UniCode_Conv(ITEMREC.G_HANBAI_KBN, "")                                 '           �̔��敪
                    Call UniCode_Conv(ITEMREC.G_SYUSHI, "")                                     '           ���x�P��
                    Call UniCode_Conv(ITEMREC.G_KUMITATE, "")                                   '           �g�����i
                    Call UniCode_Conv(ITEMREC.G_ST_URITAN, "")                                  '           �W���e�������P���@9(8)V99
                    Call UniCode_Conv(ITEMREC.G_ST_URITAN_DT, "")                               '           �W���e�������ݒ��
                    Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "")                                  '           �W���e�������P��  9(8)V99
                    Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, "")                               '           �W���e�������ݒ��
                    
                    For j = 0 To 2                                                              '�d������
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).CODE, "")                     '           �d����R�[�h
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).TANKA, "")                    '           �P��
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).TANKA_DT, "")                 '           �P���ݒ��
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LOT, "")                      '           �P���ݒ��
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LEAD_TIME, "")                '           ���[�h�^�C��
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LAST_ORDER_DT, "")            '           �ŏI������
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LAST_ORDER_QTY, "")           '           �ŏI������
                    
                    Next j
                
                    Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, "")                              '           �O���݌ɋ��z
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_KBN, "")                                 '           ���ދ敪
                    Call UniCode_Conv(ITEMREC.G_LABEL_NON, "")                                  '           ���ٓ\��t��
                    Call UniCode_Conv(ITEMREC.S_TANTO, "")                                      '���P�^�S����
                    Call UniCode_Conv(ITEMREC.ZAIKO_F, P_ZAIKO_F_ON)                            '�݌ɊǗ��Ώ�
                    
                    Call UniCode_Conv(ITEMREC.FILLER, "")                                       'Filler
        
                
                End If
                    
                Call UniCode_Conv(ITEMREC.L_HIN_NAME_E, CStr(ITEM_REC(2)))                  '�i���d
                Call UniCode_Conv(ITEMREC.L_BIKOU, CStr(ITEM_REC(5)))                       '���l
                    
                Call UniCode_Conv(ITEMREC.L_KAISHA_CODE, KAISHA_Tbl(i).C_Code)              '��Ж�
                
                Call UniCode_Conv(ITEMREC.L_KISHU1, CStr(ITEM_REC(7)))                      '�@��P
                Call UniCode_Conv(ITEMREC.L_KISHU2, CStr(ITEM_REC(9)))                      '�@��Q
                '--����ւ�
'                Call UniCode_Conv(ITEMREC.L_KISHU3, CStr(ITEM_REC(10)))                     '�@��R
                                
'                Call UniCode_Conv(ITEMREC.L_KISHU_BIKOU, CStr(ITEM_REC(18)))                '�K�p�@����l
                
                Call UniCode_Conv(ITEMREC.L_KISHU3, CStr(ITEM_REC(18)))                     '�@��R
                                
                Call UniCode_Conv(ITEMREC.L_KISHU_BIKOU, CStr(ITEM_REC(10)))                '�K�p�@����l
                '--����ւ�
                
                
                
                If Trim(CStr(ITEM_REC(11))) = "TRUE" Then                                   '��
                    Call UniCode_Conv(ITEMREC.L_PAPER, L_PAPER_ON)
                Else
                    Call UniCode_Conv(ITEMREC.L_PAPER, L_PAPER_OFF)
                End If
                
                If Trim(CStr(ITEM_REC(12))) = "TRUE" Then                                   '��׽���
                    Call UniCode_Conv(ITEMREC.L_PLASTIC, L_PLASTIC_ON)
                Else
                    Call UniCode_Conv(ITEMREC.L_PLASTIC, L_PLASTIC_OFF)
                End If
                If IsNumeric(ITEM_REC(13)) Then                                             '���i(1)
                    Call UniCode_Conv(ITEMREC.L_URIKIN1, Format(CDbl(ITEM_REC(13)), "0000000000"))
                Else
                    Call UniCode_Conv(ITEMREC.L_URIKIN1, "")
                End If
                If IsNumeric(ITEM_REC(14)) Then                                             '���i(2)
                    Call UniCode_Conv(ITEMREC.L_URIKIN2, Format(CDbl(ITEM_REC(14)), "0000000000"))
                Else
                    Call UniCode_Conv(ITEMREC.L_URIKIN2, "")
                End If
                If IsNumeric(ITEM_REC(15)) Then                                             '���i(3)
                    Call UniCode_Conv(ITEMREC.L_URIKIN3, Format(CDbl(ITEM_REC(15)), "0000000000"))
                Else
                    Call UniCode_Conv(ITEMREC.L_URIKIN3, "")
                End If
                    
                If Trim(CStr(ITEM_REC(16))) = "TRUE" Then                                   '�K�p�@������
                    Call UniCode_Conv(ITEMREC.L_LABEL, L_LABEL_ON)
                Else
                    Call UniCode_Conv(ITEMREC.L_LABEL, L_LABEL_OFF)
                End If
                
                If Trim(CStr(ITEM_REC(17))) = "TRUE" Then                                   '��������
                    Call UniCode_Conv(ITEMREC.L_MAISU, L_MAISU_ON)
                Else
                    Call UniCode_Conv(ITEMREC.L_MAISU, L_MAISU_OFF)
                End If
                    
                Call UniCode_Conv(ITEMREC.L_SAGYO_SHIJI, CStr(ITEM_REC(19)))                '��Ǝw��
                Call UniCode_Conv(ITEMREC.L_BIKOU3, CStr(ITEM_REC(20)))                     '���l�R
                    
                    
                Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, "")
                For i = 0 To UBound(KAISHA_Tbl)                                             '���ƕ�����
                
                
                    If KAISHA_Tbl(i).C_NAME = Trim(ITEM_REC(22)) Then
                        Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, KAISHA_Tbl(i).C_Code)
                        Exit For
                    End If
                
                Next i
                    
                If IsNumeric(ITEM_REC(23)) Then                                             '���萔
                    Call UniCode_Conv(ITEMREC.L_IRI_QTY, Format(CDbl(ITEM_REC(23)), "00000000"))
                Else
                    Call UniCode_Conv(ITEMREC.L_IRI_QTY, "")
                End If
            
            
                Call UniCode_Conv(ITEMREC.L_TANA1, CStr(ITEM_REC(24)))                      '�I��1
                Call UniCode_Conv(ITEMREC.L_TANA2, CStr(ITEM_REC(25)))                      '�I��2
                Call UniCode_Conv(ITEMREC.JAN_CODE, CStr(ITEM_REC(26)))                     'JAN����
            
            
                Do
                    sts = BTRV(Upd_com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
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
                            
                            Call File_Error(sts, Upd_com, "�i�ڃ}�X�^")
                            Exit Function
                    End Select
                Loop
            End If
        
        Loop
    '---------------------------------------------  �I��
    
        Cnt(1).Caption = Format(Count, "#0")
        
        Close #FileNo
    Case 2


'-----------------------------------------------------------------------------  ���i���|�|���o�n�r
                                    '���O�t�@�C������荞��
        If GetIni("FILE", "COMPO_TXT", "CONV2006", c) Then
            Beep
            MsgBox "[COMPO_TXT]�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            Unload Me
        End If
        FileNo = FreeFile
        
        
        fileName = RTrim(c)
        
            
        Open fileName For Input As FileNo
        
        
        
        
        MsgLab(1) = "���i���@���i�}�X�^�R���o�[�g�������I�I"
        Me.MousePointer = vbHourglass
        Count = 0
        DISP_INTERVAL = 0
        Cnt(2).Caption = Format(Count, "#0")
                                            
                                            
                                            
                                            
        Do Until EOF(FileNo)
            
            DoEvents
            
            Line Input #FileNo, RecordBuf
            
            ITEM_REC = Split(RecordBuf, vbTab, -1)
            
            
            Count = Count + 1
            DISP_INTERVAL = DISP_INTERVAL + 1
            If DISP_INTERVAL = 100 Then
                Cnt(2).Caption = Format(Count, "#0")
                DISP_INTERVAL = 0
            End If
            
            
            For i = 0 To UBound(ITEM_REC)
                For j = 0 To Len(ITEM_REC(i))
    
                    If Mid(CStr(ITEM_REC(i)), j + 1, 1) = """" Then
                        Mid(ITEM_REC(i), j + 1, 1) = " "
                    End If
    
                Next j
    
                ITEM_REC(i) = Trim(ITEM_REC(i))
            Next i
                                                                                                
            '�R�[�h�}�X�^�ǂݍ���
            Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN04_CD)
            Call UniCode_Conv(K0_P_CODE.C_Code, CStr(ITEM_REC(1)))
            
            Err_Flg = False
            
            
            sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            Select Case sts
                Case BtNoErr
                
                Case BtErrKeyNotFound
                    Err_Flg = True
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�R�[�h�}�X�^")
                    Exit Function
            End Select
            
            If Not Err_Flg Then
                        
            
            
            
            
                Call UniCode_Conv(K0_ITEM.JGYOBU, Trim(StrConv(P_CODEREC.OPTION1, vbUnicode)))      '���ƕ�(=����)
                Call UniCode_Conv(K0_ITEM.NAIGAI, Trim(StrConv(P_CODEREC.OPTION2, vbUnicode)))      '�����O(=����)
                Call UniCode_Conv(K0_ITEM.HIN_GAI, CStr(ITEM_REC(0)))   '�i��
                
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                    
                        Upd_com = BtOpUpdate
                    
                    
                    Case BtErrKeyNotFound
                        
                        Upd_com = BtOpInsert
                    
                    Case Else
                        
                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                        Exit Function
                End Select
                
                        
                If Upd_com = BtOpInsert Then
                        
                    Call UniCode_Conv(ITEMREC.JGYOBU, Trim(StrConv(P_CODEREC.OPTION1, vbUnicode)))       '���ƕ�=����
                    Call UniCode_Conv(ITEMREC.NAIGAI, Trim(StrConv(P_CODEREC.OPTION2, vbUnicode)))   '�����O
                    Call UniCode_Conv(ITEMREC.HIN_GAI, CStr(ITEM_REC(0)))       '�i�ں���
                    Call UniCode_Conv(ITEMREC.ST_SET_DT, "")                    '�W���I�Ԑݒ���t
                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")                      '�W�����Ɂ@�q��
                    Call UniCode_Conv(ITEMREC.ST_RETU, "")                      '�W�����Ɂ@��
                    Call UniCode_Conv(ITEMREC.ST_REN, "")                       '�W�����Ɂ@�A
                    Call UniCode_Conv(ITEMREC.ST_DAN, "")                       '�W�����Ɂ@�i
                    Call UniCode_Conv(ITEMREC.BEF_SOKO, "")                     '�O����Ɂ@�q��
                    Call UniCode_Conv(ITEMREC.BEF_RETU, "")                     '�O����Ɂ@��
                    Call UniCode_Conv(ITEMREC.BEF_REN, "")                      '�O����Ɂ@�A
                    Call UniCode_Conv(ITEMREC.BEF_DAN, "")                      '�O����Ɂ@�i
                    Call UniCode_Conv(ITEMREC.LAST_NYU_DT, "")                  '�ŏI���ɓ�
                    Call UniCode_Conv(ITEMREC.LAST_SYU_DT, "")                  '�ŏI�o�ɓ�
                    Call UniCode_Conv(ITEMREC.HIN_NAI, CStr(ITEM_REC(0)))       '�i�ԁi���j
                    Call UniCode_Conv(ITEMREC.BIKOU_SOKO, "")                   'νđq��
                    Call UniCode_Conv(ITEMREC.BIKOU_TANA, "")                   'νĒI��
                    Call UniCode_Conv(ITEMREC.HOJYU_P, "00000000")              '��[�_
                    Call UniCode_Conv(ITEMREC.AVE_SYUKA, "00000000")            '�����Ϗo�א�
                    Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "0")                  '����ِ�
                    Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "0")                  '����ِ�
                    Call UniCode_Conv(ITEMREC.LAST_INP_DT, "")                  '�ŏI���ד��t
                    Call UniCode_Conv(ITEMREC.LAST_CHK_DT, "")                  '�ŏI�ƍ����t
                    Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, "00000000")         '�ƍ����݌ɐ�
                    Call UniCode_Conv(ITEMREC.BIKOU, "")                        '������l
                    Call UniCode_Conv(ITEMREC.IRI_QTY, "")                      '������萔
                    Call UniCode_Conv(ITEMREC.JAN_CODE, "")                     'JAN����
                    Call UniCode_Conv(ITEMREC.HIN_CHANGE, "")                   '�i�ԓǂݑւ�����
                    Call UniCode_Conv(ITEMREC.GOODS_KBN, "0")                   '���i���L��
                    Call UniCode_Conv(ITEMREC.PACKING_NO, "")                   '������
                    Call UniCode_Conv(ITEMREC.RANK, "")                         '�����ݸ
                    Call UniCode_Conv(ITEMREC.NEW_RANK, "")                     '�V�ݸ
                    Call UniCode_Conv(ITEMREC.GLICS1_TANA, "")                  '��د���I��1
                    Call UniCode_Conv(ITEMREC.GLICS2_TANA, "")                  '��د���I��2
                    Call UniCode_Conv(ITEMREC.GLICS3_TANA, "")                  '��د���I��3
                
                    Call UniCode_Conv(ITEMREC.L_HIN_NAME_E, "")                 '�i��E
                    Call UniCode_Conv(ITEMREC.L_BIKOU, "")                      '���l
                    Call UniCode_Conv(ITEMREC.L_KAISHA_CODE, "")                '��Ж�
                    Call UniCode_Conv(ITEMREC.L_KISHU1, "")                     '�@��(1)
                    Call UniCode_Conv(ITEMREC.L_KISHU2, "")                     '�@��(2)
                    Call UniCode_Conv(ITEMREC.L_KISHU3, "")                     '�@��(3)
                    Call UniCode_Conv(ITEMREC.L_PAPER, "")                      '��
                    Call UniCode_Conv(ITEMREC.L_PLASTIC, "")                    '��׽���
                    Call UniCode_Conv(ITEMREC.L_URIKIN1, "")                    '���i(1)
                    Call UniCode_Conv(ITEMREC.L_URIKIN2, "")                    '���i(2)
                    Call UniCode_Conv(ITEMREC.L_URIKIN3, "")                    '���i(3)
                    Call UniCode_Conv(ITEMREC.L_LABEL, "")                      '�K�p�@������
                    Call UniCode_Conv(ITEMREC.L_MAISU, "")                      '���ٖ���
                    Call UniCode_Conv(ITEMREC.L_KISHU_BIKOU, "")                '�K�p�@����l
                    Call UniCode_Conv(ITEMREC.L_SAGYO_SHIJI, "")                '��Ǝw��
                    Call UniCode_Conv(ITEMREC.L_BIKOU3, "")                     '���l(3)
                    Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, "")                '���ƕ���
                    Call UniCode_Conv(ITEMREC.L_IRI_QTY, "")                    '���萔
                    Call UniCode_Conv(ITEMREC.L_TANA1, "")                      '�I��(1)
                    Call UniCode_Conv(ITEMREC.L_TANA2, "")                      '�I��(2)
                    
                    Call UniCode_Conv(ITEMREC.S_TANTO, "")                      '���P�^�S����
                    
                    Call UniCode_Conv(ITEMREC.ZAIKO_F, P_ZAIKO_F_ON)            '�݌ɊǗ��Ώ�
                    
                    Call UniCode_Conv(ITEMREC.HIN_NAME, "")                      '�i�ږ���
                        
                    Call UniCode_Conv(ITEMREC.ST_SET_DT, "")                                    '�W���I�Ԑݒ���t
                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")                                      '�W�����Ɂ@�q��
                    Call UniCode_Conv(ITEMREC.ST_RETU, "")                                      '�W�����Ɂ@��
                    Call UniCode_Conv(ITEMREC.ST_REN, "")                                       '�W�����Ɂ@�A
                    Call UniCode_Conv(ITEMREC.ST_DAN, "")                                       '�W�����Ɂ@�i
                    Call UniCode_Conv(ITEMREC.BEF_SOKO, "")                                     '�O����Ɂ@�q��
                    Call UniCode_Conv(ITEMREC.BEF_RETU, "")                                     '�O����Ɂ@��
                    Call UniCode_Conv(ITEMREC.BEF_REN, "")                                      '�O����Ɂ@�A
                    Call UniCode_Conv(ITEMREC.BEF_DAN, "")                                      '�O����Ɂ@�i
                    Call UniCode_Conv(ITEMREC.LAST_NYU_DT, "")                                  '�ŏI���ɓ�
                    Call UniCode_Conv(ITEMREC.LAST_SYU_DT, "")                                  '�ŏI�o�ɓ�
                    Call UniCode_Conv(ITEMREC.BIKOU_SOKO, "")                                   'νđq��
                    Call UniCode_Conv(ITEMREC.BIKOU_TANA, "")                                   'νĒI��
                    Call UniCode_Conv(ITEMREC.HOJYU_P, "00000000")                              '��[�_
                    Call UniCode_Conv(ITEMREC.AVE_SYUKA, "00000000")                            '�����Ϗo�א�
                    Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "0")                                  '����ِ�
                    Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "0")                                  '����ِ�
                    Call UniCode_Conv(ITEMREC.LAST_INP_DT, "")                                  '�ŏI���ד��t
                    Call UniCode_Conv(ITEMREC.LAST_CHK_DT, "")                                  '�ŏI�ƍ����t
                    Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, "00000000")                         '�ƍ����݌ɐ�
                    Call UniCode_Conv(ITEMREC.BIKOU, "")                                        '������l
                    Call UniCode_Conv(ITEMREC.IRI_QTY, "")                                      '������萔
                    Call UniCode_Conv(ITEMREC.JAN_CODE, "")                                     'JAN����
                    Call UniCode_Conv(ITEMREC.HIN_CHANGE, "")                                   '�i�ԓǂݑւ�����
                    Call UniCode_Conv(ITEMREC.GOODS_KBN, "1")                                   '���i���L��
                    Call UniCode_Conv(ITEMREC.PACKING_NO, "")                                   '������
                    Call UniCode_Conv(ITEMREC.RANK, "")                                         '�����ݸ
                    Call UniCode_Conv(ITEMREC.NEW_RANK, "")                                     '�V�ݸ
                    Call UniCode_Conv(ITEMREC.GLICS1_TANA, "")                                  '��د���I��1
                    Call UniCode_Conv(ITEMREC.GLICS2_TANA, "")                                  '��د���I��2
                    Call UniCode_Conv(ITEMREC.GLICS3_TANA, "")                                  '��د���I��3
                
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_KBN, "")                                 '�Ɩ��Ǘ��@ �d���敪
                    Call UniCode_Conv(ITEMREC.G_HANBAI_KBN, "")                                 '           �̔��敪
                    Call UniCode_Conv(ITEMREC.G_SYUSHI, "")                                     '           ���x�P��
                    Call UniCode_Conv(ITEMREC.G_KUMITATE, "")                                   '           �g�����i
                    Call UniCode_Conv(ITEMREC.G_ST_URITAN, "")                                  '           �W���e�������P���@9(8)V99
                    Call UniCode_Conv(ITEMREC.G_ST_URITAN_DT, "")                               '           �W���e�������ݒ��
                    Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "")                                  '           �W���e�������P��  9(8)V99
                    Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, "")                               '           �W���e�������ݒ��
                    
                    For j = 0 To 2                                                              '�d������
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).CODE, "")                     '           �d����R�[�h
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).TANKA, "")                    '           �P��
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).TANKA_DT, "")                 '           �P���ݒ��
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LOT, "")                      '           �P���ݒ��
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LEAD_TIME, "")                '           ���[�h�^�C��
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LAST_ORDER_DT, "")            '           �ŏI������
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LAST_ORDER_QTY, "")           '           �ŏI������
                    
                    Next j
                
                    Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, "")                              '           �O���݌ɋ��z
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_KBN, "")                                 '           ���ދ敪
                    Call UniCode_Conv(ITEMREC.G_LABEL_NON, "")                                  '           ���ٓ\��t��
                    Call UniCode_Conv(ITEMREC.S_TANTO, "")                                      '���P�^�S����
                    Call UniCode_Conv(ITEMREC.ZAIKO_F, P_ZAIKO_F_ON)                            '�݌ɊǗ��Ώ�
                    
                    Call UniCode_Conv(ITEMREC.FILLER, "")                                       'Filler
                    
                    
            
                
                
                    Do
                        sts = BTRV(Upd_com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
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
                                
                                Call File_Error(sts, Upd_com, "�i�ڃ}�X�^")
                                Exit Function
                        End Select
                    Loop
                
                
                End If
            End If
        
        Loop
    '---------------------------------------------  �I��
    
        Cnt(2).Caption = Format(Count, "#0")
        
        Close #FileNo
    
    End Select

    MsgBox "�R���o�[�g�I��"

End Function

Private Sub Command1_Click(index As Integer)

Dim ans     As Integer
Dim Mesg    As String

    Select Case index
    
        Case 0
            Mesg = "�u���ށv"
        
        Case 1
            Mesg = "�u���فv"
        Case 2
            Mesg = "�u���i�v"
        Case 3
            Unload Me
    End Select


    ans = MsgBox(Mesg & "���ްď��������s���܂����H", vbYesNo, "�m�F����")
    If ans = vbYes Then
    
        If Update_Proc(index) Then
            Unload Me
        End If
    
    End If
    
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
                                
                                '���ƕ���荞��
    If JGYOB_TB_Set Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
                                
                                
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                '�R�[�h�}�X�^�n�o�d�m
    If P_CODE_Open(BtOpenNomal) Then
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
    
                                            '�R�[�h�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�R�[�h�}�X�^")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set PC000101 = Nothing

    End
End Sub

