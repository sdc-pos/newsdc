VERSION 5.00
Begin VB.Form F9000701 
   Caption         =   "      "
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   7065
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text 
      Alignment       =   2  '��������
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   3960
      MaxLength       =   2
      TabIndex        =   13
      Top             =   720
      Width           =   492
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�I�@�@��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   1080
      Width           =   1692
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���@�@�s"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   1692
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Caption         =   "�ݒ��q��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   6
      Left            =   2640
      TabIndex        =   12
      Top             =   840
      Width           =   1332
   End
   Begin VB.Label SokoName 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4560
      TabIndex        =   11
      Top             =   840
      Width           =   1692
   End
   Begin VB.Label lblIn_CNT 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   2
      Left            =   3240
      TabIndex        =   8
      Top             =   3000
      Width           =   1092
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   5
      Left            =   4320
      TabIndex        =   10
      Top             =   3000
      Width           =   372
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Caption         =   "�݌ɓo�^������"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   1440
      TabIndex        =   9
      Top             =   3000
      Width           =   1692
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   3
      Left            =   4320
      TabIndex        =   7
      Top             =   2520
      Width           =   372
   End
   Begin VB.Label lblIn_CNT 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   3240
      TabIndex        =   6
      Top             =   2520
      Width           =   1092
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Caption         =   "�i�ړo�^������"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   1440
      TabIndex        =   5
      Top             =   2520
      Width           =   1692
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   4320
      TabIndex        =   4
      Top             =   2040
      Width           =   372
   End
   Begin VB.Label lblIn_CNT 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   3240
      TabIndex        =   3
      Top             =   2040
      Width           =   1092
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Caption         =   "�捞�݌����@��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Top             =   2040
      Width           =   1692
   End
End
Attribute VB_Name = "F9000701"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WS_NO As String * 2                 'ܰ��ð��ݔԍ�

Private Sub Command1_Click(Index As Integer)

Dim ans     As Integer
Dim sts     As Integer

    Select Case Index
        Case 0
            
            Call UniCode_Conv(K0_SOKO.Soko_No, Text(0).Text)
            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
            Select Case sts
                Case BtNoErr
                    SokoName.Caption = StrConv(SOKOREC.SOKO_NAME, vbUnicode)
                Case BtErrKeyNotFound
                    SokoName.Caption = ""
                    Beep
                    MsgBox "���͂������ڂ́A�G���[�ł��B"
                    Text(0).SetFocus
                    Exit Sub
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^")
                    Unload Me
            End Select
            
            
            Beep
            ans = MsgBox("�݌ɈڊǏ��������s���܂����H", vbYesNo, "�m�F")
            If ans = vbNo Then
                Text(0).SetFocus
                Exit Sub
            End If
            
            
            Call Data_Update_Proc
        Case 1
            Unload Me
    End Select

End Sub

Private Sub Form_Load()

Dim c As String * 128
Dim sts As Integer
Dim sBuffer As String * 255
Dim com     As String

                                
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

                                'ܰ��ð��ݔԍ���荞��
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "??"
    End If
    WS_NO = RTrim(com)
                                '���ƕ���荞��
    If JGYOB_TB_Set Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If


                                '�q�Ƀ}�X�^�n�o�d�m
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�v���}�X�^�n�o�d�m
    If YOIN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�S���҃}�X�^�n�o�d�m
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���ח\��n�o�d�m
    If Y_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ƀf�[�^�n�o�d�m
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ɉړ����n�o�d�m
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i�ڃ}�X�^�i�X�V�p���[�N�j�n�o�d�m
    If wITEM_Open(BtOpenNomal) Then
        Unload Me
    End If

    Text(0).Text = "90"
    
    Show
    Text(0).SetFocus

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '�q�Ƀ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�q�Ƀ}�X�^")
        End If
    End If
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '���ח\��b�k�n�r�d
    sts = BTRV(BtOpClose, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���ח\��")
        End If
    End If
                                            '�݌Ƀf�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
        End If
    End If
                                            '�݌Ɉړ����b�k�n�r�d
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ɉړ���")
        End If
    End If
                                            '�i�ڃ}�X�^�i�X�V�p���[�N�j�b�k�n�r�d
    sts = BTRV(BtOpClose, wITEM_POS, wITEMREC, Len(wITEMREC), K0_wITEM, Len(K0_wITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '�v���}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�v���}�X�^")
        End If
    End If
                                            '�S���҃}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�S���҃}�X�^")
        End If
    End If
                                            '�a���������������Z�b�g
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set F9000701 = Nothing
    
    
    End
End Sub


Private Sub Data_Update_Proc()
 
Dim Fno         As Integer
Dim ZaikoTemp   As String
Dim IN_CNT      As Integer
Dim sts         As Integer
Dim Zaiko_CNT   As Integer
Dim Item_CNT   As Integer



    If OutREC_Open_Proc() Then
        Unload Me
    End If
    
    IN_CNT = 0
    DataNo = 0
    
    Fno = FreeFile
    On Error Resume Next
    Open "c:\zaiko\IN_FILE.CSV" For Input As #Fno
    '�݌Ƀf�[�^�i�b�r�u�j�ǂݍ���
    Do While EOF(Fno) = False
        DoEvents
        
        Line Input #Fno, ZaikoTemp
        ZaikoData = Split(ZaikoTemp, ",", True, vbTextCompare)
        IN_CNT = IN_CNT + 1
    
            
    
        lblIn_CNT(0) = Format(IN_CNT, "#0")
    
        If Data_Put_Proc() Then
            Unload Me
        End If
    
    Loop

    Close #OutFno
    Close #Fno
    Fno = FreeFile
    Open "c:\zaiko\shiji79.dat" For Binary As #Fno
    Zaiko_CNT = 0
    Item_CNT = 0

    Do
                                    
        DoEvents
                                    
                                    '�w���f�[�^�ǂݍ���
        Get #Fno, , OutREC
        If Left(StrConv(OutREC.TEXT_NO, vbUnicode), 1) < " " Then    'EOF ?
            Exit Do
        End If
                                        '�g�����U�N�V�����J�n
        sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        If sts <> BtNoErr Then
            Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
            Unload Me
        End If
        
        
        If Upd_Item(Item_CNT) Then
            GoTo Abort_Tran
        End If
                                        
        If NyukaY_Put(Zaiko_CNT) Then
            GoTo Abort_Tran
        End If
                                        
                                        

                                        '�g�����U�N�V�����I��
        sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        If sts <> BtNoErr Then
            Call File_Error(sts, BtOpEndTransaction, "")
            GoTo Abort_Tran
        End If
    Loop


    Close #Fno

    MsgBox "�݌ɈڊǏ������I�����܂����B"
    Unload Me

Abort_Tran:
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If

    Unload Me

End Sub
                                            '�i�ڃ}�X�^�X�V
Private Function Upd_Item(IN_CNT As Integer) As Boolean
Dim sts As Integer
Dim ans As Integer
Dim Command As Integer
Dim Work As String


    Upd_Item = True


    Call UniCode_Conv(K0_ITEM.JGYOBU, "7")
    Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(OutREC.HIN_GAI, vbUnicode))
    
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Command = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                Command = BtOpInsert
                Call UniCode_Conv(ITEMREC.JGYOBU, "7")
                Call UniCode_Conv(ITEMREC.NAIGAI, "1")
                Call UniCode_Conv(ITEMREC.HIN_GAI, StrConv(OutREC.HIN_GAI, vbUnicode))
                Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                Call UniCode_Conv(ITEMREC.ST_SET_DT, "")
                Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                Call UniCode_Conv(ITEMREC.ST_RETU, "")
                Call UniCode_Conv(ITEMREC.ST_REN, "")
                Call UniCode_Conv(ITEMREC.ST_DAN, "")
                Call UniCode_Conv(ITEMREC.BEF_SOKO, "")
                Call UniCode_Conv(ITEMREC.BEF_RETU, "")
                Call UniCode_Conv(ITEMREC.BEF_REN, "")
                Call UniCode_Conv(ITEMREC.BEF_DAN, "")
                Call UniCode_Conv(ITEMREC.LAST_NYU_DT, "")
                Call UniCode_Conv(ITEMREC.LAST_SYU_DT, "")
                Call UniCode_Conv(ITEMREC.HIN_NAI, StrConv(OutREC.HIN_NAI, vbUnicode))
                Call UniCode_Conv(ITEMREC.BIKOU_SOKO, "")
                Call UniCode_Conv(ITEMREC.BIKOU_TANA, "")
                Call UniCode_Conv(ITEMREC.SIZAI_CD, "")
                Call UniCode_Conv(ITEMREC.HOJYU_P, "00000000")
                Call UniCode_Conv(ITEMREC.AVE_SYUKA, "00000000")
                Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "1")
                Call UniCode_Conv(ITEMREC.LAST_INP_DT, "")
                
                Call UniCode_Conv(ITEMREC.LOCK_F, "0")          '�r���t���O
                Call UniCode_Conv(ITEMREC.WEL_ID, "")           '�g�p���q�@�h�c
                Call UniCode_Conv(ITEMREC.PRG_ID, "")           '�g�p���v���O����
                
                Call UniCode_Conv(ITEMREC.LAST_CHK_DT, "")
                Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, "0000000")
                Call UniCode_Conv(ITEMREC.MOTO_JIGYOBU, "")
                Call UniCode_Conv(ITEMREC.BIKOU, "")
                Call UniCode_Conv(ITEMREC.IRI_QTY, "")
                
                Call UniCode_Conv(ITEMREC.JAN_CODE, "")
                Call UniCode_Conv(ITEMREC.HIN_CHANGE, "")
                Call UniCode_Conv(ITEMREC.GOODS_KBN, GOODS_ON)
                Call UniCode_Conv(ITEMREC.PACKING_NO, "")
                Call UniCode_Conv(ITEMREC.RANK, "")
                
                
                Call UniCode_Conv(ITEMREC.FILLER, "")
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^")
                Exit Function
        End Select
    Loop
    
    Call UniCode_Conv(ITEMREC.LAST_NYU_DT, Format(Now, "YYYYMMDD"))
    Call UniCode_Conv(ITEMREC.LAST_INP_DT, Format(Now, "YYYYMMDD"))
    Call UniCode_Conv(ITEMREC.HIN_NAI, StrConv(OutREC.HIN_NAI, vbUnicode))
    Call UniCode_Conv(ITEMREC.HIN_NAME, StrConv(OutREC.HIN_NAME, vbUnicode))
    
    IN_CNT = IN_CNT + 1
    lblIn_CNT(1) = Format(IN_CNT, "#0")
    
    Do
        sts = BTRV(Command, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr, BtErrEOF, BtErrKeyNotFound
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, Command, "�i�ڃ}�X�^")
                Exit Function
        End Select
    Loop

    Upd_Item = False

End Function

                                            '���ח\��쐬 �� ���׍X�V
Private Function NyukaY_Put(IN_CNT As Integer) As Boolean

Dim sts     As Integer
Dim Work    As String * 8
Dim ans     As Integer

    NyukaY_Put = True
'�݌ɐ����O�͑ΏۊO
    If CLng(StrConv(OutREC.YOTEI_QTY, vbUnicode)) = 0 Then
            NyukaY_Put = False
            Exit Function
    End If

'���ח\��쐬
                                '�����敪
'    Call UniCode_Conv(Y_NYUREC.KAN_KBN, KAN_KBN_ON)
'                                '�f�[�^���
'    Call UniCode_Conv(Y_NYUREC.DT_SYU, "0")
'                                '�\�萔��
'    Call UniCode_Conv(Y_NYUREC.YOTEI_QTY, Format(CLng(StrConv(OutREC.YOTEI_QTY, vbUnicode)), "00000000"))
'                                '�m�萔��
'    Call UniCode_Conv(Y_NYUREC.FIX_QTY, "00000000")
'                                '�����O
'    Call UniCode_Conv(Y_NYUREC.NAIGAI, "1")
'                                '���ƕ��敪
'    Call UniCode_Conv(Y_NYUREC.JGYOBU, StrConv(OutREC.JGYOBU, vbUnicode))
'                                '�����敪
'    Call UniCode_Conv(Y_NYUREC.CYOK_KBN, StrConv(OutREC.CYOK_KBN, vbUnicode))
'                                '�e�L�X�g��
'    Call UniCode_Conv(Y_NYUREC.TEXT_NO, StrConv(OutREC.TEXT_NO, vbUnicode))
'                                '�`�[���t
'    Call UniCode_Conv(Y_NYUREC.DEN_DT, StrConv(OutREC.DEN_DT, vbUnicode))
'                                '���o�ɋ敪
'    Call UniCode_Conv(Y_NYUREC.IO_KBN, StrConv(OutREC.IO_KBN, vbUnicode))
'                                '�ԍ��敪
'    Call UniCode_Conv(Y_NYUREC.PM_KBN, StrConv(OutREC.PM_KBN, vbUnicode))
'                                '�`�[���
'    Call UniCode_Conv(Y_NYUREC.DEN_SYU, StrConv(OutREC.DEN_SYU, vbUnicode))
'                                '�`�[��
'    Call UniCode_Conv(Y_NYUREC.DEN_NO, StrConv(OutREC.DEN_NO, vbUnicode))
'                                '�����敪
'    Call UniCode_Conv(Y_NYUREC.CYU_KBN, StrConv(OutREC.CYU_KBN, vbUnicode))
'                                '�i�ԁi�O���j
'    Call UniCode_Conv(Y_NYUREC.HIN_GAI, StrConv(OutREC.HIN_GAI, vbUnicode))
'                                '�i�ԁi�����j
'    Call UniCode_Conv(Y_NYUREC.HIN_NAI, StrConv(OutREC.HIN_NAI, vbUnicode))
'                                '�i��
'    Call UniCode_Conv(Y_NYUREC.HIN_NAME, StrConv(OutREC.HIN_NAME, vbUnicode))
'                                '�\�Z�P�ʁi���j
'    Call UniCode_Conv(Y_NYUREC.YOSAN_FROM, StrConv(OutREC.YOSAN_FROM, vbUnicode))
'                                '�\�Z�P�ʁi��j
'    Call UniCode_Conv(Y_NYUREC.YOSAN_TO, StrConv(OutREC.YOSAN_TO, vbUnicode))
'                                '�q�ɋ敪�iνāj
'    Call UniCode_Conv(Y_NYUREC.HOST_SOKO, StrConv(OutREC.HOST_SOKO, vbUnicode))
'                                '�I�ԁiνāj���@�W�����ɒI�ԁi�i��Ͻ��j
'    Work = StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
'           StrConv(ITEMREC.ST_RETU, vbUnicode) & _
'           StrConv(ITEMREC.ST_REN, vbUnicode) & _
'           StrConv(ITEMREC.ST_DAN, vbUnicode)
'    Call UniCode_Conv(Y_NYUREC.HOST_TANA, Work)
'                                '�x����^�o�א�
'    Call UniCode_Conv(Y_NYUREC.SYUK_CODE, StrConv(OutREC.SYUK_CODE, vbUnicode))
'                                '�x����^�o�א於
'    Call UniCode_Conv(Y_NYUREC.SYUK_NAME, StrConv(OutREC.SYUK_NAME, vbUnicode))
'                                '��s���א�
'    Call UniCode_Conv(Y_NYUREC.BEF_NYU_QTY, "00000000")
'                                '�������t
'    Call UniCode_Conv(Y_NYUREC.KAN_DT, Format(Now, "YYYYMMDD"))
'                                'FILLER
'    Call UniCode_Conv(Y_NYUREC.FILLER, "")
'
'���ח\��f�[�^�ǉ��i���ו��j
'    Do
'        sts = BTRV(BtOpInsert, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), ZERO)
'        Select Case sts
'            Case BtNoErr
'                Exit Do
'            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                Beep
'                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'                If ans = vbCancel Then
'                    Exit Function
'                End If
'            Case BtErrDuplicates
'                Exit Do
'            Case Else
'                Call File_Error(sts, BtOpInsert, "���ח\��")
'                Exit Function
'        End Select
'    Loop

'���א��ō݌Ƀf�[�^�X�V�i�{�j
    If Nyuko_Update_Proc(StrConv(OutREC.JGYOBU, vbUnicode), _
                            "1", _
                            StrConv(OutREC.HIN_GAI, vbUnicode), _
                            StrConv(OutREC.DEN_DT, vbUnicode), _
                            (Trim(Text(0).Text) & "01" & "01" & "01"), _
                            "10", _
                            0, _
                            CLng(StrConv(OutREC.YOTEI_QTY, vbUnicode)), _
                            WS_NO, _
                            WS_NO, _
                            , _
                            "�݌Ɉڊ�") Then
        Exit Function
    
    End If

    IN_CNT = IN_CNT + 1
    lblIn_CNT(2) = Format(IN_CNT, "#0")
    
    NyukaY_Put = False

End Function



   
Private Sub Text_GotFocus(Index As Integer)
    If Text(Index).TabStop = True Then
        Text(Index) = Trim(Text(Index).Text)
        Text(Index).SelStart = 0
        Text(Index).SelLength = Len(Text(Index).Text)
    End If

End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim i   As Integer
Dim sts As Integer
    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
    Select Case Index
        Case 0
            Call UniCode_Conv(K0_SOKO.Soko_No, Text(0).Text)
            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
            Select Case sts
                Case BtNoErr
                    SokoName.Caption = StrConv(SOKOREC.SOKO_NAME, vbUnicode)
                Case BtErrKeyNotFound
                    SokoName.Caption = ""
                    Beep
                    MsgBox "���͂������ڂ́A�G���[�ł��B"
                    Text(0).SetFocus
                    Exit Sub
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^")
                    Unload Me
            End Select
    End Select
        
    Command1(0).SetFocus

End Sub
