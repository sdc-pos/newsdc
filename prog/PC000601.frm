VERSION 5.00
Begin VB.Form PC000601 
   BackColor       =   &H00C0C0C0&
   Caption         =   "���ލ݌Ƀf�[�^�P���ݒ菈��"
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
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   7
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   4
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "�݌Ɉړ�����"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   6
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '�E����
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   5
      Top             =   3360
      Width           =   1455
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
      Caption         =   "�݌Ƀf�[�^��"
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
Attribute VB_Name = "PC000601"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Function Update_Proc() As Integer

Dim sts             As Integer
Dim com             As Integer
Dim ans             As Integer
Dim Count           As Long

Dim DISP_INTERVAL   As Long





Dim c               As String * 128

    Update_Proc = True

    
    
    
    
    
    
    MsgLab(1) = "�݌Ƀf�[�^�P���ݒ�@�������I�I"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(0).Caption = Format(Count, "#0")
                                        
                                        
    Call UniCode_Conv(K1_ZAIKO.JGYOBU, SHIZAI)
    Call UniCode_Conv(K1_ZAIKO.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K1_ZAIKO.HIN_GAI, "")
    Call UniCode_Conv(K1_ZAIKO.GOODS_ON, "")
    Call UniCode_Conv(K1_ZAIKO.NYUKA_DT, "")
    Call UniCode_Conv(K1_ZAIKO.Soko_No, "")
    Call UniCode_Conv(K1_ZAIKO.Retu, "")
    Call UniCode_Conv(K1_ZAIKO.Ren, "")
    Call UniCode_Conv(K1_ZAIKO.Dan, "")
                                        
                                        
                                        
    com = BtOpGetGreaterEqual
                                        
    Do
        
        DoEvents
        
        sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
        Select Case sts
            Case BtNoErr
            
            
                If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> SHIZAI Or _
                    StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> NAIGAI_NAI Then
                    Exit Do
                
                End If
            
            
            Case BtErrEOF
                Exit Do
            
            
            Case Else
                
                Call File_Error(sts, com, "�݌Ƀf�[�^")
                Exit Function
        End Select
        
        
        
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(ZAIKOREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(ZAIKOREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(ZAIKOREC.HIN_GAI, vbUnicode))
        
        
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            
                Call UniCode_Conv(ZAIKOREC.SHIIRE_CODE, StrConv(ITEMREC.G_SHIIRE_TBL(0).CODE, vbUnicode))
                Call UniCode_Conv(ZAIKOREC.SHIIRE_TANKA, StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode))
            
            
            
            
                sts = BTRV(BtOpUpdate, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
                Select Case sts
                    Case BtNoErr
                    
                    Case Else
                        
                        
                        
                        Call File_Error(sts, BtOpUpdate, "�݌��ް�")
                        Exit Function
                End Select
            
            
            
            
            
            Case BtErrKeyNotFound
            
            
            Case Else
                
                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                Exit Function
        End Select
        Count = Count + 1
        Cnt(0).Caption = Format(Count, "#0")
        
        com = BtOpGetNext
        
        
        
    
    Loop
'---------------------------------------------  �I��

    Cnt(0).Caption = Format(Count, "#0")
    
    Update_Proc = False
End Function

Private Function IDO_Update_Proc() As Integer

Dim sts             As Integer
Dim com             As Integer
Dim ans             As Integer
Dim Count           As Long

Dim DISP_INTERVAL   As Long



Dim c               As String * 128

    IDO_Update_Proc = True

    
    
    
    
    
    
    MsgLab(1) = "�݌Ɉړ���P���ݒ�@�������I�I"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(1).Caption = Format(Count, "#0")
                                        
                                        
    Call UniCode_Conv(K0_IDO.JGYOBU, SHIZAI)
    Call UniCode_Conv(K0_IDO.JITU_DT, "")
    Call UniCode_Conv(K0_IDO.JITU_TM, "")
                                        
                                        
                                        
    com = BtOpGetGreaterEqual
                                        
    Do
        
        DoEvents
        
        sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
        Select Case sts
            Case BtNoErr
            
            
                If StrConv(IDOREC.JGYOBU, vbUnicode) <> SHIZAI Then
                    Exit Do
                
                End If
            
            
            Case BtErrEOF
                Exit Do
            
            
            Case Else
                
                Call File_Error(sts, com, "�݌Ɉړ���")
                Exit Function
        End Select
        
        
        
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(IDOREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(IDOREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(IDOREC.HIN_GAI, vbUnicode))
        
        
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            
                Call UniCode_Conv(IDOREC.SHIIRE_CODE, StrConv(ITEMREC.G_SHIIRE_TBL(0).CODE, vbUnicode))
                Call UniCode_Conv(IDOREC.SHIIRE_TANKA, StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode))
            
            
            
            
                sts = BTRV(BtOpUpdate, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
                Select Case sts
                    Case BtNoErr
                    
                    Case Else
                        
                        
                        
                        Call File_Error(sts, BtOpUpdate, "�݌Ɉړ���")
                        Exit Function
                End Select
            
            
            
            
            
            Case BtErrKeyNotFound
            
            
            Case Else
                
                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                Exit Function
        End Select
        Count = Count + 1
        Cnt(1).Caption = Format(Count, "#0")
        
        com = BtOpGetNext
        
        
        
    
    Loop
'---------------------------------------------  �I��

    Cnt(1).Caption = Format(Count, "#0")
    
    IDO_Update_Proc = False

End Function


Private Sub Command1_Click(Index As Integer)

Dim ans As Integer
                                
                                    
                                
    Select Case Index
                                '�����I��
        
        Case 0
            Beep
            ans = MsgBox("�u�݌��ް��v ���s���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                If Update_Proc() Then
                    Unload Me
                End If
            End If
        Case 1
            Beep
            ans = MsgBox("�u�݌Ɉړ����v ���s���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                If IDO_Update_Proc() Then
                    Unload Me
                End If
            End If
    
    End Select
    
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
                                '�݌��ް��n�o�d�m
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '�݌Ɉړ����ް��n�o�d�m
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '�݌��ް�CLOSE
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌��ް�")
        End If
    End If
                                            
                                            '�݌Ɉړ����ް�CLOSE
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ɉړ���")
        End If
    End If
                                            
                                            '�i��Ͻ�CLOSE
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i��Ͻ�")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set PC000601 = Nothing

    End
End Sub

