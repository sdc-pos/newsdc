VERSION 5.00
Begin VB.Form CONV_ZAIKO_201210091 
   BackColor       =   &H00C0C0C0&
   Caption         =   "�f�[�^�R���o�[�g�����iCONV_ZAIKO_20121009 2012.10.10 08:45)"
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
   Begin VB.CheckBox Check1 
      Caption         =   "���O�o��"
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�I��"
      Height          =   435
      Index           =   1
      Left            =   7290
      TabIndex        =   10
      Top             =   1680
      Width           =   1860
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�J�n"
      Height          =   435
      Index           =   0
      Left            =   7290
      TabIndex        =   9
      Top             =   1080
      Width           =   1860
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   5940
      TabIndex        =   8
      Top             =   4800
      Width           =   1410
   End
   Begin VB.Label Label2 
      Caption         =   "��������"
      Height          =   315
      Index           =   2
      Left            =   5940
      TabIndex        =   7
      Top             =   4440
      Width           =   1410
   End
   Begin VB.Label Label2 
      Caption         =   "�Ώی���"
      Height          =   315
      Index           =   1
      Left            =   4455
      TabIndex        =   6
      Top             =   4440
      Width           =   1410
   End
   Begin VB.Label Label2 
      Caption         =   "�Ǎ��݌���"
      Height          =   315
      Index           =   0
      Left            =   2925
      TabIndex        =   5
      Top             =   4440
      Width           =   1410
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   2925
      TabIndex        =   4
      Top             =   4800
      Width           =   1410
   End
   Begin VB.Label Label1 
      Caption         =   "�i�ڃ}�X�^"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   1350
      TabIndex        =   3
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   4455
      TabIndex        =   2
      Top             =   4800
      Width           =   1410
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
      Caption         =   "�f�[�^���J�o���[����"
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
      Height          =   495
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   4800
   End
End
Attribute VB_Name = "CONV_ZAIKO_201210091"
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
Dim sel_count       As Long
Dim upd_count       As Long

Dim DISP_INTERVAL   As Long

Dim wkKISHU1        As String * 25
Dim wkKISHU2        As String * 52
Dim wkKISHU3        As String * 150
Dim wkKISHU_BIKOU   As String * 450

Dim c               As String * 128

Dim i               As Integer

Dim CHG_FLG         As Boolean

Dim Start_Now       As String


Dim wkSoko_No       As String * 2   '�q�ɇ�
Dim wkRetu          As String * 2   '�I�ԁ@��
Dim wkRen           As String * 2   '�I�ԁ@�A
Dim wkDan           As String * 2   '�I�ԁ@�i
Dim wkJGYOBU        As String * 1   '���ƕ��敪
Dim wkNAIGAI        As String * 1   '�����O
Dim wkHIN_GAI       As String * 20  '�i�ԁi�O���j
Dim wkGOODS_ON      As String * 1   '���i���^�����i��
Dim wkNYUKA_DT      As String * 8   '���ד��t

Dim wkNYUKO_DT      As String * 8   '���ɓ��t


Dim wkYUKO_Z_QTY    As String * 8   '�L���݌ɐ�
Dim YUKO_Z_QTY      As Long
    Update_Proc = True


'---------------------------------------------  ��������f�[�^�̃R���o�[�g
    MsgLab(1) = "�݌Ƀf�[�^���J�o���[�������I�I"
    Me.MousePointer = vbHourglass
    Count = 0
    sel_count = 0
    upd_count = 0
    DISP_INTERVAL = 0
    Cnt(0).Caption = Format(Count, "#0")
    Cnt(1).Caption = Format(Count, "#0")
    Cnt(2).Caption = Format(Count, "#0")
                                        
    Start_Now = Format(Now, "YYYYMMDDHHMMSS")
                                        
                                        
    com = BtOpGetFirst
    Do
        
        DoEvents
        
        
        sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�݌Ƀf�[�^")
                Exit Function
        End Select
        
        
        Count = Count + 1
        Cnt(0).Caption = Format(Count, "#0")
        
        
        
        If Left(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 4) = "9999" Then
            
            
            sel_count = sel_count + 1
            Cnt(1).Caption = Format(sel_count, "#0")
            upd_count = upd_count + 1
            Cnt(2).Caption = Format(upd_count, "#0")
            
            
            
                        
            wkSoko_No = StrConv(ZAIKOREC.Soko_No, vbUnicode)
            wkRetu = StrConv(ZAIKOREC.Retu, vbUnicode)
            wkRen = StrConv(ZAIKOREC.Ren, vbUnicode)
            wkDan = StrConv(ZAIKOREC.Dan, vbUnicode)
            wkJGYOBU = StrConv(ZAIKOREC.JGYOBU, vbUnicode)
            wkNAIGAI = StrConv(ZAIKOREC.NAIGAI, vbUnicode)
            wkHIN_GAI = StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
            wkGOODS_ON = StrConv(ZAIKOREC.GOODS_ON, vbUnicode)
            wkNYUKA_DT = StrConv(ZAIKOREC.NYUKA_DT, vbUnicode)
            wkNYUKO_DT = StrConv(ZAIKOREC.NYUKO_DT, vbUnicode)

            wkYUKO_Z_QTY = StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)

            
            
            sts = BTRV(BtOpDelete, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
            Select Case sts
                Case BtNoErr
                Case Else
                    Call File_Error(sts, BtOpDelete, "�݌Ƀf�[�^")
                    Exit Function
            End Select
            
            If Check1.Value = vbChecked Then
                Call LOG_OUT(LOG_F, "[DELETE][" & wkSoko_No & wkRetu & wkRen & wkDan & "][" & wkJGYOBU & "][" & wkNAIGAI & "][" & wkHIN_GAI & "][" & wkGOODS_ON & "][" & wkNYUKA_DT & "]" & wkYUKO_Z_QTY)
            End If
            
            
            Call UniCode_Conv(ZAIKOREC.NYUKA_DT, Left(wkNYUKO_DT, 4) & Right(wkNYUKA_DT, 4))
            sts = BTRV(BtOpInsert, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
            Select Case sts
                Case BtNoErr
                    If Check1.Value = vbChecked Then
                        Call LOG_OUT(LOG_F, "[INSERT][" & wkSoko_No & wkRetu & wkRen & wkDan & "][" & wkJGYOBU & "][" & wkNAIGAI & "][" & wkHIN_GAI & "][" & wkGOODS_ON & "][" & Left(wkNYUKO_DT, 4) & Right(wkNYUKA_DT, 4) & "]" & wkYUKO_Z_QTY)
                    End If
                
                Case BtErrDuplicates
                
                
                
                    Call UniCode_Conv(K0_ZAIKO.Soko_No, wkSoko_No)
                    Call UniCode_Conv(K0_ZAIKO.Retu, wkRetu)
                    Call UniCode_Conv(K0_ZAIKO.Ren, wkRen)
                    Call UniCode_Conv(K0_ZAIKO.Dan, wkDan)
                    Call UniCode_Conv(K0_ZAIKO.JGYOBU, wkJGYOBU)
                    Call UniCode_Conv(K0_ZAIKO.NAIGAI, wkNAIGAI)
                    Call UniCode_Conv(K0_ZAIKO.HIN_GAI, wkHIN_GAI)
                    Call UniCode_Conv(K0_ZAIKO.GOODS_ON, wkGOODS_ON)
            
                    Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, Left(wkNYUKO_DT, 4) & Right(wkNYUKA_DT, 4))
                        
                    sts = BTRV(BtOpGetEqual, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
                    Select Case sts
                        Case BtNoErr
                        
                    If Check1.Value = vbChecked Then
                        Call LOG_OUT(LOG_F, "[UPDATE][" & wkSoko_No & wkRetu & wkRen & wkDan & "][" & wkJGYOBU & "][" & wkNAIGAI & "][" & wkHIN_GAI & "][" & wkGOODS_ON & "][" & Left(wkNYUKO_DT, 4) & Right(wkNYUKA_DT, 4) & "]" & StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode) & "+" & wkYUKO_Z_QTY)
                    End If
                        
                            wkYUKO_Z_QTY = Val(wkYUKO_Z_QTY) + Val(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                            Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, Format(wkYUKO_Z_QTY, "00000000"))
                        
                            sts = BTRV(BtOpUpdate, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
                            Select Case sts
                                Case BtNoErr
                                Case Else
                                    Call File_Error(sts, BtOpDelete, "�݌Ƀf�[�^")
                                    Exit Function
                            End Select
                        
                        
                        Case BtErrKeyNotFound
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "�݌Ƀf�[�^")
                            Exit Function
                    End Select
                
                                    
                
                
                Case Else
                    Call File_Error(sts, BtOpInsert, "�݌Ƀf�[�^")
                    Exit Function
            End Select
            
            
            Call UniCode_Conv(K0_ZAIKO.Soko_No, wkSoko_No)
            Call UniCode_Conv(K0_ZAIKO.Retu, wkRetu)
            Call UniCode_Conv(K0_ZAIKO.Ren, wkRen)
            Call UniCode_Conv(K0_ZAIKO.Dan, wkDan)
            Call UniCode_Conv(K0_ZAIKO.JGYOBU, wkJGYOBU)
            Call UniCode_Conv(K0_ZAIKO.NAIGAI, wkNAIGAI)
            Call UniCode_Conv(K0_ZAIKO.HIN_GAI, wkHIN_GAI)
            Call UniCode_Conv(K0_ZAIKO.GOODS_ON, wkGOODS_ON)
            Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, wkNYUKA_DT)
            
            
            com = BtOpGetGreaterEqual
        Else
            com = BtOpGetNext
        End If
    Loop

    Cnt(0).Caption = Format(Count, "#0")


'---------------------------------------------  �I��
Update_End:
    
    Update_Proc = False

End Function

Private Sub Command1_Click(Index As Integer)
    
Dim ans As Integer
    
    Select Case Index
        Case 0
            ans = MsgBox("���s���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                If Update_Proc() Then
                    Unload Me
                End If
            
                MsgBox "�I�����܂���"
                Unload Me
            
            End If


        Case 1
            Unload Me
    End Select

End Sub

Private Sub Form_Activate()

Dim ans As Integer
                                
                                
                                
                                

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
    
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                    
    Check1.Value = vbChecked
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌��ް�")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set CONV_ZAIKO_201210091 = Nothing

    End
End Sub
