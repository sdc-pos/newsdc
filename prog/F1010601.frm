VERSION 5.00
Begin VB.Form F1010601 
   BackColor       =   &H00FFFFFF&
   Caption         =   "������Ǘ��}�X�^�����e�i���X"
   ClientHeight    =   11325
   ClientLeft      =   2130
   ClientTop       =   2430
   ClientWidth     =   16875
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
   MaxButton       =   0   'False
   ScaleHeight     =   11325
   ScaleWidth      =   16875
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   6
      Left            =   7875
      MaxLength       =   2
      TabIndex        =   28
      Top             =   1440
      Width           =   390
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   5
      Left            =   5520
      MaxLength       =   3
      TabIndex        =   6
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   6840
      MaxLength       =   8
      TabIndex        =   3
      Top             =   960
      Width           =   1092
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   4  '�S�p�Ђ炪��
      Index           =   3
      Left            =   8040
      MaxLength       =   20
      TabIndex        =   4
      Top             =   960
      Width           =   4935
   End
   Begin VB.ListBox List1 
      Height          =   7980
      ItemData        =   "F1010601.frx":0000
      Left            =   600
      List            =   "F1010601.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   2280
      Width           =   12375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   4  '�S�p�Ђ炪��
      Index           =   4
      Left            =   3000
      MaxLength       =   10
      TabIndex        =   5
      Top             =   1440
      Width           =   1332
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   4  '�S�p�Ђ炪��
      Index           =   1
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   2
      Top             =   960
      Width           =   4935
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   600
      MaxLength       =   8
      TabIndex        =   1
      Top             =   960
      Width           =   1092
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   600
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command 
      Caption         =   "�I  ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   10320
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   10560
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   9480
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   10560
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   8640
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   10560
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�f�[�^"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   7800
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   10560
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   6480
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   10560
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   5640
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   10560
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4800
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   10560
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3960
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   10560
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "��  ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2640
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   10560
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1800
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   10560
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   10560
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�X  �V"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   10560
      Width           =   855
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   8295
      TabIndex        =   29
      Top             =   1560
      Width           =   120
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�o�׋敪�R�[�h"
      Height          =   240
      Index           =   6
      Left            =   6195
      TabIndex        =   27
      Top             =   1560
      Width           =   1680
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�\������"
      Height          =   240
      Index           =   7
      Left            =   4440
      TabIndex        =   26
      Top             =   1560
      Width           =   960
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�r�r����"
      Height          =   240
      Index           =   5
      Left            =   6840
      TabIndex        =   25
      Top             =   720
      Width           =   960
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�r�r����"
      Height          =   255
      Index           =   4
      Left            =   8040
      TabIndex        =   24
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�X�L���i�\���p����"
      Height          =   240
      Index           =   3
      Left            =   600
      TabIndex        =   23
      Top             =   1560
      Width           =   2160
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���Ӑ於��"
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   22
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "���Ӑ溰��"
      Height          =   240
      Index           =   2
      Left            =   600
      TabIndex        =   21
      Top             =   720
      Width           =   1200
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����O"
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   20
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "F1010601"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const Combo_Max% = 0
Private Const Command_Max% = 11

Private Const ptxMUKE_CODE% = 0
Private Const ptxMUKE_NAME% = 1
Private Const ptxSS_CODE% = 2
Private Const ptxSS_NAME% = 3
Private Const ptxMUKE_DNAME% = 4

Private Const ptxRANKING% = 5
Private Const ptxSYUKA_KBN% = 6


Private Const Text_Max% = 6


Private Const pcmbNaiGai% = 0

Private MTS_CSV As String

Private wkMTS_CHG_CD(0 To 1295) As String * 2
Private Const LAST_UPDATE_DAY$ = "[F101060] 2019.06.25 11:15"  '2019.06.25 ��ʃT�C�Y�g��


Private Function List_Proc() As Integer
'----------------------------------------------------------------------------
'                   ���X�g�{�b�N�X�\������
'----------------------------------------------------------------------------
Dim sts As Integer
Dim com As Integer

    List_Proc = True
    
    List1.Clear
    
    com = BtOpGetFirst
    Do
        sts = BTRV(com, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "������Ǘ��}�X�^")
                Exit Function
        End Select
        
        Call List_Edit_Proc
         
        com = BtOpGetNext
    Loop
    
    List_Proc = False
    
End Function
Private Sub Clear_Field(Mode As Integer)
'----------------------------------------------------------------------------
'                   ��ʏ�������
'----------------------------------------------------------------------------
Dim i As Integer

    
    For i = Mode To Text_Max
            Text(i) = ""
    Next i

End Sub
Private Function Err_Chk() As Integer
'----------------------------------------------------------------------------
'                   ���͍��ڂ̃G���[�`�F�b�N
'----------------------------------------------------------------------------
Dim sts     As Integer
    
    Err_Chk = True
    If Len(Text(ptxMUKE_CODE).Text) = 0 Then
        Beep
        MsgBox "���͂������ڂ̓G���[�ł��B(�K�{����)"
        Text(ptxMUKE_CODE).SetFocus
        Exit Function
    End If
        
    If IsNumeric(Text(ptxRANKING).Text) Then
        Text(ptxRANKING).Text = Format(CInt(Text(ptxRANKING).Text), "000")
    End If
        
    If Trim(Text(ptxSYUKA_KBN).Text) = "" Then
        Label(8).Caption = ""
    Else
        Call UniCode_Conv(K0_SE_SHIP_TANKA_M.SE_SYUKA_KBN, Text(ptxSYUKA_KBN).Text)
        sts = BTRV(BtOpGetEqual, SE_SHIP_TANKA_M_POS, SE_SHIP_TANKA_M_REC, Len(SE_SHIP_TANKA_M_REC), K0_SE_SHIP_TANKA_M, Len(K0_SE_SHIP_TANKA_M), 0)
        Select Case sts
            Case BtNoErr
                Label(8).Caption = Trim(StrConv(SE_SHIP_TANKA_M_REC.SE_SYUKA_NAME, vbUnicode))
            Case BtErrKeyNotFound
                Label(8).Caption = ""
                Beep
                MsgBox "���͂������ڂ̓G���[�ł��B(�P���ݒ薢�o�^)"
                Text(ptxSYUKA_KBN).SetFocus
                Exit Function
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�o�א�ʒP���ݒ�}�X�^")
                Exit Function
        End Select
    End If
        
        
    Err_Chk = False
End Function
Private Function Dislpay_Proc() As Integer
'----------------------------------------------------------------------------
'                   ���R�[�h���e�̕\��
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim i       As Integer

    Dislpay_Proc = True

    Call UniCode_Conv(K0_MTS.MUKE_CODE, Text(ptxMUKE_CODE).Text)
    Call UniCode_Conv(K0_MTS.SS_CODE, Text(ptxSS_CODE).Text)
    
    sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Dislpay_Proc = False
            Exit Function
        Case Else
            Call File_Error(sts, BtOpGetEqual, "������Ǘ��}�X�^")
            Exit Function
    End Select
    
    
    For i = 0 To 1
        
        If Right(Combo(pcmbNaiGai).List(i), 1) = StrConv(MTSREC.NAIGAI, vbUnicode) Then
            Combo(pcmbNaiGai).ListIndex = i
            Exit For
        End If
    
    Next i
    
    Text(ptxMUKE_NAME).Text = StrConv(MTSREC.MUKE_NAME, vbUnicode)
    Text(ptxSS_NAME).Text = StrConv(MTSREC.SS_NAME, vbUnicode)
    Text(ptxMUKE_DNAME).Text = StrConv(MTSREC.MUKE_DNAME, vbUnicode)

    Text(ptxRANKING).Text = StrConv(MTSREC.DISPLAY_RANKING, vbUnicode)

    Text(ptxSYUKA_KBN).Text = StrConv(MTSREC.SYUKA_KBN, vbUnicode)
    Call UniCode_Conv(K0_SE_SHIP_TANKA_M.SE_SYUKA_KBN, Text(ptxSYUKA_KBN).Text)
    sts = BTRV(BtOpGetEqual, SE_SHIP_TANKA_M_POS, SE_SHIP_TANKA_M_REC, Len(SE_SHIP_TANKA_M_REC), K0_SE_SHIP_TANKA_M, Len(K0_SE_SHIP_TANKA_M), 0)
    Select Case sts
        Case BtNoErr
            Label(8).Caption = Trim(StrConv(SE_SHIP_TANKA_M_REC.SE_SYUKA_NAME, vbUnicode))
        Case BtErrKeyNotFound
            Label(8).Caption = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�o�א�ʒP���ݒ�}�X�^")
            Exit Function
    End Select


    Dislpay_Proc = False
End Function
Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   �ǉ��^�ύX����
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer
Dim ans             As Integer

Dim wkMUKE_CHG_CD   As String * 2
    
    Update_Proc = True

    Call UniCode_Conv(K0_MTS.MUKE_CODE, Text(ptxMUKE_CODE).Text)
    Call UniCode_Conv(K0_MTS.SS_CODE, Text(ptxSS_CODE).Text)


    Do
        
        sts = BTRV(BtOpGetEqual + BtSNoWait, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<MTS.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Update_Proc = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "������Ǘ��}�X�^")
                Exit Function
        End Select
    
    Loop

    If com = BtOpInsert Then
        Call UniCode_Conv(MTSREC.MUKE_CODE, Text(ptxMUKE_CODE).Text)
        Call UniCode_Conv(MTSREC.SS_CODE, Text(ptxSS_CODE).Text)
        Call UniCode_Conv(MTSREC.DATA_KBN, "")
        Call UniCode_Conv(MTSREC.FILLER, "")
    End If

    Call UniCode_Conv(MTSREC.NAIGAI, Right(Combo(pcmbNaiGai).Text, 1))
    Call UniCode_Conv(MTSREC.MUKE_NAME, Text(ptxMUKE_NAME).Text)
    Call UniCode_Conv(MTSREC.SS_NAME, Text(ptxSS_NAME).Text)
    Call UniCode_Conv(MTSREC.MUKE_DNAME, Text(ptxMUKE_DNAME).Text)
    Call UniCode_Conv(MTSREC.DISPLAY_RANKING, Text(ptxRANKING).Text)
    Call UniCode_Conv(MTSREC.SYUKA_KBN, Text(ptxSYUKA_KBN).Text)


    Do
        sts = BTRV(com, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<MTS.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Update_Proc = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, com, "������Ǘ��}�X�^")
                Exit Function
        End Select
    Loop

    Call List_Update_Proc(0)

    Call Clear_Field(0)

    Update_Proc = False

End Function
Private Function Delete_Proc() As Integer
'----------------------------------------------------------------------------
'                   �폜����
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer
Dim ans             As Integer

    
    Delete_Proc = True

    Call UniCode_Conv(K0_MTS.MUKE_CODE, Text(ptxMUKE_CODE).Text)
    Call UniCode_Conv(K0_MTS.SS_CODE, Text(ptxSS_CODE).Text)


    Do
        
        sts = BTRV(BtOpGetEqual + BtSNoWait, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
        Select Case sts
            Case BtNoErr
                Do
                    sts = BTRV(BtOpDelete, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<MTS.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                Delete_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        Case Else
                            Call File_Error(sts, BtOpDelete, "������Ǘ��}�X�^")
                            Exit Function
                    End Select
                Loop
                Exit Do
            Case BtErrKeyNotFound
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<MTS.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Delete_Proc = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "������Ǘ��}�X�^")
                Exit Function
        End Select
    
    Loop

    Call List_Update_Proc(1)

    Call Clear_Field(0)

    Delete_Proc = False

End Function

Private Sub Command_Click(Index As Integer)

Dim yn  As Integer
Dim sts As Integer

    Select Case Index
        Case 0
                                            
            sts = Err_Chk()             '�G���[�`�F�b�N
            Select Case sts
                Case False
                Case True
                    Exit Sub
                Case SYS_ERR
                    Unload Me
            End Select
            Beep
            yn = MsgBox("�X�V���܂����H", vbYesNo + vbQuestion, "�m�F����")
            
            If yn = vbYes Then
                sts = Update_Proc()
                Select Case sts
                    Case False
                    Case True
                        Unload Me
                    Case SYS_CANCEL
                End Select
            End If
            
            Text(ptxMUKE_CODE) = ""
        
        Case 3
            If Trim(Text(ptxMUKE_CODE)) = "" Then
                Beep
                MsgBox "�폜����R�[�h���w�肵�ĉ������B", vbExclamation
            Else
                Beep
                yn = MsgBox("�폜���܂����H", vbYesNo + vbQuestion, "�m�F����")
                If yn = vbYes Then
                    If Delete_Proc() Then
                        Unload Me
                    End If
                End If
            End If
        Case 8                  '�f�[�^�o��
            Beep
            yn = MsgBox("�f�[�^�o�͂��܂����H", vbYesNo + vbQuestion, "�m�F����")
            If yn = vbYes Then
                If Data_Proc() Then
                    Unload Me
                End If
            End If
        Case 11
            Unload Me
        Case Else
            Beep
    End Select
    
    Text(ptxMUKE_CODE).SetFocus

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
Dim i       As Integer
Dim j       As Integer
Dim k       As Integer
Dim c       As String * 128
Dim sts     As Integer

Dim Wk      As String * 36


    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)
                                '�b�r�u�t�@�C������荞��
    If GetIni("FILE", "MTS_CSV", "SYS", c) Then
        Beep
        MsgBox "������Ǘ��}�X�^�f�[�^�o�͗p�t�@�C��[MTS_CSV]�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    MTS_CSV = Trim(c)
    
    Me.Caption = Me.Caption & " " & LAST_UPDATE_DAY '2019.06.25 �^�C�g���o�[�\���p�Œǉ�
    
                                '������Ǘ��}�X�^�n�o�d�m
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '�o�וʒP���ݒ�}�X�^�n�o�d�m
    If SE_SHIP_TANKA_M_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                
    Wk = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    k = 0
    For i = 1 To 36
        
        For j = 1 To 36
            wkMTS_CHG_CD(k) = Mid(Wk, i, 1) & Mid(Wk, j, 1)
            k = k + 1
        Next j
    Next i
                                
                                
                                '�����O�ݒ�
    Combo(pcmbNaiGai).Clear
    Combo(pcmbNaiGai).AddItem NAIGAI1 & Space(4) & NAIGAI_NAI   '����
    Combo(pcmbNaiGai).AddItem NAIGAI2 & Space(4) & NAIGAI_GAI   '�C�O
    Combo(pcmbNaiGai).ListIndex = 0
    
    Show
                                
    If List_Proc() Then
        Unload Me
    End If
                                '��ʏ����ݒ�
    Clear_Field (0)
    
    Text(ptxMUKE_CODE).SetFocus
    
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
                                            '�o�וʒP���ݒ�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, SE_SHIP_TANKA_M_POS, SE_SHIP_TANKA_M_REC, Len(SE_SHIP_TANKA_M_REC), K0_SE_SHIP_TANKA_M, Len(K0_SE_SHIP_TANKA_M), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�o�וʒP���ݒ�")
        End If
    End If
    sts = BTRV(BtOpReset, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "������Ǘ��}�X�^")
    End If
    Set F1010601 = Nothing
    End
End Sub
Private Sub List1_DblClick()

Dim i       As Integer
Dim CODE    As String * 16

        CODE = Right(List1.List(List1.ListIndex), 16)

        Text(ptxMUKE_CODE).Text = Left(CODE, 8)
        Text(ptxSS_CODE).Text = Right(CODE, 8)

        If Dislpay_Proc() Then
            Unload Me
        End If

        Text(ptxMUKE_CODE).SetFocus

End Sub

Private Sub List1_GotFocus()
    
    If List1.ListCount > 0 Then
        If List1.ListIndex <= 0 Then
            List1.ListIndex = 0
        End If
    End If
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim sts As Integer

    Select Case KeyCode
        Case vbKeyReturn
            
            Call List1_DblClick
        Case vbKeyF12
            Command(11).Value = True
    End Select

End Sub


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
        Case ptxSS_CODE
            If Len(Trim(Text(ptxMUKE_CODE).Text)) = 0 Then
                Beep
                MsgBox "���͂������ڂ̓G���[�ł��B�i�K�{���́j"
                Text(ptxMUKE_CODE).SetFocus
                Exit Sub
            End If
    
            If Dislpay_Proc() Then
                Unload Me
            End If
    
    End Select
    
    For i = Index + 1 To Text_Max
        If Text(i).Enabled And Text(i).Visible And Text(i).TabStop Then
            Text(i).SetFocus
            Exit For
        End If
    Next i
    
End Sub
Private Function Data_Proc() As Integer

Dim FileNo          As Integer
Dim FileName        As String
Dim Ret             As Integer

Dim com             As Integer
Dim sts             As Integer

    Call Input_Lock

    FileNo = FreeFile
    FileName = MTS_CSV
    
    Ret = InStr(1, Trim(FileName), ".") - 1
    FileName = Left(Trim(FileName), Ret) & Right(Trim(FileName), Len(Trim(FileName)) - Ret)

    On Error GoTo Error_Proc

    Open (FileName) For Output As FileNo
    
    Write #FileNo, "�����O", "���Ӑ溰��", "���Ӑ於��", "�q�Ɂ^�r�r����", "�q�Ɂ^�r�r����", "�\�����́i�X�L���i�p�j", "�Ǒւ����ށi���������p�j"

    com = BtOpGetFirst
    Do
        DoEvents
        sts = BTRV(com, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "������Ǘ��}�X�^")
                Exit Function
        End Select
    
        Write #FileNo, StrConv(MTSREC.NAIGAI, vbUnicode),
        Write #FileNo, StrConv(MTSREC.MUKE_CODE, vbUnicode),
        Write #FileNo, StrConv(MTSREC.MUKE_NAME, vbUnicode),
        Write #FileNo, StrConv(MTSREC.SS_CODE, vbUnicode),
        Write #FileNo, StrConv(MTSREC.SS_NAME, vbUnicode),
        Write #FileNo, StrConv(MTSREC.MUKE_DNAME, vbUnicode)
    
    
        com = BtOpGetNext
    Loop

    Close #FileNo
    
    Call Input_UnLock
    
    Beep
    MsgBox "�u" & FileName & "�v�͐���ɏo�͂���܂����B"


    Exit Function
Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox FileName & "���g�p���ł��B"
        Data_Proc = False
    Else
        MsgBox "Err.Number= " & Err.Number
        Data_Proc = True
    End If

    Call Input_UnLock



End Function


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------
Dim i As Integer

    F1010601.MousePointer = vbHourglass

    Call Ctrl_Lock(F1010601)

End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------
Dim i As Integer

    Call Ctrl_UnLock(F1010601)

    F1010601.MousePointer = vbDefault

End Sub


Public Sub List_Edit_Proc()
'----------------------------------------------------------------------------
'                   ���X�g�{�b�N�X���ו\��
'----------------------------------------------------------------------------

Dim Edit    As String

    
        
    Select Case StrConv(MTSREC.NAIGAI, vbUnicode)
        Case NAIGAI_NAI
            Edit = NAIGAI1 & " "
        Case NAIGAI_GAI
            Edit = NAIGAI2 & " "
    End Select
    Edit = Edit & StrConv(MTSREC.MUKE_CODE, vbUnicode) & " "
    Edit = Edit & StrConv(MTSREC.MUKE_NAME, vbUnicode) & " "
    Edit = Edit & StrConv(MTSREC.SS_CODE, vbUnicode) & " "
    Edit = Edit & StrConv(MTSREC.SS_NAME, vbUnicode) & StrConv(MTSREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.SS_CODE, vbUnicode)
    
    List1.AddItem Edit

End Sub


Private Sub List_Update_Proc(Mode As Integer)
'----------------------------------------------------------------------------
'                   ���X�g�{�b�N�X�X�V
'----------------------------------------------------------------------------
Dim i       As Integer
Dim CODE    As String * 16


    For i = 0 To List1.ListCount - 1
        
        CODE = Right(List1.List(i), 16)
        
        If Trim(Text(ptxMUKE_CODE).Text) = Trim(Left(CODE, 8)) And _
            Trim(Text(ptxSS_CODE).Text) = Trim(Right(CODE, 8)) Then
                List1.RemoveItem i
        End If
    
    Next i

    If Mode = 0 Then
        Call List_Edit_Proc
    End If
End Sub
