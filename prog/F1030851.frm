VERSION 5.00
Begin VB.Form F1030851 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�o�ח\�胁���e�i���X"
   ClientHeight    =   7425
   ClientLeft      =   2130
   ClientTop       =   2715
   ClientWidth     =   13605
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
   ScaleHeight     =   7425
   ScaleWidth      =   13605
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text 
      Height          =   360
      Index           =   0
      Left            =   3360
      MaxLength       =   8
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�S���Ώ�"
      Height          =   495
      Left            =   10290
      TabIndex        =   43
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox txtID_No 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Left            =   2595
      Locked          =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1455
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   1
      Left            =   4620
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   2
      Top             =   360
      Width           =   5355
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   4
      Left            =   4320
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   5
      Left            =   5280
      MaxLength       =   2
      TabIndex        =   7
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   6
      Left            =   6000
      MaxLength       =   2
      TabIndex        =   8
      Top             =   1080
      Width           =   375
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2760
      Width           =   13110
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   7
      Left            =   1110
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1170
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   2880
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   2160
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   10
      Left            =   10290
      TabIndex        =   12
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   9
      Left            =   6795
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1800
      Width           =   3135
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   8
      Left            =   4155
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1800
      Width           =   2535
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   1320
      Locked          =   -1  'True
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�I ��"
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
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6480
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
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   6480
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
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   6480
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
      Index           =   8
      Left            =   7800
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�\ ��"
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
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6480
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
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6480
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6480
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
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�� ��"
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
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6480
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
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6480
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�X �V"
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6480
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�h�c��"
      Height          =   255
      Index           =   0
      Left            =   2625
      TabIndex        =   42
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�ϐ�"
      Height          =   255
      Index           =   13
      Left            =   11445
      TabIndex        =   40
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�\�萔"
      Height          =   255
      Index           =   12
      Left            =   10395
      TabIndex        =   39
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�`�[���t"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   38
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�`�[��"
      Height          =   255
      Index           =   10
      Left            =   1575
      TabIndex        =   37
      Top             =   2520
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   10920
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "������"
      Height          =   252
      Index           =   9
      Left            =   3360
      TabIndex        =   36
      Top             =   1080
      Width           =   732
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   255
      Index           =   8
      Left            =   5040
      TabIndex        =   35
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   255
      Index           =   7
      Left            =   5760
      TabIndex        =   34
      Top             =   1200
      Width           =   255
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
      TabIndex        =   33
      Top             =   7080
      Width           =   180
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "������"
      Height          =   255
      Index           =   17
      Left            =   2520
      TabIndex        =   32
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   255
      Index           =   3
      Left            =   2640
      TabIndex        =   31
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   30
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�`�[���t"
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   29
      Top             =   1080
      Width           =   972
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����敪"
      Height          =   255
      Index           =   14
      Left            =   240
      TabIndex        =   28
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i  ��"
      Height          =   255
      Index           =   6
      Left            =   5985
      TabIndex        =   27
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i  ��"
      Height          =   255
      Index           =   5
      Left            =   4200
      TabIndex        =   26
      Top             =   2520
      Width           =   735
   End
   Begin VB.Menu MainMenu 
      Caption         =   "���ƕ�"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1030851"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim P_ID As String * 8
Dim WS_ID As String * 3


Private Const Text_Max% = 10
Private Const ptxMUKE_CODE% = 0         '������i�R�[�h���͗p�j
Private Const ptxM_YY% = 1
Private Const ptxM_MM% = 2
Private Const ptxM_DD% = 3
Private Const ptxC_YY% = 4
Private Const ptxC_MM% = 5
Private Const ptxC_DD% = 6
Private Const ptxNO% = 7
Private Const ptxITEM% = 8
Private Const ptxI_NM% = 9
Private Const ptxQTY% = 10




Private Const pcmbC_KBN% = 0
Private Const pcmbMUKE_CODE% = 1

Private Upd_Mode    As Integer

Private Const LAST_UPDATE_DAY$ = "[F103085]2018.04.21 09:00"

                                    '��ʏ�����Ԃ�ݒ肷��
Private Sub Clear_Field(Start_Pos As Integer)
Dim i As Integer
    
    For i = Start_Pos To Text_Max
        Text(i).Text = ""
    Next i
    
    txtID_No.Text = ""
    
End Sub
Private Function P_Off() As Integer

Dim sts As Integer
Dim com As Integer
Dim yn As Integer
    
    P_Off = True
    
    Call UniCode_Conv(K4_Y_SYU.WEL_ID, WS_ID)
    Call UniCode_Conv(K4_Y_SYU.PRG_ID, P_ID)
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K4_Y_SYU, Len(K4_Y_SYU), 4)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
                P_Off = False
                Exit Function
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                yn = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                
                If yn = vbCancel Then
                    P_Off = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�o�ח\��f�[�^")
                Exit Function
        End Select
    Loop
        
    Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
    Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
        
'    Select Case StrConv(Y_SYUREC.KAN_KBN, vbUnicode)
'        Case KAN_L_SOFF_POFF_KOFF
'            Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_SOFF_POFF_KOFF)
'        Case KAN_L_SING_POFF_KOFF
'            Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_SING_POFF_KOFF)
'        Case KAN_L_SOFF_PON_KOFF
'            Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_SOFF_PON_KOFF)
'        Case KAN_L_SING_PON_KOFF
'            Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_SING_PON_KOFF)
'    End Select
        
    Do
        sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K4_Y_SYU, Len(K4_Y_SYU), 4)
        Select Case sts
            Case BtNoErr
                Exit Do
                
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                yn = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                
                If yn = vbCancel Then
                    
                    sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K4_Y_SYU, Len(K4_Y_SYU), 4)
                    If sts Then
                        Call File_Error(sts, BtOpUnlock, "�o�ח\��f�[�^")
                        Exit Function
                    End If
                    P_Off = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�o�ח\��f�[�^")
                Exit Function
        End Select
    Loop

    P_Off = False
End Function

Private Function Item_Dsp() As Integer
                                    '�e���ڂ�\������
Dim sts         As Integer
Dim yn          As Integer
Dim i           As Integer
Dim Qty         As Long
Dim ans         As Integer
    
    Item_Dsp = True
    
        
    sts = P_Off
    Select Case sts
        Case False
        Case SYS_ERR
            Exit Function
        Case SYS_CANCEL
            List1.SetFocus
            Item_Dsp = SYS_CANCEL
            Exit Function
    End Select
    
                                                '�o�ח\��ǂݍ���
    Call UniCode_Conv(K0_Y_SYU.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, Mid(List1.List(List1.ListIndex), 22, 7) & Mid(List1.List(List1.ListIndex), 30, 2))
    
    
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
        Select Case sts
            Case BtNoErr
                If Len(Trim(StrConv(Y_SYUREC.WEL_ID, vbUnicode))) <> 0 And _
                    StrConv(Y_SYUREC.WEL_ID, vbUnicode) <> WS_ID Then
                    
                    sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                    If sts Then
                        Call File_Error(sts, BtOpUnlock, "�o�ח\��f�[�^")
                        Exit Function
                    End If
                    Beep
                    yn = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If yn = vbCancel Then
                        List1.SetFocus
                        Item_Dsp = SYS_CANCEL
                        Exit Function
                    End If
                Else
                    Exit Do
                End If
            Case BtErrKeyNotFound
                Beep
                MsgBox "�f�[�^���e���ύX����Ă��܂��B�ŐV���e��\�����܂��B"
                If List_Dsp() Then
                    Exit Function
                End If
                List1.SetFocus
                Item_Dsp = False
                Exit Function
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                yn = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If yn = vbCancel Then
                    sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K4_Y_SYU, Len(K4_Y_SYU), 4)
                    If sts Then
                        Call File_Error(sts, BtOpUnlock, "�o�ח\��f�[�^")
                        Exit Function
                    End If
                    List1.SetFocus
                    Item_Dsp = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�o�ח\��f�[�^")
                Exit Function
        End Select
    Loop
                                        '�����̃`�F�b�N
    If Upd_Mode = 1 Then
    Else
        If CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) = CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)) Then
            sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
            If sts Then
                Call File_Error(sts, BtOpUnlock, "�o�ח\��f�[�^")
                Exit Function
            End If
                
            Beep
            MsgBox "�w��̓`�[�͏o�Ɋ����ׁ̈A�ύX�ł��܂���B"
            List1.SetFocus
            List1.ListIndex = 0
            Item_Dsp = False
            Exit Function
        End If
    End If
    
    If CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) <> 0 Then
        Beep
        yn = MsgBox("�w��̓`�[�͍�ƒ��ł��B�������p�����܂����H", vbYesNo + vbQuestion + vbDefaultButton2, "�m�F����")
        If yn = vbNo Then
            sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
            If sts Then
                Call File_Error(sts, BtOpUnlock, "�o�ח\��f�[�^")
                Item_Dsp = SYS_ERR
                Exit Function
            End If
            
            List1.SetFocus
            List1.ListIndex = 0
            Item_Dsp = False
            Exit Function
        End If
    End If
        
        
    If Len(Trim(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode))) <> 0 Then
        Beep
        yn = MsgBox("�w��̓`�[�͏o�ɕ\����ςł��B�������p�����܂����H", vbYesNo + vbQuestion + vbDefaultButton2, "�m�F����")
        If yn = vbNo Then
            sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
            If sts Then
                Call File_Error(sts, BtOpUnlock, "�o�ח\��f�[�^")
                Item_Dsp = SYS_ERR
                Exit Function
            End If
                
            List1.SetFocus
            List1.ListIndex = 0
            Item_Dsp = False
            Exit Function
            
        End If
    End If
        
    
    Text(ptxM_YY).Text = Mid(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 1, 4)
    Text(ptxM_MM).Text = Mid(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 5, 2)
    Text(ptxM_DD).Text = Mid(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 7, 2)

    Text(ptxC_YY).Text = Text(ptxM_YY)
    Text(ptxC_MM).Text = Text(ptxM_MM)
    Text(ptxC_DD).Text = Text(ptxM_DD)

    Text(ptxNO).Text = Left(StrConv(Y_SYUREC.DEN_NO, vbUnicode), 7) & "-" & Right(Trim(StrConv(Y_SYUREC.DEN_NO, vbUnicode)), 1)
    txtID_No.Text = Left(StrConv(Y_SYUREC.ID_NO, vbUnicode), 7) & "-" & Right(Trim(StrConv(Y_SYUREC.ID_NO, vbUnicode)), 2)
    
    Text(ptxITEM).Text = StrConv(Y_SYUREC.HIN_NO, vbUnicode)
                                                        '�i�ڃ}�X�^�Ǎ���
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(Y_SYUREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Item_Dsp = SYS_ERR
            Exit Function
    End Select
    
    Text(ptxI_NM) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
    Text(ptxQTY) = Format(CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)), "#0")
    
'    Text(ptxHS_C_KBN) = StrConv(Y_SYUREC.CYU_KBN, vbUnicode)
    
    Call UniCode_Conv(Y_SYUREC.PRG_ID, P_ID)
    Call UniCode_Conv(Y_SYUREC.WEL_ID, WS_ID)
    
    Do
        sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                yn = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If yn = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpUpdate, "�o�ח\��f�[�^")
                Item_Dsp = SYS_ERR
                Exit Function
        End Select
    Loop
    Text(ptxC_YY).SetFocus
    
    Item_Dsp = False
    
    
End Function

Private Function Err_Chk(Mode As Integer) As Integer
                                    
Dim sts As Integer
Dim i   As Integer
                                    '���͍��ڂ̃G���[�`�F�b�N

    Err_Chk = True
                                    '�o�ח\��ǂݍ���
    Call UniCode_Conv(K0_Y_SYU.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, Left(txtID_No, 7) & Right(Trim(txtID_No), 2))
sts = BTRV(BtOpGetEqual, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    Select Case sts
        Case BtNoErr
            If StrConv(Y_SYUREC.WEL_ID, vbUnicode) = WS_ID And _
                StrConv(Y_SYUREC.PRG_ID, vbUnicode) = P_ID Then
            Else
                Beep
                MsgBox "�X�V�Ώۂ̏o�ח\�肪�m�肵�Ă��܂���B"
                List1.SetFocus
                Exit Function
            End If
                                    '�폜���̃`�F�b�N
            If Mode = 1 Then
                If Upd_Mode = 0 Then
                    If CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) <> 0 Then
                        Beep
                        MsgBox "�o�Ɏ��т��L��̂ō폜�ł��܂���B"
                        List1.SetFocus
                    Else
                        Err_Chk = False
                        Exit Function
                    End If
                End If
            End If
        Case BtErrKeyNotFound
            Beep
            MsgBox "�X�V�Ώۂ̏o�ח\�肪�m�肵�Ă��܂���B"
            List1.SetFocus
            Exit Function
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�o�ח\��f�[�^")
            Err_Chk = SYS_ERR
            Exit Function
    End Select


    For i = ptxC_YY To ptxC_DD
        If Trim(Text(i).Text) = "" Then
            Text(i).Text = "0"
        End If
    
        If Not IsNumeric(Text(i).Text) Then
            Beep
            MsgBox "���͂������ڂ̓G���[�ł��B"
            Text(i).SetFocus
            Exit Function
        Else
            Text(i).Text = Right(Format(CInt(Text(i).Text), "0000"), Text(i).MaxLength)
        End If
    
    Next i

    If Not IsDate(Text(ptxC_YY).Text & "/" & Text(ptxC_MM).Text & "/" & Text(ptxC_DD).Text) Then
        Beep
        MsgBox "���͂������ڂ̓G���[�ł��B"
        Text(ptxC_YY).SetFocus
        Exit Function
    End If
    If (Text(ptxC_YY).Text & "/" & Text(ptxC_MM).Text & "/" & Text(ptxC_DD).Text) < Format(Date, "YYYY/MM/DD") Then
        Beep
        MsgBox "���͂������ڂ̓G���[�i���{���j�ł��B"
        Text(ptxC_YY).SetFocus
        Exit Function
    End If

    If Not IsNumeric(Text(ptxQTY).Text) Then
        Beep
        MsgBox "���͂������ڂ̓G���[�ł��B"
        Text(ptxQTY).SetFocus
        Exit Function
    Else
        Text(ptxQTY).Text = Format(CLng(Text(ptxQTY).Text), "#0")
    End If

    If CLng(Text(ptxQTY).Text) <= 0 Then
        Beep
        MsgBox "���͂������ڂ̓G���[�ł��B"
        Text(ptxQTY).SetFocus
        Exit Function
    End If
                                    '���ʕύX�̃`�F�b�N
    If CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) > CLng(Text(ptxQTY).Text) Then
        Beep
        MsgBox "�o�Ɏ��і����ւ̐��ʕύX�͂ł��܂���B"
        Text(ptxQTY).SetFocus
        Exit Function
    End If
    
    Err_Chk = False
    
End Function

                                            '�o�ח\��̒����^�폜
Private Function Update_Proc(Mode As Integer) As Integer

Dim sts As Integer
Dim ans As Integer
                                            
    Update_Proc = True
    
                                    '�g�����U�N�V�����J�n
    sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
        Exit Function
    End If
    
    
    
    '-------------------------------------------    �o�ח\��̓ǂݍ���
    Call UniCode_Conv(K0_Y_SYU.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, Left(txtID_No, 7) & Right(Trim(txtID_No), 2))
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
        Select Case sts
            Case BtNoErr
                
                Exit Do
            Case BtErrKeyNotFound
                Beep
                MsgBox "�X�V�Ώۂ̏o�ח\�肪�m�肵�Ă��܂���B"
                List1.SetFocus
                Update_Proc = False
                GoTo Abort_Tran
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    List1.SetFocus
                    Update_Proc = False
                    GoTo Abort_Tran
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�o�ח\��f�[�^")
                Update_Proc = SYS_ERR
                GoTo Abort_Tran
        End Select
    Loop
    
    
    '-------------------------------------------    �o�ח\��(νĲҰ��)�̓ǂݍ���
    Call UniCode_Conv(K4_Y_SYU_H.ID_NO, StrConv(Y_SYUREC.ID_NO, vbUnicode))
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
        Select Case sts
            Case BtNoErr
                
                Exit Do
            Case BtErrKeyNotFound
                Beep
                MsgBox "�ް��ݸ�ُ�ł��B"
                List1.SetFocus
                Update_Proc = False
                GoTo Abort_Tran
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA_H.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    List1.SetFocus
                    Update_Proc = False
                    GoTo Abort_Tran
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�o�ח\��f�[�^(νĲҰ��)")
                Update_Proc = SYS_ERR
                GoTo Abort_Tran
        End Select
    Loop
    
    
    
    Select Case Mode
        Case 0                              '����
                                            
            If SYUKA_LOG_ON Then
                Call SYUKA_LOG_OUT_PROC("UPD", "BEF")
            End If
                                            
            '--------------------------------   '�o�ח\��f�[�^�X�V
            Call UniCode_Conv(Y_SYUREC.SURYO, Format(CLng(Text(ptxQTY).Text), "0000000"))
            Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, _
                            Text(ptxC_YY).Text & Text(ptxC_MM).Text & Text(ptxC_DD).Text)
            Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
            Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
            
            If CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) = CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)) Then
                                            '�o�Ɋ���
                Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_KBN_FIN)
            
            End If
            
            Do
                sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            List1.SetFocus
                            Update_Proc = False
                            GoTo Abort_Tran
                        End If
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "�o�ח\��f�[�^")
                        Update_Proc = SYS_ERR
                        GoTo Abort_Tran
                End Select
            Loop
            
            
            '-------------------------------------------    �o�ח\��(νĲҰ��)�̍X�V
            Call UniCode_Conv(Y_SYU_HREC.SURYO, Format(CLng(Text(ptxQTY).Text), "0000000"))
            Call UniCode_Conv(Y_SYU_HREC.SYUKA_YMD, _
                            Text(ptxC_YY).Text & Text(ptxC_MM).Text & Text(ptxC_DD).Text)
            
            Do
                sts = BTRV(BtOpUpdate, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA_H.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            List1.SetFocus
                            Update_Proc = False
                            GoTo Abort_Tran
                        End If
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "�o�ח\��f�[�^(νĲҰ��)")
                        Update_Proc = SYS_ERR
                        GoTo Abort_Tran
                End Select
            Loop
                                '�g�����U�N�V�����I��
            sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call File_Error(sts, BtOpAbortTransaction, "")
                GoTo Abort_Tran
            End If
            
            If SYUKA_LOG_ON Then
                Call SYUKA_LOG_OUT_PROC("UPD", "AFT")
            End If
            
            List1.RemoveItem List1.ListIndex
            Call List_Edit
        Case 1
            '--------------------------------   '�o�ח\��f�[�^�폜
            Do
                sts = BTRV(BtOpDelete, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            Update_Proc = False
                            List1.SetFocus
                            GoTo Abort_Tran
                        End If
                    Case Else
                        Call File_Error(sts, BtOpDelete, "�o�ח\��f�[�^")
                        Update_Proc = SYS_ERR
                        GoTo Abort_Tran
                End Select
            Loop
            
            
            '-------------------------------------------    �o�ח\��(νĲҰ��)�̍폜
            Do
                sts = BTRV(BtOpDelete, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            Update_Proc = False
                            List1.SetFocus
                            GoTo Abort_Tran
                        End If
                    Case Else
                        Call File_Error(sts, BtOpDelete, "�o�ח\��f�[�^(νĲҰ��)")
                        Update_Proc = SYS_ERR
                        GoTo Abort_Tran
                End Select
            Loop
            
            
                                '�g�����U�N�V�����I��
            sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call File_Error(sts, BtOpAbortTransaction, "")
                GoTo Abort_Tran
            End If
            
            
            
            If SYUKA_LOG_ON Then
                Call SYUKA_LOG_OUT_PROC("DEL", "AFT")
            End If
            
            List1.RemoveItem List1.ListIndex
        
    End Select
    Update_Proc = False
    Exit Function

'�ُ�I��
Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If
    
End Function

Private Function MTS_Set_Proc() As Integer

Dim sts     As Integer
Dim com     As Integer
    
Dim Edit    As String
    
    MTS_Set_Proc = True
    
'    Call Input_Lock
    
    
    Combo(pcmbMUKE_CODE).Clear
    
    com = BtOpGetFirst
    Do
        sts = BTRV(com, MTS_POS, MTSREC, Len(MTSREC), K1_MTS, Len(K1_MTS), 1)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "������}�X�^")
                Exit Function
        End Select
        
        Edit = StrConv(MTSREC.MUKE_NAME, vbUnicode) & "   "
        Edit = Edit & StrConv(MTSREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.SS_CODE, vbUnicode)
        
        
        Combo(pcmbMUKE_CODE).AddItem Edit
    
        com = BtOpGetNext
    
    Loop

    If Combo(pcmbMUKE_CODE).ListCount <= 0 Then
    Else
        Combo(pcmbMUKE_CODE).ListIndex = 0
    End If

'    Call Input_UnLock

    MTS_Set_Proc = False
End Function

Private Function List_Dsp() As Integer
Dim sts As Integer
Dim com As Integer
Dim yn As Integer
Dim WS01 As Integer
Dim W_Str As String

    List_Dsp = True

    List1.Clear
    
    
                                                    '���ƕ�
    Call UniCode_Conv(K2_Y_SYU.JGYOBU, Last_JGYOBU)
                                                    '�����敪
    Call UniCode_Conv(K2_Y_SYU.KEY_CYU_KBN, Right(Combo(pcmbC_KBN).Text, 1))
                                                    '������
    Call UniCode_Conv(K2_Y_SYU.KEY_MUKE_CODE, Left(Right(Combo(pcmbMUKE_CODE).Text, 16), 8))
    Call UniCode_Conv(K2_Y_SYU.KEY_SS_CODE, Right(Combo(pcmbMUKE_CODE).Text, 8))
    
    
    
    com = BtOpGetGreaterEqual
    Do
        sts = BTRV(com, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K2_Y_SYU, Len(K2_Y_SYU), 2)
        Select Case sts
            Case BtNoErr
                If StrConv(Y_SYUREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�o�ח\��f�[�^")
                Exit Function
        End Select
        
        If StrConv(Y_SYUREC.KEY_CYU_KBN, vbUnicode) <> Right(Combo(pcmbC_KBN).Text, 1) Then
            Exit Do
        End If
        
        If StrConv(Y_SYUREC.KEY_MUKE_CODE, vbUnicode) <> Left(Right(Combo(pcmbMUKE_CODE).Text, 16), 8) Or _
           StrConv(Y_SYUREC.SS_CODE, vbUnicode) <> Right(Combo(pcmbMUKE_CODE).Text, 8) Then
            Exit Do
        End If
        
        
        If Upd_Mode = 0 Then
            If Len(Trim(StrConv(Y_SYUREC.KAN_YMD, vbUnicode))) <> 0 Then
            Else
                If List_Edit() Then
                    Exit Function
                End If
            End If
        Else
            If List_Edit() Then
                Exit Function
            End If
        End If
        
        com = BtOpGetNext
    
    Loop

    List_Dsp = False

End Function
Private Function List_Edit() As Integer
Dim sts     As Integer
Dim com     As Integer
Dim yn      As Integer
Dim WS01    As Integer
Dim Edit    As String
Dim RetBuf  As String
    
    List_Edit = True
    
    Edit = ""
    Edit = Edit & Mid(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 1, 4) & "/" & Mid(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 5, 2) & "/" & Mid(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 7, 2) & " "
    Edit = Edit & Left(StrConv(Y_SYUREC.DEN_NO, vbUnicode), 7) & "-" & Right(Trim(StrConv(Y_SYUREC.DEN_NO, vbUnicode)), 1)
    Edit = Edit & "(" & Left(StrConv(Y_SYUREC.ID_NO, vbUnicode), 7) & "-" & Right(Trim(StrConv(Y_SYUREC.ID_NO, vbUnicode)), 2) & ") "
    Edit = Edit & Left(StrConv(Y_SYUREC.HIN_NO, vbUnicode), 13) & "  "
                                                        '�i�ڃ}�X�^�Ǎ���
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(Y_SYUREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
        
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�o�ח\��f�[�^")
            Exit Function
    End Select
    
    Edit = Edit & Left(StrConv(ITEMREC.HIN_NAME, vbUnicode), 25) & " "
    
            
    RetBuf = Format(CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)), "#,##0")
    RetBuf = Space(6 - Len(RetBuf)) & Trim(RetBuf) & "("
    Edit = Edit & RetBuf
    RetBuf = Format(CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)), "#,##0")
    RetBuf = Space(6 - Len(RetBuf)) & Trim(RetBuf) & ")"
    Edit = Edit & RetBuf
    
    
    List1.AddItem Edit

    List_Edit = False

End Function


Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim sts As Integer
  
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    Select Case Index
        
        Case pcmbC_KBN
        Case pcmbMUKE_CODE
            
            Text(ptxMUKE_CODE).Text = Trim(Left(Right(Combo(pcmbMUKE_CODE).Text, 16), 8))
            
            
            
            '�o�ח\��\��
            If List_Dsp() Then
                Unload Me
            End If
    
    End Select
End Sub



Private Sub Command_Click(Index As Integer)
Dim yn As Integer
Dim sts As Integer

    Select Case Index
        Case 0
                                            '�G���[�`�F�b�N
            sts = Err_Chk(0)
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
                sts = Update_Proc(0)
                Select Case sts
                    Case True, False
                    Case SYS_ERR
                        Unload Me
                End Select
            Else
                sts = P_Off()
                Select Case sts
                    Case False
                    Case SYS_ERR
                        Unload Me
                    Case SYS_CANCEL
                        List1.SetFocus
                        Exit Sub
                    End Select
            End If
            
            Call Clear_Field(1)
            
            If List1.ListCount > 0 Then
                List1.ListIndex = 0
                List1.SetFocus
            Else
                Text(ptxMUKE_CODE).SetFocus
            End If
        Case 3
                                            '�G���[�`�F�b�N
            sts = Err_Chk(1)
            Select Case sts
                Case False
                Case True
                    Exit Sub
                Case SYS_ERR
                    Unload Me
            End Select
            
            Beep
            yn = MsgBox("�폜���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If yn = vbYes Then
                If Update_Proc(1) Then
                    Unload Me
                End If
            End If
            
            Call Clear_Field(1)
            
            If List1.ListCount <> 0 Then
                List1.SetFocus
            Else
                Combo(pcmbC_KBN).SetFocus
            End If
        Case 7
            If List_Dsp() Then
            End If
        Case 11
            Unload Me
        Case Else
            Beep
    End Select
    
End Sub


Private Sub Command1_Click()
    
    If Command(0).Enabled = False Then
        Command(0).Enabled = True
        Upd_Mode = 0
    Else
        Command(0).Enabled = False
        Upd_Mode = 1
    End If

    If List_Dsp() Then
        Unload Me
    End If

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
Dim c       As String * 128
Dim com     As Integer
Dim sts     As Integer
Dim yn      As Integer
Dim RetBuf  As String * 255

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
                                '�o�׃��O�t�@�C������荞��
    If SYUKA_LOGF_GET_PROC() Then
        Beep
        MsgBox "�o�׃��O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    
    P_ID = StrConv(App.EXEName, vbUpperCase)
    
    If GetComputerNameA(RetBuf, 255) <> 0 Then
        WS_ID = Trim(Left(RetBuf, InStr(RetBuf, vbNullChar) - 1))
    Else
        WS_ID = "???"
    End If
                                '���ƕ���荞��
    If JGYOB_TB_Set Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If

    For i = 0 To UBound(JGYOBU_T)
'        If JGYOBU_T(i).Code = " " Then
'            Unload SubMenu(i)
'            Exit For
'        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F1030851.Caption = "���o�b�����@�o�ח\�胁���e�i���X�i" + RTrim(JGYOBU_T(i).NAME) + ")" & LAST_UPDATE_DAY
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)


                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '������Ǘ��}�X�^�n�o�d�m
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '�o�ח\��(νĲҰ��)�f�[�^�t�@�C���n�o�d�m
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                '�o�ח\��(νĲҰ��)�f�[�^�t�@�C���n�o�d�m
    If Y_SYU_H_Open(BtOpenNomal) Then
        Unload Me
    End If
                                

                                
    Upd_Mode = 0
                                '��ʏ����ݒ�
                                    '�����敪
    Combo(pcmbC_KBN).Clear
    Combo(pcmbC_KBN).AddItem CYU_KBN_1 & "   " & CYU_KBN_TUK
    Combo(pcmbC_KBN).AddItem CYU_KBN_2 & "   " & CYU_KBN_SPO
    Combo(pcmbC_KBN).AddItem CYU_KBN_3 & "   " & CYU_KBN_HJU
'    Combo(pcmbC_KBN).AddItem CYU_KBN_4
    Combo(pcmbC_KBN).AddItem CYU_KBN_E & "   " & CYU_KBN_BOU
    Combo(pcmbC_KBN).AddItem CYU_KBN_T & "   " & CYU_KBN_KIN
    
                                '���O�t�@�C������荞��
    If GetIni(App.EXEName, "CYU_KBN", "SYS", c) Then
        Beep
        MsgBox "�����敪�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    
    For i = 0 To Combo(pcmbC_KBN).ListCount - 1
        If RTrim(c) = Right(Combo(pcmbC_KBN).List(i), 1) Then
            Combo(pcmbC_KBN).ListIndex = i
            Exit For
        End If
    Next i
    
    If MTS_Set_Proc() Then
        Unload Me
    End If
    
    Call Clear_Field(0)
        
    Text(ptxMUKE_CODE).SetFocus

    End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
    sts = P_Off
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '������Ǘ��}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "������Ǘ��}�X�^")
        End If
    End If
                                            '�o�ח\��f�[�^�t�@�C���b�k�n�r�d
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�o�ח\��f�[�^�t�@�C��")
        End If
    End If
                                            '�o�ח\��f�[�^(νĲҰ��)�t�@�C���b�k�n�r�d
    sts = BTRV(BtOpClose, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�o�ח\��f�[�^(νĲҰ��)�t�@�C��")
        End If
    End If
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1030851 = Nothing
    
    End
End Sub

Private Sub List1_DblClick()
Dim sts As Integer

    sts = Item_Dsp()
    Select Case sts
        Case False
        Case SYS_CANCEL
        Case Else
            Unload Me
    End Select
    
End Sub


Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)

Dim sts As Integer
    
    If List1.ListCount = 0 Then
        Exit Sub
    End If
        
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    sts = Item_Dsp()
    
    Select Case sts
        Case False
        Case SYS_CANCEL
        Case Else
            Unload Me
    End Select

End Sub


Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    '���j���[���I���v��
'    If JGYOBU_T(Index).CODE = " " Then
'        Unload Me
'    End If

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '���ƕ��؂�ւ�
    F1030851.Caption = "���o�b�����@�o�ח\�胁���e�i���X�i" + RTrim(JGYOBU_T(Index).NAME) + "�j" & LAST_UPDATE_DAY
    Last_JGYOBU = JGYOBU_T(Index).CODE
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)

End Sub
Private Sub Text_GotFocus(Index As Integer)
    If Text(Index).TabStop = True Then
        Text(Index) = Trim(Text(Index).Text)
        Text(Index).SelStart = 0
        Text(Index).SelLength = Len(Text(Index).Text)
    End If
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim RetBuf As String
Dim i As Integer
Dim sts As Integer
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    Select Case Index
        
        Case ptxMUKE_CODE
            Call UniCode_Conv(K2_MTS.MUKE_CODE, Text(Index).Text)
            sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K2_MTS, Len(K2_MTS), 2)
            Select Case sts
                Case BtNoErr
                    If Len(Trim(StrConv(MTSREC.SS_CODE, vbUnicode))) <> 0 Then
                        Beep
                        MsgBox "���͂������ڂ̓G���[�ł��B(������R�[�h)"
                        Exit Sub
                    End If
                                
                Case BtErrKeyNotFound
                                
                    Call UniCode_Conv(K3_MTS.SS_CODE, Text(Index).Text)
                                                        
                    sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K3_MTS, Len(K3_MTS), 3)
                    Select Case sts
                        Case BtNoErr
                                        
                        Case BtErrKeyNotFound
                            Beep
                            MsgBox "���͂������ڂ̓G���[�ł��B(������R�[�h)"
                            Exit Sub
                                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "������Ǘ��}�X�^")
                            Unload Me
                    End Select

                Case Else
                    Call File_Error(sts, BtOpGetEqual, "������Ǘ��}�X�^")
                    Unload Me
            End Select


            For i = 0 To Combo(pcmbMUKE_CODE).ListCount - 1 '������
    
                If Right(Combo(pcmbMUKE_CODE).List(i), 16) = StrConv(MTSREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.SS_CODE, vbUnicode) Then
                    Combo(pcmbMUKE_CODE).ListIndex = i
                    Exit For
                End If
            
    
            Next

            If List_Dsp() Then
                Unload Me
            End If
        
            If List1.ListCount > 0 Then
                List1.ListIndex = 0
                List1.SetFocus
            End If
        
        Case Else
        
    End Select
    For i = Index + 1 To Text_Max
        If Text(i).Visible And Text(i).Enabled And Text(i).TabStop Then
            Text(i).SetFocus
            Call Text_GotFocus(Index)
            Exit Sub
        End If
    Next i
End Sub
