VERSION 5.00
Begin VB.Form F1100801 
   BackColor       =   &H00C0C0C0&
   Caption         =   "�o�n�r�i�ڃ}�X�^�[�ϊ�����"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   345
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
   ScaleHeight     =   6495
   ScaleWidth      =   10095
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.CommandButton Command1 
      Caption         =   "�}�X�^�ϊ�"
      Height          =   555
      Index           =   1
      Left            =   8040
      TabIndex        =   4
      Top             =   5880
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label Lab_File 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4620
      TabIndex        =   7
      Top             =   3300
      Width           =   120
   End
   Begin VB.Label Lab 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�i�ڃ}�X�^�i�����݁j��"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   1980
      TabIndex        =   6
      Top             =   3960
      Width           =   2640
   End
   Begin VB.Label Lab 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�o�m�f�[�^�i�Ǎ��݁j��"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   1980
      TabIndex        =   5
      Top             =   3000
      Width           =   2640
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '�E����
      BackStyle       =   0  '����
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   4740
      TabIndex        =   2
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '�E����
      BackStyle       =   0  '����
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   4740
      TabIndex        =   3
      Top             =   3960
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
      Left            =   840
      TabIndex        =   1
      Top             =   2160
      Width           =   240
   End
   Begin VB.Label MsgLab 
      Alignment       =   2  '��������
      BackColor       =   &H008080FF&
      Caption         =   "�o�n�r�i�ڃ}�X�^�[�ϊ�"
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
      Left            =   1920
      TabIndex        =   0
      Top             =   0
      Width           =   6495
   End
End
Attribute VB_Name = "F1100801"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function Convert_Proc() As Integer
Dim IN_Count        As Long
Dim UP_Count        As Long
Dim DISP_INTERVAL   As Long

Dim sts             As Integer
Dim i               As Integer
Dim c               As String

Dim IN_Path         As String       '����̧�� �߽
Dim IN_Fil          As Variant      '����̧�ٖ��z��


    Convert_Proc = True


'---------------------------------------------  �o�m�t�@�C���p�X��荞��
    If GetIni("FILE", "PN_PATH", "F110080", c) Then
        c = "[F110080] �o�m�t�@�C���p�X �擾�G���[(ini)"
        Call Log_Out(LOG_F, c)
        Exit Function
    End If
    IN_Path = Trim(c)

'---------------------------------------------  �o�m�t�@�C������荞��
    If GetIni("FILE", "PN_FILES", "F110080", c) Then
        c = "[F110080] �o�m�t�@�C���� �擾�G���[(ini)"
        Call Log_Out(LOG_F, c)
        Exit Function
    End If
    IN_Fil = Split(Trim(c), ",", -1)


'---------------------------------------------  �i�ڃ}�X�^�ϊ�
    MsgLab(1) = "�o�n�r�i�ڃ}�X�^�ϊ� �������I�I"
    Me.MousePointer = vbHourglass
    UP_Count = 0

    On Error Resume Next

    For i = 0 To UBound(IN_Fil)

        Open IN_Path & IN_Fil(i) For Binary As #1

        Lab_File.Caption = "(" & IN_Fil(i) & ")"
        IN_Count = 0
        DISP_INTERVAL = 0

        Do
            Get #1, , INREC

            If EOF(1) Then
                Exit Do
            End If

            IN_Count = IN_Count + 1
            DISP_INTERVAL = DISP_INTERVAL + 1
            If DISP_INTERVAL = 100 Then
                Cnt(0).Caption = Format(IN_Count, "#0")
                DISP_INTERVAL = 0
            End If
            DoEvents


            If Trim(StrConv(INREC.HIN_GAI, vbUnicode)) <> "" And _
              (Trim(StrConv(INREC.HIN_NAME, vbUnicode)) <> "" Or _
               Trim(StrConv(INREC.U_TANKA2, vbUnicode)) <> "" Or _
               Trim(StrConv(INREC.U_TANKA3, vbUnicode)) <> "" Or _
               Trim(StrConv(INREC.U_TANKA4, vbUnicode)) <> "" Or _
               Trim(StrConv(INREC.HIN_NAI, vbUnicode)) <> "" Or _
               Trim(StrConv(INREC.KOSO_CD, vbUnicode)) <> "" Or _
               Trim(StrConv(INREC.UNIT_BUHIN, vbUnicode)) <> "" Or _
               Trim(StrConv(INREC.NAI_BUHIN, vbUnicode)) <> "" Or _
               Trim(StrConv(INREC.GAI_BUHIN, vbUnicode)) <> "" Or _
               Trim(StrConv(INREC.HYO_TANKA, vbUnicode)) <> "") Then
                                                                        '���ƕ� �Ǒւ�
                If GetIni("JIGYOBU", StrConv(INREC.SISAN_JGYOBA, vbUnicode), "F110080", c) Then
                    Beep
                    c = "���ƕ� �Ǒւ��G���[(�i�ځF" & StrConv(INREC.HIN_GAI, vbUnicode) & _
                                       "�A���Ə�F" & StrConv(INREC.JGYOBA, vbUnicode) & ")"
                    Call Log_Out(LOG_F, c)
                Else
                    Call UniCode_Conv(K0_ITEM.JGYOBU, Trim(c))              '���ƕ�
                    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)           '�����O�i������ �Œ�j
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, _
                                        StrConv(INREC.HIN_GAI, vbUnicode))  '�i��
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    If sts = BtNoErr Then
'''                        If Trim(StrConv(INREC.HIN_NAME, vbUnicode)) <> "" Then  '�i��
'''                            Call UniCode_Conv(ITEMREC.HIN_NAME, Trim(StrConv(INREC.HIN_NAME, vbUnicode)))
'''                        End If
                        
                        If Trim(StrConv(INREC.HIN_NAME, vbUnicode)) <> "" Then  '�i��
                            If Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode)) = "" Then  '�i��
                                Call UniCode_Conv(ITEMREC.HIN_NAME, Trim(StrConv(INREC.HIN_NAME, vbUnicode)))
                            End If
                        End If
                        
                        
                        If Trim(StrConv(INREC.U_TANKA2, vbUnicode)) <> "" Then  '���i(1)
                            
                            If IsNumeric(StrConv(INREC.U_TANKA2, vbUnicode)) Then
                                If Val(StrConv(INREC.U_TANKA2, vbUnicode)) <> 0 Then
                                    Call UniCode_Conv(ITEMREC.L_URIKIN1, Format(Int(Trim(StrConv(INREC.U_TANKA2, vbUnicode))), "0000000000"))
                                End If
                            End If
                        
                        End If
                        
                        If Trim(StrConv(INREC.U_TANKA3, vbUnicode)) <> "" Then  '���i(2)
                            
                            If IsNumeric(StrConv(INREC.U_TANKA3, vbUnicode)) Then
                                If Val(StrConv(INREC.U_TANKA3, vbUnicode)) <> 0 Then
                                    Call UniCode_Conv(ITEMREC.L_URIKIN2, Format(Int(Trim(StrConv(INREC.U_TANKA3, vbUnicode))), "0000000000"))
                                End If
                            End If
                        
                        End If
                        
                        If Trim(StrConv(INREC.U_TANKA4, vbUnicode)) <> "" Then  '���i(3)
                            
                            If IsNumeric(StrConv(INREC.U_TANKA4, vbUnicode)) Then
                                
                                If Val(StrConv(INREC.U_TANKA4, vbUnicode)) <> 0 Then
                            
                                    Call UniCode_Conv(ITEMREC.L_URIKIN3, Format(Int(Trim(StrConv(INREC.U_TANKA4, vbUnicode))), "0000000000"))
                                End If
                        
                            End If
                        End If
                        
                        If Trim(StrConv(INREC.HIN_NAI, vbUnicode)) <> "" Then   '�i��(����)
                            Call UniCode_Conv(ITEMREC.HIN_NAI, Trim(StrConv(INREC.HIN_NAI, vbUnicode)))
                        End If

                                                                                '���`��
                        If Trim(StrConv(INREC.KOSO_CD, vbUnicode)) <> "" Then
                            Call UniCode_Conv(ITEMREC.K_KEITAI, StrConv(INREC.KOSO_CD, vbUnicode))
                        End If
                                                                                '�Ưĕ��i�敪
                        If Trim(StrConv(INREC.UNIT_BUHIN, vbUnicode)) <> "" Then
                            Call UniCode_Conv(ITEMREC.UNIT_BUHIN, StrConv(INREC.UNIT_BUHIN, vbUnicode))
                        End If
                                                                                '�����������i�敪
                        If Trim(StrConv(INREC.NAI_BUHIN, vbUnicode)) <> "" Then
                            Call UniCode_Conv(ITEMREC.NAI_BUHIN, StrConv(INREC.NAI_BUHIN, vbUnicode))
                        End If
                                                                                '�C�O�������i�敪
                        If Trim(StrConv(INREC.GAI_BUHIN, vbUnicode)) <> "" Then
                            Call UniCode_Conv(ITEMREC.GAI_BUHIN, StrConv(INREC.GAI_BUHIN, vbUnicode))
                        End If
                                                                                '�C�O�������i�敪
                        If Trim(StrConv(INREC.GAI_BUHIN, vbUnicode)) <> "" Then
                            Call UniCode_Conv(ITEMREC.GAI_BUHIN, StrConv(INREC.GAI_BUHIN, vbUnicode))
                        End If
                                                                                '�W���P��
                        If Trim(StrConv(INREC.HYO_TANKA, vbUnicode)) <> "" Then
                            Call UniCode_Conv(ITEMREC.HYO_TANKA, StrConv(INREC.HYO_TANKA, vbUnicode))
                        End If
                                                                            


                        sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts = BtNoErr Then
                            UP_Count = UP_Count + 1
                            Cnt(1).Caption = Format(UP_Count, "#0")
                            DoEvents
                        Else
                            c = "�i�ڃ}�X�^ ERROR Opretion = " & BtOpUpdate & " " & "sts= " & sts
                            Call Log_Out(LOG_F, c)
                        End If
                    
                    
                    Else
'''                        c = "�i�ڃ}�X�^���o�^ [�i�ځF" & StrConv(INREC.HIN_GAI, vbUnicode) & _
'''                                        "�A���Ə�F" & StrConv(INREC.JGYOBA, vbUnicode) & "]"
'''                        Call Log_Out(LOG_F, c)
                    End If
                End If
            End If
        Loop

        Close #1

    Next i


    Convert_Proc = False

End Function

Private Sub Command1_Click(Index As Integer)
Dim ans     As Integer
Dim i       As Integer

    Select Case Index

        Case 1      'Ͻ��ϊ��J�n
            Command1(1).Enabled = False
            DoEvents

            If Convert_Proc() Then
                Unload Me
            End If
            Unload Me

    End Select

End Sub

Private Sub Form_Activate()
    If Command1(1).Enabled = True Then
        Command1(1).Value = True
    End If
End Sub

Private Sub Form_DblClick()
'    PrintForm
End Sub

Private Sub Form_Load()
Dim sts     As Integer
Dim c       As String

    If App.PrevInstance Then
        c = "[F110080] ����v���O�������s���ׁ̈A�����𒆎~���܂����B"
        Call Log_Out(LOG_F, c)
        End
    End If

    Show
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        c = "[F110080] ���O�t�@�C�����̊l���Ɏ��s�����ׁA�����𒆎~���܂����B"
        Call Log_Out(LOG_F, c)
        End
    End If
    LOG_F = RTrim(c)

                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts     As Integer
Dim c       As String

                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            c = "�i�ڃ}�X�^ ERROR Opretion = " & BtOpClose & " " & "sts= " & sts
            Call Log_Out(LOG_F, c)
        End If
    End If

    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)

    Set F1100801 = Nothing

    End
End Sub
