VERSION 5.00
Begin VB.Form PI000104 
   Appearance      =   0  '�ׯ�
   Caption         =   "�\���R�s�[�i�ԓ��͉��"
   ClientHeight    =   2610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6840
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
   ScaleHeight     =   2610
   ScaleWidth      =   6840
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.CommandButton Command1 
      Caption         =   "�L�����Z��"
      Height          =   615
      Index           =   1
      Left            =   4320
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "  �m�@�@��@"
      Height          =   615
      Index           =   0
      Left            =   1920
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   2280
      MaxLength       =   20
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "<2019.05.28>"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   1260
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Caption         =   "�R�s�[��i��"
      Height          =   255
      Index           =   14
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "PI000104"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const ptxHIN_GAI% = 0
Private Const ptxHIN_NAME% = 1

Private Sub Command1_Click(Index As Integer)
    
    
    PI000104_CANCEL_F = 0
    
    Select Case Index
        
        Case 0
            
            PI000104_HIN_GAI = ""
            
            If Error_Check_Proc(ptxHIN_GAI) Then
                Text1(ptxHIN_GAI).SetFocus
                Exit Sub
            End If
            
            
            PI000104_HIN_GAI = Text1(ptxHIN_GAI).text
            Text1(ptxHIN_GAI).text = ""
            
            PI000104.Visible = False
        Case 1
            PI000104_CANCEL_F = 1
            PI000104_HIN_GAI = ""
            PI000104.Visible = False
    End Select
End Sub


Private Function Error_Check_Proc(Index As Integer) As Integer
'----------------------------------------------------------------------------
'                   ERROR CHECK����
'----------------------------------------------------------------------------
    
Dim sts     As Integer
Dim wkTanto As String * 5
   
'2019.05.28 �ǉ��i����j
Dim WS01    As Integer
Dim W_COMBO As String       '�e��ʂ̃R���{���e

    Error_Check_Proc = True


    Select Case Index
        Case ptxHIN_GAI
    
            Text1(ptxHIN_GAI).text = StrConv(Text1(ptxHIN_GAI).text, vbUpperCase)
            If Trim(Text1(ptxHIN_GAI).text) = "" Then
                MsgBox "�i�ڃR�[�h�͕K�����͂��ĉ������B"
                Exit Function
            End If
    
    
    
    
            W_COMBO = Right(PI000101.Combo1(0), 4)
            Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(PI000101.Combo1(0), 4), 3, 1))
            Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(PI000101.Combo1(0), 4), 4, 1))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).text)



            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                
                Case BtErrKeyNotFound

                    wkTanto = PI000101.Text1(2).text
                    If Trim(wkTanto) = "" Then
                        wkTanto = "PSHIJ"
                    End If

                    Last_JGYOBU = Mid(Right(PI000101.Combo1(0), 4), 3, 1)
                    If PN_CHK(Text1(Index), "G", wkTanto, 1) Then          '�O���i��
                        Exit Function
                    End If

                Case Else
                    
                    
                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                    Exit Function

            End Select

    
            '2019.05.28 �\���l�̑��݃`�F�b�N��ǉ��i����j
            '           �w��i�Ԃ��V�K�̎��̍\���R�s�[����{�d�l�����݂���ꍇ�̓G���[�I
            W_COMBO = Right(PI000101.Combo1(0), 4)
            Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(PI000101.Combo1(0), 4), 1, 2)) '�e��ʓ��e
            Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(PI000101.Combo1(0), 4), 3, 1))       ' �V
            Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(PI000101.Combo1(0), 4), 4, 1))       '�@�V
            
            Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHIN_GAI).text)                       '���Y��ʂ̓��͒l
            '------------------------------------------ ��/�O��/�����E�\��
'            Public Const P_HEAD$ = "0"                  'ͯ�ް
'            Public Const P_KOSOU$ = "1"                 '������
'            Public Const P_GAISOU$ = "2"                '�O������
'            Public Const P_DOUKON$ = "3"                '�����E�\��
            Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_KOSOU)                                     '�Œ�l�i�����i�j
            Call UniCode_Conv(K0_P_COMPO.SEQNO, "")                                          '�@�V
            
            Do
                sts = BTRV(BtOpGetGreater, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrEOF
                        Exit Do
'                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                        Sleep (500)
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                        Exit Function
                End Select
            Loop
            WS01 = 1
            If sts = BtNoErr Then
                If StrConv(P_COMPO_O_REC.SHIMUKE_CODE, vbUnicode) <> Mid(W_COMBO, 1, 2) Then WS01 = 0
                If StrConv(P_COMPO_O_REC.JGYOBU, vbUnicode) <> Mid(W_COMBO, 3, 1) Then WS01 = 0
                If StrConv(P_COMPO_O_REC.NAIGAI, vbUnicode) <> Mid(W_COMBO, 4, 1) Then WS01 = 0
                If Trim(StrConv(P_COMPO_O_REC.HIN_GAI, vbUnicode)) <> Trim(Text1(ptxHIN_GAI).text) Then WS01 = 0
                If Trim(StrConv(P_COMPO_O_REC.DATA_KBN, vbUnicode)) > P_DOUKON Then WS01 = 0        '�������E�\���H
            Else
                WS01 = 0
            End If
            If WS01 = 1 Then
                MsgBox "�w�肵���i�ڃR�[�h�̍\���́A�o�^�ς݂ł��B"
                Exit Function
            End If
    
    End Select



    Error_Check_Proc = False
End Function
