VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form F1060601 
   BackColor       =   &H00FFFFFF&
   Caption         =   "作業監視モニター"
   ClientHeight    =   7275
   ClientLeft      =   2025
   ClientTop       =   2550
   ClientWidth     =   11250
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   11250
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000B&
      Height          =   360
      Index           =   9
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5880
      Width           =   372
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000B&
      Height          =   360
      Index           =   8
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5880
      Width           =   372
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000B&
      Height          =   360
      Index           =   7
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5880
      Width           =   372
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000B&
      Height          =   360
      Index           =   6
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5880
      Width           =   372
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000B&
      Height          =   360
      Index           =   5
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5880
      Width           =   732
   End
   Begin VB.CommandButton Command 
      Caption         =   "終  了"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
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
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
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
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "最　新"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
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
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
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
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
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
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
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
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
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
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
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
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6360
      Width           =   855
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Height          =   5652
      Left            =   840
      OleObjectBlob   =   "F1060601.frx":0000
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   120
      Width           =   9012
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "分現在"
      Height          =   252
      Index           =   10
      Left            =   9120
      TabIndex        =   17
      Top             =   6000
      Width           =   732
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "時"
      Height          =   252
      Index           =   9
      Left            =   8280
      TabIndex        =   16
      Top             =   6000
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "日"
      Height          =   252
      Index           =   8
      Left            =   7320
      TabIndex        =   15
      Top             =   6000
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "月"
      Height          =   252
      Index           =   7
      Left            =   6480
      TabIndex        =   14
      Top             =   6000
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "年"
      Height          =   252
      Index           =   6
      Left            =   5640
      TabIndex        =   13
      Top             =   6000
      Width           =   252
   End
   Begin VB.Label LabJIGYO 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Left            =   120
      TabIndex        =   12
      Top             =   6840
      Width           =   180
   End
   Begin VB.Menu MainMenu 
      Caption         =   "事業部"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1060601"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const ptxDATE_YY% = 5           '現在　年
Private Const ptxDATE_MM% = 6           '現在　月
Private Const ptxDATE_DD% = 7           '現在　日
Private Const ptxTIME_HH% = 8           '現在　時
Private Const ptxTIME_MM% = 9           '現在　分

Dim Y_SYUKA     As New XArrayDB

Private Const Min_Row% = 1              '最小行数
'Private Const Max_Row& = 2000           '最大行数

Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 5              '最大列数

Private Const ColMUKE_NAME% = 0         '列 向け先名称
Private Const ColALL_Su% = 1            '列 合計
Private Const ColTUK_Su% = 2            '列 月切り
Private Const ColSPO_Su% = 3            '列 ｽﾎﾟｯﾄ
Private Const ColHJU_Su% = 4            '列 補充
Private Const ColBOU_Su% = 5            '列 貿易

Private Const RowTotal% = 1             '行 合計

Private Type Set_Tbl_Tag                '集計用中間テーブル
    MUKE_CODE   As String * 8              '向け先コード
    SS_CODE     As String * 8
    MUKE_NAME   As String * 10              '向け先名称
    ALL_SU_JI   As Integer                  '全数　済
    ALL_SU_YO   As Integer                  '全数　予定
    
    TUK_SU_JI   As Integer                  '月切り　済
    TUK_SU_YO   As Integer                  '月切り　予定
    
    SPO_SU_JI   As Integer                  'スポット　済
    SPO_SU_YO   As Integer                  'スポット　予定
    
    HJU_SU_JI   As Integer                  'スポット　済
    HJU_SU_YO   As Integer                  'スポット　予定
    
    BOU_SU_JI   As Integer                  'スポット　済
    BOU_SU_YO   As Integer                  'スポット　予定
End Type

Private Set_tbl_MTS()   As Set_Tbl_Tag  'ＭＴＳ用集計テーブル
Private MTS_NON         As Integer

Private Set_tbl_KEN()   As Set_Tbl_Tag  '検査工場用集計テーブル
Private KEN_NON         As Integer

Private Set_tbl_BOU()   As Set_Tbl_Tag  '貿易用集計テーブル
Private BOU_NON         As Integer




Private Function List_Dsp_Proc() As Integer
    
Dim com             As Integer
Dim sts             As Integer
Dim i               As Integer

Dim Row             As Integer
    
Dim GK_ALL_SU_JI    As Integer
Dim GK_ALL_SU_YO    As Integer
Dim GK_TUK_SU_JI    As Integer
Dim GK_TUK_SU_YO    As Integer
Dim GK_SPO_SU_JI    As Integer
Dim GK_SPO_SU_YO    As Integer
Dim GK_HJU_SU_JI    As Integer
Dim GK_HJU_SU_YO    As Integer
Dim GK_BOU_SU_JI    As Integer
Dim GK_BOU_SU_YO    As Integer
    
    List_Dsp_Proc = True
    
    Call Input_Lock
                                    '中間テーブルクリア
    If Not MTS_NON Then
        For i = 0 To UBound(Set_tbl_MTS)
                    
            Set_tbl_MTS(i).ALL_SU_JI = 0
            Set_tbl_MTS(i).ALL_SU_YO = 0
    
            Set_tbl_MTS(i).TUK_SU_JI = 0
            Set_tbl_MTS(i).TUK_SU_YO = 0
    
            Set_tbl_MTS(i).SPO_SU_JI = 0
            Set_tbl_MTS(i).SPO_SU_YO = 0
    
            Set_tbl_MTS(i).HJU_SU_JI = 0
            Set_tbl_MTS(i).HJU_SU_YO = 0
            
            Set_tbl_MTS(i).BOU_SU_JI = 0
            Set_tbl_MTS(i).BOU_SU_YO = 0
        
        
        Next i
    End If
    
    If Not KEN_NON Then
        For i = 0 To UBound(Set_tbl_KEN)
                    
            Set_tbl_KEN(i).ALL_SU_JI = 0
            Set_tbl_KEN(i).ALL_SU_YO = 0
    
            Set_tbl_KEN(i).TUK_SU_JI = 0
            Set_tbl_KEN(i).TUK_SU_YO = 0
    
            Set_tbl_KEN(i).SPO_SU_JI = 0
            Set_tbl_KEN(i).SPO_SU_YO = 0
    
            Set_tbl_KEN(i).HJU_SU_JI = 0
            Set_tbl_KEN(i).HJU_SU_YO = 0
            
            Set_tbl_KEN(i).BOU_SU_JI = 0
            Set_tbl_KEN(i).BOU_SU_YO = 0
        
        
        Next i
    End If
    
    If Not BOU_NON Then
        For i = 0 To UBound(Set_tbl_BOU)
                    
            Set_tbl_BOU(i).ALL_SU_JI = 0
            Set_tbl_BOU(i).ALL_SU_YO = 0
    
            Set_tbl_BOU(i).TUK_SU_JI = 0
            Set_tbl_BOU(i).TUK_SU_YO = 0
    
            Set_tbl_BOU(i).SPO_SU_JI = 0
            Set_tbl_BOU(i).SPO_SU_YO = 0
    
            Set_tbl_BOU(i).HJU_SU_JI = 0
            Set_tbl_BOU(i).HJU_SU_YO = 0
            
            Set_tbl_BOU(i).BOU_SU_JI = 0
            Set_tbl_BOU(i).BOU_SU_YO = 0
        
        
        Next i
    End If
                                    
    GK_ALL_SU_JI = 0
    GK_ALL_SU_YO = 0
    GK_TUK_SU_JI = 0
    GK_TUK_SU_YO = 0
    GK_SPO_SU_JI = 0
    GK_SPO_SU_YO = 0
    GK_HJU_SU_JI = 0
    GK_HJU_SU_YO = 0
    GK_BOU_SU_JI = 0
    GK_BOU_SU_YO = 0
    
    If Not MTS_NON Then
        For i = 0 To UBound(Set_tbl_MTS) - 1
                                       'ＭＴＳ分集計処理開始
            Call UniCode_Conv(K3_Y_SYU.JGYOBU, Last_JGYOBU)
            Call UniCode_Conv(K3_Y_SYU.KEY_MUKE_CODE, Set_tbl_MTS(i).MUKE_CODE)
            Call UniCode_Conv(K3_Y_SYU.KEY_SS_CODE, Set_tbl_MTS(i).SS_CODE)
            Call UniCode_Conv(K3_Y_SYU.KEY_CYU_KBN, "")
            Call UniCode_Conv(K3_Y_SYU.NAIGAI, "")
            Call UniCode_Conv(K3_Y_SYU.KEY_HIN_NO, "")
            Call UniCode_Conv(K3_Y_SYU.KEY_ID_NO, "")
                                                            
            com = BtOpGetGreater
                                    
            Do
            
                DoEvents
        
                sts = BTRV(com, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K3_Y_SYU, Len(K3_Y_SYU), 3)
                Select Case sts
                    Case BtNoErr
                                            '事業部／向け先ブレーク
                        If StrConv(Y_SYUREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                            Exit Do
                        End If
            
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        Call File_Error(sts, com, "出荷予定データ")
                        List_Dsp_Proc = SYS_ERR
                        Exit Function
            
                End Select
                                        
                
                If Set_tbl_MTS(i).MUKE_CODE = StrConv(Y_SYUREC.MUKE_CODE, vbUnicode) Then
                
                
                    Set_tbl_MTS(i).ALL_SU_YO = Set_tbl_MTS(i).ALL_SU_YO + 1
                    Set_tbl_MTS(UBound(Set_tbl_MTS)).ALL_SU_YO = Set_tbl_MTS(UBound(Set_tbl_MTS)).ALL_SU_YO + 1
                    GK_ALL_SU_YO = GK_ALL_SU_YO + 1
                                                '検品済なら（？？？）
                    If Len(Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode))) <> 0 Then
                        Set_tbl_MTS(i).ALL_SU_JI = Set_tbl_MTS(i).ALL_SU_JI + 1
                        Set_tbl_MTS(UBound(Set_tbl_MTS)).ALL_SU_JI = Set_tbl_MTS(UBound(Set_tbl_MTS)).ALL_SU_JI + 1
                        GK_ALL_SU_JI = GK_ALL_SU_JI + 1
                    End If
                                            
                    Select Case StrConv(Y_SYUREC.CYU_KBN, vbUnicode)
                        Case CYU_KBN_TUK            '月切
                            Set_tbl_MTS(i).TUK_SU_YO = Set_tbl_MTS(i).TUK_SU_YO + 1
                            Set_tbl_MTS(UBound(Set_tbl_MTS)).TUK_SU_YO = Set_tbl_MTS(UBound(Set_tbl_MTS)).TUK_SU_YO + 1
                            GK_TUK_SU_YO = GK_TUK_SU_YO + 1
                                                '検品済なら（？？？）
                            If Len(Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode))) <> 0 Then
                                Set_tbl_MTS(i).TUK_SU_JI = Set_tbl_MTS(i).TUK_SU_JI + 1
                                Set_tbl_MTS(UBound(Set_tbl_MTS)).TUK_SU_JI = Set_tbl_MTS(UBound(Set_tbl_MTS)).TUK_SU_JI + 1
                                GK_TUK_SU_JI = GK_TUK_SU_JI + 1
                            End If
                    
                        Case CYU_KBN_SPO            'スポット
                            Set_tbl_MTS(i).SPO_SU_YO = Set_tbl_MTS(i).SPO_SU_YO + 1
                            Set_tbl_MTS(UBound(Set_tbl_MTS)).SPO_SU_YO = Set_tbl_MTS(UBound(Set_tbl_MTS)).SPO_SU_YO + 1
                            GK_SPO_SU_YO = GK_SPO_SU_YO + 1
                                                '検品済なら（？？？）
                            If Len(Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode))) <> 0 Then
                                Set_tbl_MTS(i).SPO_SU_JI = Set_tbl_MTS(i).SPO_SU_JI + 1
                                Set_tbl_MTS(UBound(Set_tbl_MTS)).SPO_SU_JI = Set_tbl_MTS(UBound(Set_tbl_MTS)).SPO_SU_JI + 1
                                GK_SPO_SU_JI = GK_SPO_SU_JI + 1
                            End If
                        Case CYU_KBN_HJU            '補充
                            Set_tbl_MTS(i).HJU_SU_YO = Set_tbl_MTS(i).HJU_SU_YO + 1
                            Set_tbl_MTS(UBound(Set_tbl_MTS)).HJU_SU_YO = Set_tbl_MTS(UBound(Set_tbl_MTS)).HJU_SU_YO + 1
                            GK_HJU_SU_YO = GK_HJU_SU_YO + 1
                                                '検品済なら（？？？）
                            If Len(Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode))) <> 0 Then
                                Set_tbl_MTS(i).HJU_SU_JI = Set_tbl_MTS(i).HJU_SU_JI + 1
                                Set_tbl_MTS(UBound(Set_tbl_MTS)).HJU_SU_JI = Set_tbl_MTS(UBound(Set_tbl_MTS)).HJU_SU_JI + 1
                                GK_HJU_SU_JI = GK_HJU_SU_JI + 1
                            End If
                        Case CYU_KBN_BOU            '貿易
                            Set_tbl_MTS(i).BOU_SU_YO = Set_tbl_MTS(i).BOU_SU_YO + 1
                            Set_tbl_MTS(UBound(Set_tbl_MTS)).BOU_SU_YO = Set_tbl_MTS(UBound(Set_tbl_MTS)).BOU_SU_YO + 1
                            GK_BOU_SU_YO = GK_BOU_SU_YO + 1
                                                '検品済なら（？？？）
                            If Len(Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode))) <> 0 Then
                                Set_tbl_MTS(i).BOU_SU_JI = Set_tbl_MTS(i).BOU_SU_JI + 1
                                Set_tbl_MTS(UBound(Set_tbl_MTS)).BOU_SU_JI = Set_tbl_MTS(UBound(Set_tbl_MTS)).BOU_SU_JI + 1
                                GK_BOU_SU_JI = GK_BOU_SU_JI + 1
                            End If
                    End Select
                End If
                
                com = BtOpGetNext
            
            Loop
        Next i
    End If
    
    If Not KEN_NON Then
        For i = 0 To UBound(Set_tbl_KEN) - 1
                                       '検査工場分集計処理開始
            Call UniCode_Conv(K3_Y_SYU.JGYOBU, Last_JGYOBU)
            Call UniCode_Conv(K3_Y_SYU.KEY_MUKE_CODE, Set_tbl_KEN(i).MUKE_CODE)
            Call UniCode_Conv(K3_Y_SYU.KEY_SS_CODE, Set_tbl_KEN(i).SS_CODE)
            Call UniCode_Conv(K3_Y_SYU.KEY_CYU_KBN, "")
            Call UniCode_Conv(K3_Y_SYU.NAIGAI, "")
            Call UniCode_Conv(K3_Y_SYU.KEY_HIN_NO, "")
            Call UniCode_Conv(K3_Y_SYU.KEY_ID_NO, "")
                                                            
            com = BtOpGetGreater
                                    
            Do
            
                DoEvents
        
                sts = BTRV(com, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K3_Y_SYU, Len(K3_Y_SYU), 3)
                Select Case sts
                    Case BtNoErr
                                            '事業部／向け先ブレーク
                        If StrConv(Y_SYUREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                            Exit Do
                        End If
            
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        Call File_Error(sts, com, "出荷予定データ")
                        List_Dsp_Proc = SYS_ERR
                        Exit Function
            
                End Select
                                        
                
                If Set_tbl_KEN(i).MUKE_CODE = StrConv(Y_SYUREC.MUKE_CODE, vbUnicode) Then
                    Set_tbl_KEN(i).ALL_SU_YO = Set_tbl_KEN(i).ALL_SU_YO + 1
                    Set_tbl_KEN(UBound(Set_tbl_KEN)).ALL_SU_YO = Set_tbl_KEN(UBound(Set_tbl_KEN)).ALL_SU_YO + 1
                    GK_ALL_SU_YO = GK_ALL_SU_YO + 1
                                                '検品済なら（？？？）
                    If Len(Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode))) <> 0 Then
                        Set_tbl_KEN(i).ALL_SU_JI = Set_tbl_KEN(i).ALL_SU_JI + 1
                        Set_tbl_KEN(UBound(Set_tbl_KEN)).ALL_SU_JI = Set_tbl_KEN(UBound(Set_tbl_KEN)).ALL_SU_JI + 1
                        GK_ALL_SU_JI = GK_ALL_SU_JI + 1
                    End If
                                            
                    Select Case StrConv(Y_SYUREC.CYU_KBN, vbUnicode)
                        Case CYU_KBN_TUK            '月切
                            Set_tbl_KEN(i).TUK_SU_YO = Set_tbl_KEN(i).TUK_SU_YO + 1
                            Set_tbl_KEN(UBound(Set_tbl_KEN)).TUK_SU_YO = Set_tbl_KEN(UBound(Set_tbl_KEN)).TUK_SU_YO + 1
                            GK_TUK_SU_YO = GK_TUK_SU_YO + 1
                                                '検品済なら（？？？）
                            If Len(Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode))) <> 0 Then
                                Set_tbl_KEN(i).TUK_SU_JI = Set_tbl_KEN(i).TUK_SU_JI + 1
                                Set_tbl_KEN(UBound(Set_tbl_KEN)).TUK_SU_JI = Set_tbl_KEN(UBound(Set_tbl_KEN)).TUK_SU_JI + 1
                                GK_TUK_SU_JI = GK_TUK_SU_JI + 1
                            End If
                    
                        Case CYU_KBN_SPO            'スポット
                            Set_tbl_KEN(i).SPO_SU_YO = Set_tbl_KEN(i).SPO_SU_YO + 1
                            Set_tbl_KEN(UBound(Set_tbl_KEN)).SPO_SU_YO = Set_tbl_KEN(UBound(Set_tbl_KEN)).SPO_SU_YO + 1
                            GK_SPO_SU_YO = GK_SPO_SU_YO + 1
                                                '検品済なら（？？？）
                            If Len(Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode))) <> 0 Then
                                Set_tbl_KEN(i).SPO_SU_JI = Set_tbl_KEN(i).SPO_SU_JI + 1
                                Set_tbl_KEN(UBound(Set_tbl_KEN)).SPO_SU_JI = Set_tbl_KEN(UBound(Set_tbl_KEN)).SPO_SU_JI + 1
                                GK_SPO_SU_JI = GK_SPO_SU_JI + 1
                            End If
                        Case CYU_KBN_HJU            '補充
                            Set_tbl_KEN(i).HJU_SU_YO = Set_tbl_KEN(i).HJU_SU_YO + 1
                            Set_tbl_KEN(UBound(Set_tbl_KEN)).HJU_SU_YO = Set_tbl_KEN(UBound(Set_tbl_KEN)).HJU_SU_YO + 1
                            GK_HJU_SU_YO = GK_HJU_SU_YO + 1
                                                '検品済なら（？？？）
                            If Len(Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode))) <> 0 Then
                                Set_tbl_KEN(i).HJU_SU_JI = Set_tbl_KEN(i).HJU_SU_JI + 1
                                Set_tbl_KEN(UBound(Set_tbl_KEN)).HJU_SU_JI = Set_tbl_KEN(UBound(Set_tbl_KEN)).HJU_SU_JI + 1
                                GK_HJU_SU_JI = GK_HJU_SU_JI + 1
                            End If
                        Case CYU_KBN_BOU            '貿易
                            Set_tbl_KEN(i).BOU_SU_YO = Set_tbl_KEN(i).BOU_SU_YO + 1
                            Set_tbl_KEN(UBound(Set_tbl_KEN)).BOU_SU_YO = Set_tbl_KEN(UBound(Set_tbl_KEN)).BOU_SU_YO + 1
                            GK_BOU_SU_YO = GK_BOU_SU_YO + 1
                                                '検品済なら（？？？）
                            If Len(Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode))) <> 0 Then
                                Set_tbl_KEN(i).BOU_SU_JI = Set_tbl_KEN(i).BOU_SU_JI + 1
                                Set_tbl_KEN(UBound(Set_tbl_KEN)).BOU_SU_JI = Set_tbl_KEN(UBound(Set_tbl_KEN)).BOU_SU_JI + 1
                                GK_BOU_SU_JI = GK_BOU_SU_JI + 1
                            End If
                    End Select
                End If
                com = BtOpGetGreater

            Loop
        Next i
    End If
    
    If Not BOU_NON Then
        For i = 0 To UBound(Set_tbl_BOU) - 1
                                       '貿易分集計処理開始
            Call UniCode_Conv(K3_Y_SYU.JGYOBU, Last_JGYOBU)
            Call UniCode_Conv(K3_Y_SYU.KEY_MUKE_CODE, Set_tbl_BOU(i).MUKE_CODE)
            Call UniCode_Conv(K3_Y_SYU.KEY_SS_CODE, Set_tbl_BOU(i).SS_CODE)
            Call UniCode_Conv(K3_Y_SYU.KEY_CYU_KBN, "")
            Call UniCode_Conv(K3_Y_SYU.NAIGAI, "")
            Call UniCode_Conv(K3_Y_SYU.KEY_HIN_NO, "")
            Call UniCode_Conv(K3_Y_SYU.KEY_ID_NO, "")
                                                            
            com = BtOpGetGreater
                                    
            Do
            
                DoEvents
        
                sts = BTRV(com, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K3_Y_SYU, Len(K3_Y_SYU), 3)
                Select Case sts
                    Case BtNoErr
                                            '事業部／向け先ブレーク
                    If StrConv(Y_SYUREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                        Exit Do
                    End If
            
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        Call File_Error(sts, com, "出荷予定データ")
                        List_Dsp_Proc = SYS_ERR
                        Exit Function
            
                End Select
                                        
                If Set_tbl_BOU(i).MUKE_CODE = StrConv(Y_SYUREC.MUKE_CODE, vbUnicode) Then
                                            
                    Set_tbl_BOU(i).ALL_SU_YO = Set_tbl_BOU(i).ALL_SU_YO + 1
                    Set_tbl_BOU(UBound(Set_tbl_BOU)).ALL_SU_YO = Set_tbl_BOU(UBound(Set_tbl_BOU)).ALL_SU_YO + 1
                    GK_ALL_SU_YO = GK_ALL_SU_YO + 1
                                                '検品済なら（？？？）
                    If Len(Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode))) <> 0 Then
                        Set_tbl_BOU(i).ALL_SU_JI = Set_tbl_BOU(i).ALL_SU_JI + 1
                        Set_tbl_BOU(UBound(Set_tbl_KEN)).ALL_SU_JI = Set_tbl_BOU(UBound(Set_tbl_BOU)).ALL_SU_JI + 1
                        GK_ALL_SU_JI = GK_ALL_SU_JI + 1
                    End If
                                            
                    Select Case StrConv(Y_SYUREC.CYU_KBN, vbUnicode)
                        Case CYU_KBN_TUK            '月切
                            Set_tbl_BOU(i).TUK_SU_YO = Set_tbl_BOU(i).TUK_SU_YO + 1
                            Set_tbl_BOU(UBound(Set_tbl_BOU)).TUK_SU_YO = Set_tbl_BOU(UBound(Set_tbl_BOU)).TUK_SU_YO + 1
                            GK_TUK_SU_YO = GK_TUK_SU_YO + 1
                                                '検品済なら（？？？）
                            If Len(Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode))) <> 0 Then
                                Set_tbl_BOU(i).TUK_SU_JI = Set_tbl_BOU(i).TUK_SU_JI + 1
                                Set_tbl_BOU(UBound(Set_tbl_BOU)).TUK_SU_JI = Set_tbl_BOU(UBound(Set_tbl_BOU)).TUK_SU_JI + 1
                                GK_TUK_SU_JI = GK_TUK_SU_JI + 1
                            End If
                    
                        Case CYU_KBN_SPO            'スポット
                            Set_tbl_BOU(i).SPO_SU_YO = Set_tbl_BOU(i).SPO_SU_YO + 1
                            Set_tbl_BOU(UBound(Set_tbl_BOU)).SPO_SU_YO = Set_tbl_BOU(UBound(Set_tbl_BOU)).SPO_SU_YO + 1
                            GK_SPO_SU_YO = GK_SPO_SU_YO + 1
                                                '検品済なら（？？？）
                            If Len(Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode))) <> 0 Then
                                Set_tbl_BOU(i).SPO_SU_JI = Set_tbl_BOU(i).SPO_SU_JI + 1
                                Set_tbl_BOU(UBound(Set_tbl_BOU)).SPO_SU_JI = Set_tbl_BOU(UBound(Set_tbl_BOU)).SPO_SU_JI + 1
                                GK_SPO_SU_JI = GK_SPO_SU_JI + 1
                            End If
                        Case CYU_KBN_HJU            '補充
                            Set_tbl_BOU(i).HJU_SU_YO = Set_tbl_BOU(i).HJU_SU_YO + 1
                            Set_tbl_BOU(UBound(Set_tbl_BOU)).HJU_SU_YO = Set_tbl_BOU(UBound(Set_tbl_BOU)).HJU_SU_YO + 1
                            GK_HJU_SU_YO = GK_HJU_SU_YO + 1
                                                '検品済なら（？？？）
                            If Len(Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode))) <> 0 Then
                                Set_tbl_BOU(i).HJU_SU_JI = Set_tbl_BOU(i).HJU_SU_JI + 1
                                Set_tbl_BOU(UBound(Set_tbl_BOU)).HJU_SU_JI = Set_tbl_BOU(UBound(Set_tbl_BOU)).HJU_SU_JI + 1
                                GK_HJU_SU_JI = GK_HJU_SU_JI + 1
                            End If
                        Case CYU_KBN_BOU            '貿易
                            Set_tbl_BOU(i).BOU_SU_YO = Set_tbl_BOU(i).BOU_SU_YO + 1
                            Set_tbl_BOU(UBound(Set_tbl_BOU)).BOU_SU_YO = Set_tbl_BOU(UBound(Set_tbl_BOU)).BOU_SU_YO + 1
                            GK_BOU_SU_YO = GK_BOU_SU_YO + 1
                                                '検品済なら（？？？）
                            If Len(Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode))) <> 0 Then
                                Set_tbl_BOU(i).BOU_SU_JI = Set_tbl_BOU(i).BOU_SU_JI + 1
                                Set_tbl_BOU(UBound(Set_tbl_KEN)).BOU_SU_JI = Set_tbl_BOU(UBound(Set_tbl_KEN)).BOU_SU_JI + 1
                                GK_BOU_SU_JI = GK_BOU_SU_JI + 1
                            End If
                    End Select
                End If
                com = BtOpGetNext
            
            Loop
        
        Next i
    End If
                                    
                                    
                                    'テーブルリセット
    Set Y_SYUKA = Nothing
    
    
   
    Row = 0
    
    If Not MTS_NON Then
        For i = 0 To UBound(Set_tbl_MTS)
            
            Row = Row + 1
            Y_SYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
        
'            Y_SYUKA(Row, ColMUKE_CODE) = Set_tbl_MTS(i).MUKE_CODE
            Y_SYUKA(Row, ColMUKE_NAME) = Set_tbl_MTS(i).MUKE_NAME
        
            
            If Set_tbl_MTS(i).ALL_SU_YO <> 0 Then
                Y_SYUKA(Row, ColALL_Su) = Format(Set_tbl_MTS(i).ALL_SU_JI, "#0") & "/" & Format(Set_tbl_MTS(i).ALL_SU_YO, "#0")
            End If
            
            If Set_tbl_MTS(i).TUK_SU_YO <> 0 Then
                Y_SYUKA(Row, ColTUK_Su) = Format(Set_tbl_MTS(i).TUK_SU_JI, "#0") & "/" & Format(Set_tbl_MTS(i).TUK_SU_YO, "#0")
            End If
            
            If Set_tbl_MTS(i).SPO_SU_YO <> 0 Then
                Y_SYUKA(Row, ColSPO_Su) = Format(Set_tbl_MTS(i).SPO_SU_JI, "#0") & "/" & Format(Set_tbl_MTS(i).SPO_SU_YO, "#0")
            End If
            
            If Set_tbl_MTS(i).HJU_SU_YO <> 0 Then
                Y_SYUKA(Row, ColHJU_Su) = Format(Set_tbl_MTS(i).HJU_SU_JI, "#0") & "/" & Format(Set_tbl_MTS(i).HJU_SU_YO, "#0")
            End If
            
            If Set_tbl_MTS(i).BOU_SU_YO <> 0 Then
                Y_SYUKA(Row, ColBOU_Su) = Format(Set_tbl_MTS(i).BOU_SU_JI, "#0") & "/" & Format(Set_tbl_MTS(i).BOU_SU_YO, "#0")
            End If
        Next i
    
        Row = Row + 1
        Y_SYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
        
'        Y_SYUKA(Row, ColMUKE_CODE) = "------"
        Y_SYUKA(Row, ColMUKE_NAME) = "---------------------"
        Y_SYUKA(Row, ColALL_Su) = "--------------"
        Y_SYUKA(Row, ColTUK_Su) = "--------------"
        Y_SYUKA(Row, ColSPO_Su) = "--------------"
        Y_SYUKA(Row, ColHJU_Su) = "--------------"
        Y_SYUKA(Row, ColBOU_Su) = "--------------"
    
    
    End If
   
   
    If Not KEN_NON Then
        For i = 0 To UBound(Set_tbl_KEN)
            Row = Row + 1
            Y_SYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
        
'            Y_SYUKA(Row, ColMUKE_CODE) = Set_tbl_KEN(i).MUKE_CODE
            Y_SYUKA(Row, ColMUKE_NAME) = Set_tbl_KEN(i).MUKE_NAME
        
            If Set_tbl_KEN(i).ALL_SU_YO <> 0 Then
                Y_SYUKA(Row, ColALL_Su) = Format(Set_tbl_KEN(i).ALL_SU_JI, "#0") & "/" & Format(Set_tbl_KEN(i).ALL_SU_YO, "#0")
            End If
            If Set_tbl_KEN(i).TUK_SU_YO <> 0 Then
                Y_SYUKA(Row, ColTUK_Su) = Format(Set_tbl_KEN(i).TUK_SU_JI, "#0") & "/" & Format(Set_tbl_KEN(i).TUK_SU_YO, "#0")
            End If
            If Set_tbl_KEN(i).SPO_SU_YO <> 0 Then
                Y_SYUKA(Row, ColSPO_Su) = Format(Set_tbl_KEN(i).SPO_SU_JI, "#0") & "/" & Format(Set_tbl_KEN(i).SPO_SU_YO, "#0")
            End If
            If Set_tbl_KEN(i).HJU_SU_YO <> 0 Then
                Y_SYUKA(Row, ColHJU_Su) = Format(Set_tbl_KEN(i).HJU_SU_JI, "#0") & "/" & Format(Set_tbl_KEN(i).HJU_SU_YO, "#0")
            End If
            If Set_tbl_KEN(i).BOU_SU_YO <> 0 Then
                Y_SYUKA(Row, ColBOU_Su) = Format(Set_tbl_KEN(i).BOU_SU_JI, "#0") & "/" & Format(Set_tbl_KEN(i).BOU_SU_YO, "#0")
            End If
        Next i
    
        Row = Row + 1
        Y_SYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
        
'        Y_SYUKA(Row, ColMUKE_CODE) = "------"
        Y_SYUKA(Row, ColMUKE_NAME) = "---------------------"
        Y_SYUKA(Row, ColALL_Su) = "--------------"
        Y_SYUKA(Row, ColTUK_Su) = "--------------"
        Y_SYUKA(Row, ColSPO_Su) = "--------------"
        Y_SYUKA(Row, ColHJU_Su) = "--------------"
        Y_SYUKA(Row, ColBOU_Su) = "--------------"
    End If
   
    If Not BOU_NON Then
        For i = 0 To UBound(Set_tbl_BOU)
            Row = Row + 1
            
            Y_SYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
        
'            Y_SYUKA(Row, ColMUKE_CODE) = Set_tbl_BOU(i).MUKE_CODE
            Y_SYUKA(Row, ColMUKE_NAME) = Set_tbl_BOU(i).MUKE_NAME
        
            If Set_tbl_BOU(i).ALL_SU_YO <> 0 Then
                Y_SYUKA(Row, ColALL_Su) = Format(Set_tbl_BOU(i).ALL_SU_JI, "#0") & "/" & Format(Set_tbl_BOU(i).ALL_SU_YO, "#0")
            End If
            If Set_tbl_BOU(i).TUK_SU_YO <> 0 Then
                Y_SYUKA(Row, ColTUK_Su) = Format(Set_tbl_BOU(i).TUK_SU_JI, "#0") & "/" & Format(Set_tbl_BOU(i).TUK_SU_YO, "#0")
            End If
            If Set_tbl_BOU(i).SPO_SU_YO <> 0 Then
                Y_SYUKA(Row, ColSPO_Su) = Format(Set_tbl_BOU(i).SPO_SU_JI, "#0") & "/" & Format(Set_tbl_BOU(i).SPO_SU_YO, "#0")
            End If
            If Set_tbl_BOU(i).HJU_SU_YO <> 0 Then
                Y_SYUKA(Row, ColHJU_Su) = Format(Set_tbl_BOU(i).HJU_SU_JI, "#0") & "/" & Format(Set_tbl_BOU(i).HJU_SU_YO, "#0")
            End If
            If Set_tbl_BOU(i).BOU_SU_YO <> 0 Then
                Y_SYUKA(Row, ColBOU_Su) = Format(Set_tbl_BOU(i).BOU_SU_JI, "#0") & "/" & Format(Set_tbl_BOU(i).BOU_SU_YO, "#0")
            End If
        Next i
    
        Row = Row + 1
        Y_SYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
        
'        Y_SYUKA(Row, ColMUKE_CODE) = "------"
        Y_SYUKA(Row, ColMUKE_NAME) = "---------------------"
        Y_SYUKA(Row, ColALL_Su) = "--------------"
        Y_SYUKA(Row, ColTUK_Su) = "--------------"
        Y_SYUKA(Row, ColSPO_Su) = "--------------"
        Y_SYUKA(Row, ColHJU_Su) = "--------------"
        Y_SYUKA(Row, ColBOU_Su) = "--------------"
    End If
                        '合計
    Row = Row + 1
            
    Y_SYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
        
'    Y_SYUKA(Row, ColMUKE_CODE) = ""
    Y_SYUKA(Row, ColMUKE_NAME) = "総合計"
        
    If GK_ALL_SU_YO <> 0 Then
        Y_SYUKA(Row, ColALL_Su) = Format(GK_ALL_SU_JI, "#0") & "/" & Format(GK_ALL_SU_YO, "#0")
    End If
    If GK_TUK_SU_YO <> 0 Then
        Y_SYUKA(Row, ColTUK_Su) = Format(GK_TUK_SU_JI, "#0") & "/" & Format(GK_TUK_SU_YO, "#0")
    End If
    If GK_SPO_SU_YO <> 0 Then
        Y_SYUKA(Row, ColSPO_Su) = Format(GK_SPO_SU_JI, "#0") & "/" & Format(GK_SPO_SU_YO, "#0")
    End If
    If GK_HJU_SU_YO <> 0 Then
        Y_SYUKA(Row, ColHJU_Su) = Format(GK_HJU_SU_JI, "#0") & "/" & Format(GK_HJU_SU_YO, "#0")
    End If
    If GK_BOU_SU_YO <> 0 Then
        Y_SYUKA(Row, ColBOU_Su) = Format(GK_BOU_SU_JI, "#0") & "/" & Format(GK_BOU_SU_YO, "#0")
    End If
    
    
    Text(ptxDATE_YY).Text = Left(Format(Now, "yyyymmdd"), 4)
    Text(ptxDATE_MM).Text = Mid(Format(Now, "yyyymmdd"), 5, 2)
    Text(ptxDATE_DD).Text = Right(Format(Now, "yyyymmdd"), 2)
    Text(ptxTIME_HH).Text = Left(Format(Now, "HHmmss"), 2)
    Text(ptxTIME_MM).Text = Mid(Format(Now, "HHmmss"), 3, 2)
        
                                    'DBテーブルリンク
'    Y_SYUKA.QuickSort Min_Row, (Y_SYUKA.UpperBound(1)), 1, XORDER_ASCEND, XTYPE_STRING
    
    Set TDBGrid1.Array = Y_SYUKA
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    
        
    Call Input_UnLock
    
    List_Dsp_Proc = False
    
End Function

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1060601.MousePointer = vbHourglass

    Call Ctrl_Lock(F1060601)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1060601)


    F1060601.MousePointer = vbDefault

End Sub
Private Sub Command_Click(Index As Integer)

Dim sts As Integer
    
    Select Case Index
        Case 7                              '最新表示
            If List_Dsp_Proc Then           '集計＆表示
                Unload Me
            End If
            Command(7).SetFocus
        
        Case 11                             '終了
            Unload Me
    End Select
    
End Sub


Private Sub Form_Activate()
                                '集計＆表示
    If List_Dsp_Proc Then
        Unload Me
    End If
            
    Command(7).SetFocus

End Sub

Private Sub Form_DblClick()
    PrintForm
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   Ｋｅｙ Ｄｏｗｎ 前処理
'----------------------------------------------------------------------------
    Select Case KeyCode
        Case vbKeyF1 To vbKeyF12
            Command(KeyCode - vbKeyF1).Value = True
    End Select


End Sub

Private Sub Form_Load()
Dim i As Integer
Dim c As String * 128
Dim sts As Integer
    
    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If
    
    Show
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).Code = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).Code = Last_JGYOBU Then
            F1060601.Caption = "作業監視モニター（" + RTrim(JGYOBU_T(i).NAME) + ")"
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)
                                
                                
                                '向け先マスタOPEN
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '出荷予定データOPEN
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '対象向け先テーブル設定
    If MTS_SET_Proc() Then
        Unload Me
    End If
    
    End Sub



Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer
                                            '向け先管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "向け先管理マスタ")
        End If
    End If
                                            '出荷予定データＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定")
        End If
    End If
    
    sts = BTRV(BtOpReset, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1060601 = Nothing

    End
End Sub



Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    'メニューより終了要求
'    If JGYOBU_T(Index).CODE = " " Then
'        Unload Me
'    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).Code = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '事業部切り替え
    F1060601.Caption = "作業監視モニター（" + RTrim(JGYOBU_T(Index).NAME) + ")"
    Last_JGYOBU = JGYOBU_T(Index).Code
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)

End Sub


Private Function MTS_SET_Proc() As Integer

Dim sts     As Integer
Dim c       As String * 128
Dim i       As Integer
                                
    MTS_SET_Proc = True
                                
'----------------------------    対象ＭＴＳ取込み
    i = -1
    MTS_NON = False
    Do
        If GetIni(Format(App.EXEName), "MTS" & Format(i + 2, "00"), "SYS", c) Then
            Beep
            MsgBox "対象向け先コードの獲得に失敗しました。処理を中止して下さい。"
            Exit Function
        End If
    
    
        If Trim(c) = "END" Then
            
            If i = (-1) Then            '対象向け先なし
                MTS_NON = True
                Exit Do
            End If
            
            i = i + 1
            ReDim Preserve Set_tbl_MTS(i)
            Set_tbl_MTS(i).MUKE_CODE = ""
            Set_tbl_MTS(i).MUKE_NAME = "ＭＴＳ合計"
            Exit Do
        Else
            i = i + 1
            ReDim Preserve Set_tbl_MTS(i)
            
            If Len(Trim(c)) > 8 Then
                Set_tbl_MTS(i).MUKE_CODE = Left(Trim(c), 8)
                Set_tbl_MTS(i).SS_CODE = Mid(Trim(c), 8, 8 - Len(Trim(c)))
            Else
                Set_tbl_MTS(i).MUKE_CODE = Trim(c)
                Set_tbl_MTS(i).SS_CODE = ""
            End If
                                                    
                                        '向け先名称獲得
            Call UniCode_Conv(K0_MTS.MUKE_CODE, Set_tbl_MTS(i).MUKE_CODE)
            Call UniCode_Conv(K0_MTS.SS_CODE, Set_tbl_MTS(i).SS_CODE)
            sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
            Select Case sts
                Case BtNoErr
                    Set_tbl_MTS(i).MUKE_NAME = StrConv(MTSREC.MUKE_NAME, vbUnicode)
                Case BtErrKeyNotFound
                    Set_tbl_MTS(i).MUKE_NAME = ""
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ")
                    Exit Function
            End Select

        End If
    
    Loop
'----------------------------    対象検査工場取込み
    i = -1
    KEN_NON = False
    Do
        If GetIni(Format(App.EXEName), "KEN" & Format(i + 2, "00"), "SYS", c) Then
            Beep
            MsgBox "対象向け先コードの獲得に失敗しました。処理を中止して下さい。"
            Exit Function
        End If
    
    
        If Trim(c) = "END" Then
            
            If i = (-1) Then            '対象向け先なし
                KEN_NON = True
                Exit Do
            End If
            
            i = i + 1
            ReDim Preserve Set_tbl_KEN(i)
            Set_tbl_KEN(i).MUKE_CODE = ""
            Set_tbl_KEN(i).MUKE_NAME = "検査工場合計"
            Exit Do
        Else
            i = i + 1
            ReDim Preserve Set_tbl_KEN(i)
            
            Set_tbl_KEN(i).MUKE_CODE = Trim(c)
                                        '向け先名称獲得
            Call UniCode_Conv(K0_MTS.MUKE_CODE, Set_tbl_KEN(i).MUKE_CODE)
            Call UniCode_Conv(K0_MTS.SS_CODE, Set_tbl_KEN(i).SS_CODE)
            sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
            Select Case sts
                Case BtNoErr
                    Set_tbl_KEN(i).MUKE_NAME = StrConv(MTSREC.MUKE_NAME, vbUnicode)
                Case BtErrKeyNotFound
                    Set_tbl_KEN(i).MUKE_NAME = ""
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ")
                    Exit Function
            End Select

        End If
    
    Loop
'----------------------------    対象貿易分取込み
    i = -1
    BOU_NON = False
    Do
        If GetIni(Format(App.EXEName), "BOU" & Format(i + 2, "00"), "SYS", c) Then
            Beep
            MsgBox "対象向け先コードの獲得に失敗しました。処理を中止して下さい。"
            Exit Function
        End If
    
    
        If Trim(c) = "END" Then
            
            If i = (-1) Then            '対象向け先なし
                BOU_NON = True
                Exit Do
            End If
            
            i = i + 1
            ReDim Preserve Set_tbl_BOU(i)
            Set_tbl_BOU(i).MUKE_CODE = ""
            Set_tbl_BOU(i).MUKE_NAME = "貿易合計"
            Exit Do
        Else
            i = i + 1
            ReDim Preserve Set_tbl_BOU(i)
            
            Set_tbl_BOU(i).MUKE_CODE = Trim(c)
                                        '向け先名称獲得
            Call UniCode_Conv(K0_MTS.MUKE_CODE, Set_tbl_BOU(i).MUKE_CODE)
            sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
            Select Case sts
                Case BtNoErr
                    Set_tbl_BOU(i).MUKE_NAME = StrConv(MTSREC.MUKE_NAME, vbUnicode)
                Case BtErrKeyNotFound
                    Set_tbl_BOU(i).MUKE_NAME = ""
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ")
                    Exit Function
            End Select

        End If
    
    Loop

    MTS_SET_Proc = False

End Function

