Attribute VB_Name = "wP_COMPO"
Option Explicit
'********************************************************************
'*                                                                  *
'*              構成マスタ  ファイル定義                            *
'*                                                                  *
'*          CREATE 2005.11.11                                       *
'********************************************************************
'ファイルＩＤ
Public Const wP_COMPO_ID$ = "P_COMPO"

'ページサイズ
Private Const wP_COMPO_PG_SIZ% = 1024

'ポジション・ブロック
Public wP_COMPO_POS      As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
'データ・バッファ
Public wP_COMPO_O_REC        As P_COMPO_O_REC_Tag


'データ・バッファ
Public wP_COMPO_K_REC        As P_COMPOREC_K_Tag

'キー定義

    
    
    
'キー・データ
Public wK0_P_COMPO           As KEY0_P_COMPO

Public wK1_P_COMPO           As KEY1_P_COMPO             '2014.06.23

Public wK2_P_COMPO           As KEY2_P_COMPO             '2018.0.220



Type P_COMPO_FSpeck
    fs                      As BtFileSpeck  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks3                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks4                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks5                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体

    ks6                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体    '2014.06.23

    ks7                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体    '2018.02.20
    ks8                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体    '2018.02.20
    ks9                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体    '2018.02.20
    ks10                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体    '2018.02.20
    ks11                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体    '2018.02.20
    ks12                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体    '2018.02.20

End Type

Private wP_COMPO_Speck       As P_COMPO_FSpeck

Public Function wP_COMPO_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              構成マスタ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    wP_COMPO_Open = True
                                            '構成マスタフルパス取込み
    sts = GetIni("FILE", wP_COMPO_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_COMPO]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, wP_COMPO_POS, wP_COMPO_O_REC, Len(wP_COMPO_O_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                Call File_Error(sts, BtOpOpen, "構成マスタ")
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "構成マスタ")
                Exit Function
        End Select
    Loop
    
    wP_COMPO_Open = False

End Function
