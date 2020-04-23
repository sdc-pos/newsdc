Attribute VB_Name = "PI00020com"
Option Explicit

Public Taget_Key    As String * 8       '更新対象の指図票№

Public BUNNOU_CNT   As Integer          '分納回数

Public Doukon_Tbl_No(0 To 19)   As String * 1

Public POS_UMU          As Boolean

Public PRI_S_TANTO      As Boolean      '収支／担当者印刷 OFF:印刷なし ON:印刷あり
Public PRI_MAIN_BCR     As Boolean      'ﾒｲﾝﾊﾞｰｺｰﾄﾞ OFF:印刷なし ON:印刷あり
Public PRI_BIKOU_BCR    As Boolean      '備考欄　OFF：入力値　ON:出荷BCR
Public PRI_DOUKON       As Boolean      '商品化検査　同梱 OFF:印刷なし ON:印刷あり

Public PRI_NYUKO_IN     As Boolean      '入庫完了印　同梱 OFF:印刷なし ON:印刷あり
Public PRI_INPUT_IN     As Boolean      '入力完了印　同梱 OFF:印刷なし ON:印刷あり

Public PRI_SAGYO_DAY    As Boolean      '作業日／数量／担当 OFF:印刷なし ON:印刷あり 2007.05.22
Public PRI_HINBAN_BIKOU As Boolean      '下部　品番／№／数量 OFF:印刷なし ON:印刷あり 2007.05.22


Public JISEKI_TITLE     As Variant      '自責の名称タイトル
Public TASEKI_TITLE     As Variant      '他責の名称タイトル

Public JISSEKI_DSP      As String * 1   '2008.08.19

'---------------------------------------------- *商品化指図ﾃﾞｰﾀ（親）別ポインタ
'ポジショニング
Public wP_SSHIJI_O_POS  As POSBLK
'データ・バッファ
Public wP_SSHIJI_O_REC  As P_SSHIJI_O_REC_Tag
'キー・データ
Public K0_wP_SSHIJI_O   As KEY0_P_SSHIJI_O
Public K1_wP_SSHIJI_O   As KEY1_P_SSHIJI_O
'2016.01.06 ダミー
Private Const LAST_UPDATE_DAY$ = "([PI00020] 2016.01.06 15:30) "


' ------------------------------------------------------------------------
'       指定した精度の数値に四捨五入します。
'
' @Param    dValue      丸め対象の倍精度浮動小数点数。
' @Param    iDigits     戻り値の有効桁数の精度。
' @Return               iDigits に等しい精度の数値に四捨五入された数値。
' ------------------------------------------------------------------------
Public Function ToHalfAdjust(ByVal dValue As Currency, ByVal iDigits As Integer) As Currency
    Dim dCoef As Double

    dCoef = (10 ^ iDigits)

    If dValue > 0 Then
        ToHalfAdjust = Int(CDbl(dValue * dCoef + 0.5)) / dCoef
    Else
        ToHalfAdjust = Fix(CDbl(dValue * dCoef - 0.5)) / dCoef
    End If
End Function


Sub Main()
    
Dim lngReturnValue      As Long
Dim strMyTitle          As String
Dim lngPrevHwnd         As Long
Dim lngTopHwnd          As Long
Dim lngThreadID1        As Long
Dim lngThreadID2        As Long
    
    
    


    ' 2重起動の場合は、手前に持ってきて自分自身は終了する
    strMyTitle = App.Title
    App.Title = "$" & App.Title
    lngPrevHwnd = FindWindow("ThunderRT6Main", strMyTitle)
    If lngPrevHwnd <> 0 Then
    lngTopHwnd = GetLastActivePopup(lngPrevHwnd)
    If IsIconic(lngTopHwnd) = WIN32API_TRUE Then
    lngReturnValue = ShowWindow(lngTopHwnd, SW_NORMAL)
    End If
    lngThreadID1 = GetWindowThreadProcessId(GetForegroundWindow(), ByVal 0&)
    lngThreadID2 = GetCurrentThreadId()
    lngReturnValue = AttachThreadInput(lngThreadID2, lngThreadID1, 1)
    lngReturnValue = SetForegroundWindow(lngTopHwnd)
    lngReturnValue = AttachThreadInput(lngThreadID2, lngThreadID1, 0)
    Exit Sub
    End If
    App.Title = strMyTitle




    PI000201.Show
End Sub

