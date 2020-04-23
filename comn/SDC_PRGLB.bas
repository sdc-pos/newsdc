Attribute VB_Name = "SDC_PRGLB"
Option Explicit
'フォント名
Global Const SDC_PGL_FONT_GOTHIC = "ＭＳ ゴシック"
Global Const SDC_PGL_FONT_PGOTHIC = "ＭＳ Ｐゴシック"
Global Const SDC_PGL_FONT_AGOTHIC = "@ＭＳ ゴシック"
Global Const SDC_PGL_FONT_MINCYO = "ＭＳ 明朝"
Global Const SDC_PGL_FONT_PMINCYO = "ＭＳ Ｐ明朝"
Global Const SDC_PGL_FONT_AMINCYO = "@ＭＳ 明朝"
Global Const SDC_PGL_FONT_BCODE39 = "3 of 9 Barcode"

'印刷用紙形態
'vbPRPSA3               A3、297 x 420 mm
Global Const SDC_PGL_SHEET_A3V% = 0          '縦
Global Const SDC_PGL_SHEET_A3H% = 1          '横
'vbPRPSA4               A4、210 x 297 mm
Global Const SDC_PGL_SHEET_A4V% = 2          '縦
Global Const SDC_PGL_SHEET_A4H% = 3          '横
'vbPRPSA4Small          A4 Small、210 x 297 mm
Global Const SDC_PGL_SHEET_A4SV% = 4         '縦
Global Const SDC_PGL_SHEET_A4SH% = 5         '横
'vbPRPSA5               A5、148 x 210 mm
Global Const SDC_PGL_SHEET_A5V% = 6          '縦
Global Const SDC_PGL_SHEET_A5H% = 7          '横
'vbPRPSB4               B4、250 x 354 mm
Global Const SDC_PGL_SHEET_B4V% = 8          '縦
Global Const SDC_PGL_SHEET_B4H% = 9          '横
'vbPRPSB5               B5、182 x 257 mm
Global Const SDC_PGL_SHEET_B5V% = 10         '縦
Global Const SDC_PGL_SHEET_B5H% = 11         '横
'vbPRPSLetter           レター、8 1/2 x 11 インチ
Global Const SDC_PGL_SHEET_LETV% = 12        '縦
Global Const SDC_PGL_SHEET_LETH% = 13        '横
'vbPRPSLetterSmall      レター スモール、8 1/2 x 11 インチ
Global Const SDC_PGL_SHEET_LTSV% = 14        '縦
Global Const SDC_PGL_SHEET_LTSH% = 15        '横
'vbPRPSTabloid          タブロイド、11 x 17 インチ
Global Const SDC_PGL_SHEET_TABV% = 16        '縦
Global Const SDC_PGL_SHEET_TABH% = 17        '横
'vbPRPSLedger           レジャー、17 x 11 インチ
Global Const SDC_PGL_SHEET_LEDV% = 18        '縦
Global Const SDC_PGL_SHEET_LEDH% = 19        '横
'vbPRPSLegal            リーガル、8 1/2 x 14 インチ
Global Const SDC_PGL_SHEET_LEGV% = 20        '縦
Global Const SDC_PGL_SHEET_LEGH% = 21        '横
'vbPRPSStatement        ステートメント、5 1/2 x 8 1/2 インチ
Global Const SDC_PGL_SHEET_STMV% = 22        '縦
Global Const SDC_PGL_SHEET_STMH% = 23        '横
'vbPRPSExecutive        エグゼクティブ、7 1/2 x 10 1/2 インチ
Global Const SDC_PGL_SHEET_EXEV% = 24        '縦
Global Const SDC_PGL_SHEET_EXEH% = 25        '横
'vbPRPSFolio            フォリオ、8 1/2 x 13 インチ
Global Const SDC_PGL_SHEET_FOLV% = 26        '縦
Global Const SDC_PGL_SHEET_FOLH% = 27        '横
'vbPRPSQuarto           クォート、215 x 275 mm
Global Const SDC_PGL_SHEET_QUAV% = 28        '縦
Global Const SDC_PGL_SHEET_QUAH% = 29        '横
'vbPRPS10x14            10 x 14 インチ
Global Const SDC_PGL_SHEET_10x14V% = 30      '縦
Global Const SDC_PGL_SHEET_10x14H% = 31      '横
'vbPRPS11x17            11 x 17 インチ
Global Const SDC_PGL_SHEET_11x17V% = 32      '縦
Global Const SDC_PGL_SHEET_11x17H% = 33      '横
'vbPRPSNote             ノート、8 1/2 x 11 インチ
Global Const SDC_PGL_SHEET_NOTV% = 34        '縦
Global Const SDC_PGL_SHEET_NOTH% = 35        '横
'vbPRPSCSheet           C サイズ シート
Global Const SDC_PGL_SHEET_CV% = 36          '縦
Global Const SDC_PGL_SHEET_CH% = 37          '横
'vbPRPSDSheet           D サイズ シート
Global Const SDC_PGL_SHEET_DV% = 38          '縦
Global Const SDC_PGL_SHEET_DH% = 39          '横
'vbPRPSESheet           E サイズ シート
Global Const SDC_PGL_SHEET_EV% = 40          '縦
Global Const SDC_PGL_SHEET_EH% = 41          '横
'vbPRPSFanfoldUS        U.S. ｽﾀﾝﾀﾞｰﾄﾞ ﾌｧﾝﾌｫｰﾙﾄﾞ、14 7/8 x 11 ｲﾝﾁ
Global Const SDC_PGL_SHEET_USV% = 42         '縦
Global Const SDC_PGL_SHEET_USH% = 43         '横
'vbPRPSUser             ユーザー定義
Global Const SDC_PGL_SHEET_USRV% = 44        '縦
Global Const SDC_PGL_SHEET_USRH% = 45        '横


'印刷用ワーク
Global Const SDC_PGL_LINI% = 99              '行数カウンタ初期値
Global SDC_PGL_Lcnt As Integer               '行数カウンタ

Global SDC_PGL_Pdate As String               '印刷開始日付（ﾍｯﾀﾞｰ用）
Global SDC_PGL_Ptime As String               '印刷開始時刻（ﾍｯﾀﾞｰ用）

Global SDC_PGL_PRT_CAN As Boolean            '印刷ｷｬﾝｾﾙ ﾌﾗｸﾞ

Function SDC_PGL_Init(Printr_Nm As String, Font_Nm As String, Font_Siz As Integer, Sheet_Type As Integer) As Integer
'----------------------------------------------------------------------
'　　　プリンター　初期設定
'
'  Printr_Nm ：プリンタ名取得用キー文字列
'  Font_Nm   ：フォント名（Null値ならプリンタ名の設定のみ）
'  Font_Siz  ：フォントサイズ
'  Sheet_Type：印刷用紙形態（詳細は本ﾓｼﾞｭｰﾙのGlobal定義参照）
'
'　戻り値：なし
'          CREATE 1999.04.17  S.Shibano
'----------------------------------------------------------------------
Dim Wk_Printer As PRINTER
Dim sts As Integer
Dim c As String
Dim USE_PRINTER As String

    SDC_PGL_Init = True

'指定帳票用プリンタ名　取得
    If GetIni("PRINTER", "SYSTEM", "SYS", c) Then
        Beep
        MsgBox "システムプリンタが定義されていません。", vbCritical
        Exit Function
    End If
    USE_PRINTER = RTrim(c)      'ﾃﾞﾌｫﾙﾄｾｯﾄ

    If GetIni("PRINTER", Printr_Nm, "SYS", c) = False Then
        USE_PRINTER = RTrim(c)
    Else
        Beep
        MsgBox Printr_Nm & "用プリンタの設定値(SYS.INI)無し", vbExclamation
        Exit Function
    End If

'指定帳票用プリンタ情報取得
    For Each Wk_Printer In Printers
        c = RTrim(Wk_Printer.DeviceName)
        If c = USE_PRINTER Then
            Set PRINTER = Wk_Printer
            Exit For
        End If
    Next
    If Font_Nm = "" Then
        SDC_PGL_Init = False
        Exit Function
    End If
'印刷フォント設定
    Call SDC_PGL_Font(Font_Nm, Font_Siz)

'印刷用紙形態　設定
    Select Case Sheet_Type

        Case SDC_PGL_SHEET_A3V             'Ａ３縦
            PRINTER.PaperSize = vbPRPSA3
            PRINTER.Orientation = vbPRORPortrait   '用紙の短辺を上にして印刷

        Case SDC_PGL_SHEET_A3H             'Ａ３横
            PRINTER.PaperSize = vbPRPSA3
            PRINTER.Orientation = vbPRORLandscape  '用紙の長辺を上にして印刷

        Case SDC_PGL_SHEET_A4V             'Ａ４縦
            PRINTER.PaperSize = vbPRPSA4
            PRINTER.Orientation = vbPRORPortrait   '用紙の短辺を上にして印刷

        Case SDC_PGL_SHEET_A4H             'Ａ４横
            PRINTER.PaperSize = vbPRPSA4
            PRINTER.Orientation = vbPRORLandscape  '用紙の長辺を上にして印刷

        Case SDC_PGL_SHEET_A4SV            'A4 Small縦
            PRINTER.PaperSize = vbPRPSA4Small
            PRINTER.Orientation = vbPRORPortrait   '用紙の短辺を上にして印刷

        Case SDC_PGL_SHEET_A4SH            'A4 Small横
            PRINTER.PaperSize = vbPRPSA4Small
            PRINTER.Orientation = vbPRORLandscape  '用紙の長辺を上にして印刷

        Case SDC_PGL_SHEET_A5V             'Ａ５縦
            PRINTER.PaperSize = vbPRPSA5
            PRINTER.Orientation = vbPRORPortrait   '用紙の短辺を上にして印刷

        Case SDC_PGL_SHEET_A5H             'Ａ５横
            PRINTER.PaperSize = vbPRPSA5
            PRINTER.Orientation = vbPRORLandscape  '用紙の長辺を上にして印刷

        Case SDC_PGL_SHEET_B4V             'Ｂ４縦
            PRINTER.PaperSize = vbPRPSB4
            PRINTER.Orientation = vbPRORPortrait   '用紙の短辺を上にして印刷

        Case SDC_PGL_SHEET_B4H             'Ｂ４横
            PRINTER.PaperSize = vbPRPSB4
            PRINTER.Orientation = vbPRORLandscape  '用紙の長辺を上にして印刷

        Case SDC_PGL_SHEET_B5V             'Ｂ５縦
            PRINTER.PaperSize = vbPRPSB5
            PRINTER.Orientation = vbPRORPortrait   '用紙の短辺を上にして印刷

        Case SDC_PGL_SHEET_B5H             'Ｂ５横
            PRINTER.PaperSize = vbPRPSB5
            PRINTER.Orientation = vbPRORLandscape  '用紙の長辺を上にして印刷

        Case SDC_PGL_SHEET_LETV            'レター縦
            PRINTER.PaperSize = vbPRPSLetter
            PRINTER.Orientation = vbPRORPortrait   '用紙の短辺を上にして印刷

        Case SDC_PGL_SHEET_LETH            'レター横
            PRINTER.PaperSize = vbPRPSLetter
            PRINTER.Orientation = vbPRORLandscape  '用紙の長辺を上にして印刷

        Case SDC_PGL_SHEET_LTSV            'レター スモール縦
            PRINTER.PaperSize = vbPRPSLetterSmall
            PRINTER.Orientation = vbPRORLandscape  '用紙の長辺を上にして印刷

        Case SDC_PGL_SHEET_LTSH            'レター スモール横
            PRINTER.PaperSize = vbPRPSLetterSmall
            PRINTER.Orientation = vbPRORLandscape  '用紙の長辺を上にして印刷

        Case SDC_PGL_SHEET_TABV             'タブロイド縦
            PRINTER.PaperSize = vbPRPSTabloid
            PRINTER.Orientation = vbPRORPortrait   '用紙の短辺を上にして印刷

        Case SDC_PGL_SHEET_TABH             'タブロイド横
            PRINTER.PaperSize = vbPRPSTabloid
            PRINTER.Orientation = vbPRORLandscape  '用紙の長辺を上にして印刷

        Case SDC_PGL_SHEET_LEDV             'レジャー縦
            PRINTER.PaperSize = vbPRPSLedger
            PRINTER.Orientation = vbPRORPortrait   '用紙の短辺を上にして印刷

        Case SDC_PGL_SHEET_LEDH             'レジャー横
            PRINTER.PaperSize = vbPRPSLedger
            PRINTER.Orientation = vbPRORLandscape  '用紙の長辺を上にして印刷

        Case SDC_PGL_SHEET_LEGV             'リーガル縦
            PRINTER.PaperSize = vbPRPSLegal
            PRINTER.Orientation = vbPRORPortrait   '用紙の短辺を上にして印刷

        Case SDC_PGL_SHEET_LEGH             'リーガル横
            PRINTER.PaperSize = vbPRPSLegal
            PRINTER.Orientation = vbPRORLandscape  '用紙の長辺を上にして印刷

        Case SDC_PGL_SHEET_STMV             'ステートメント縦
            PRINTER.PaperSize = vbPRPSStatement
            PRINTER.Orientation = vbPRORPortrait   '用紙の短辺を上にして印刷

        Case SDC_PGL_SHEET_STMH             'ステートメント横
            PRINTER.PaperSize = vbPRPSStatement
            PRINTER.Orientation = vbPRORLandscape  '用紙の長辺を上にして印刷

        Case SDC_PGL_SHEET_EXEV             'エグゼクティブ縦
            PRINTER.PaperSize = vbPRPSExecutive
            PRINTER.Orientation = vbPRORPortrait   '用紙の短辺を上にして印刷

        Case SDC_PGL_SHEET_EXEH             'エグゼクティブ横
            PRINTER.PaperSize = vbPRPSExecutive
            PRINTER.Orientation = vbPRORLandscape  '用紙の長辺を上にして印刷

        Case SDC_PGL_SHEET_FOLV             'フォリオ縦
            PRINTER.PaperSize = vbPRPSFolio
            PRINTER.Orientation = vbPRORPortrait   '用紙の短辺を上にして印刷

        Case SDC_PGL_SHEET_FOLH             'フォリオ縦
            PRINTER.PaperSize = vbPRPSFolio
            PRINTER.Orientation = vbPRORLandscape  '用紙の長辺を上にして印刷

        Case SDC_PGL_SHEET_QUAV             'クォート縦
            PRINTER.PaperSize = vbPRPSQuarto
            PRINTER.Orientation = vbPRORPortrait   '用紙の短辺を上にして印刷

        Case SDC_PGL_SHEET_QUAH             'クォート横
            PRINTER.PaperSize = vbPRPSQuarto
            PRINTER.Orientation = vbPRORLandscape  '用紙の長辺を上にして印刷

        Case SDC_PGL_SHEET_10x14V             '10 x 14縦
            PRINTER.PaperSize = vbPRPS10x14
            PRINTER.Orientation = vbPRORPortrait   '用紙の短辺を上にして印刷

        Case SDC_PGL_SHEET_10x14H             '10 x 14横
            PRINTER.PaperSize = vbPRPS10x14
            PRINTER.Orientation = vbPRORLandscape  '用紙の長辺を上にして印刷

        Case SDC_PGL_SHEET_11x17V             '11 x 17縦
            PRINTER.PaperSize = vbPRPS11x17
            PRINTER.Orientation = vbPRORPortrait   '用紙の短辺を上にして印刷

        Case SDC_PGL_SHEET_11x17H             '11 x 17横
            PRINTER.PaperSize = vbPRPS11x17
            PRINTER.Orientation = vbPRORLandscape  '用紙の長辺を上にして印刷

        Case SDC_PGL_SHEET_NOTV             'ノート縦
            PRINTER.PaperSize = vbPRPSNote
            PRINTER.Orientation = vbPRORPortrait   '用紙の短辺を上にして印刷

        Case SDC_PGL_SHEET_NOTH             'ノート横
            PRINTER.PaperSize = vbPRPSNote
            PRINTER.Orientation = vbPRORLandscape  '用紙の長辺を上にして印刷

        Case SDC_PGL_SHEET_CV             'C サイズ縦
            PRINTER.PaperSize = vbPRPSCSheet
            PRINTER.Orientation = vbPRORPortrait   '用紙の短辺を上にして印刷

        Case SDC_PGL_SHEET_CH             'C サイズ横
            PRINTER.PaperSize = vbPRPSCSheet
            PRINTER.Orientation = vbPRORLandscape  '用紙の長辺を上にして印刷

        Case SDC_PGL_SHEET_DV             'D サイズ縦
            PRINTER.PaperSize = vbPRPSDSheet
            PRINTER.Orientation = vbPRORPortrait   '用紙の短辺を上にして印刷

        Case SDC_PGL_SHEET_DH             'D サイズ横
            PRINTER.PaperSize = vbPRPSDSheet
            PRINTER.Orientation = vbPRORLandscape  '用紙の長辺を上にして印刷

        Case SDC_PGL_SHEET_EV             'E サイズ縦
            PRINTER.PaperSize = vbPRPSESheet
            PRINTER.Orientation = vbPRORPortrait   '用紙の短辺を上にして印刷

        Case SDC_PGL_SHEET_EH             'E サイズ横
            PRINTER.PaperSize = vbPRPSESheet
            PRINTER.Orientation = vbPRORLandscape  '用紙の長辺を上にして印刷

        Case SDC_PGL_SHEET_USV             'U.S. ｽﾀﾝﾀﾞｰﾄﾞ縦
            PRINTER.PaperSize = vbPRPSFanfoldUS
            PRINTER.Orientation = vbPRORPortrait   '用紙の短辺を上にして印刷

        Case SDC_PGL_SHEET_USH             'U.S. ｽﾀﾝﾀﾞｰﾄﾞ横
            PRINTER.PaperSize = vbPRPSFanfoldUS
            PRINTER.Orientation = vbPRORLandscape  '用紙の長辺を上にして印刷
            
        Case SDC_PGL_SHEET_USRV              'ユーザ定義　縦
            'Printer.PaperSize = vbPRPSUser
            PRINTER.Orientation = vbPRORPortrait   '用紙の短辺を上にして印刷
            
        Case SDC_PGL_SHEET_USRH              'ユーザ定義　横
            'Printer.PaperSize = vbPRPSUser
            PRINTER.Orientation = vbPRORLandscape  '用紙の長辺を上にして印刷
            
        Case Else
        
    End Select

'印刷用ワーク初期設定
    SDC_PGL_Lcnt = SDC_PGL_LINI         '行数カウンタ初期化
    SDC_PGL_Pdate = Date                '印刷開始　日付
    SDC_PGL_Ptime = Time                '　　　　　時刻

    SDC_PGL_Init = False

End Function

Sub SDC_PGL_Font(Font_Nm As String, Font_Siz As Integer)
'----------------------------------------------------------------------
'　　　             フォント設定
'
'  Font_Nm   ：フォント名
'  Font_Siz  ：フォントサイズ
'
'          CREATE 1999.05.28  S.Shibano
'----------------------------------------------------------------------
Dim W_Font As New StdFont

'印刷フォント設定
    With W_Font
        .NAME = Font_Nm
        .Size = Font_Siz
    End With
    Set PRINTER.Font = W_Font

End Sub
