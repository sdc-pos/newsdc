Attribute VB_Name = "GLB_Window"
Option Explicit

Public hStatusWnd As Long                       'ステータスバーウィンド言うのハンドル保存用
'ステータスバーコントロールを作成する
Public Declare Function CreateStatusWindow Lib "comctl32.dll" _
(ByVal style As Long, ByVal lpszText As String, _
    ByVal hwndParent As Long, ByVal wID As Long) As Long
'stateの定数
Public Const CCS_TOP = &H1                      'クライアント領域の上辺に配置
Public Const CCS_BOTTOM = &H3                   '同、底辺に配置
Public Const SBARS_SIZEGRIP = &H100             'サイズ変更グリップをつける
'ウィンドウスタイルの定数
Public Const WS_BORDER = &H800000               'フォームの枠線がある
Public Const WS_CAPTION = &HC00000              'WS_BORDER Or WS_DLGFRAME
Public Const WS_CHILD = &H40000000              '親ウインドウを持つｺﾝﾄﾛｰﾙ(子ウインドウ)を作成する
Public Const WS_CHILDWINDOW = (WS_CHILD)        '子ウインドウ
Public Const WS_CLIPCHILDREN = &H2000000        'フォームの更新時にコントロールの再描画を抑制する
Public Const WS_CLIPSIBLINGS = &H4000000        'コントロールの更新時に他のコントロールの再描画を抑制する
Public Const WS_DISABLED = &H8000000            'Enabled = False
Public Const WS_DLGFRAME = &H400000             'リサイズできない枠線を持つ
Public Const WS_GROUP = &H20000                 'コントロール グループの最初のコントロールである
Public Const WS_HSCROLL = &H100000              '水平スクロールバーがある
Public Const WS_MAXIMIZE = &H1000000            '初期状態で最大化する
Public Const WS_MAXIMIZEBOX = &H10000           '最大化ボタンを持つ
Public Const WS_MINIMIZE = &H20000000           '初期状態で最小化する
Public Const WS_MINIMIZEBOX = &H20000           '最小化ボタンを持つ
Public Const WS_ICONIC = WS_MINIMIZE            'WS_MINIMIZE と同じ
Public Const WS_POPUP = &H80000000              'ポップアップ型ウインドウ
Public Const WS_SYSMENU = &H80000               'システムメニューがある
Public Const WS_TABSTOP = &H10000               'タブストップ可能
Public Const WS_THICKFRAME = &H40000            'リサイズ可能な枠線を持つ
Public Const WS_VISIBLE = &H10000000            'Visibleである
Public Const WS_VSCROLL = &H200000              '垂直スクロールバーがある
Public Const WS_OVERLAPPED = &H0                'フォーム枠線とキャプションバーがある
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX) '標準的なスタイル
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU) 'ポップアップ型の標準的なスタイル
Public Const WS_SIZEBOX = WS_THICKFRAME         'WS_THICKFRAME と同じ
Public Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW '32ビット版ではこれを使うことが多い
Public Const WS_TILED = WS_OVERLAPPED           'WS_OVERLAPPED のと同じ
Public Const CW_USEDEFAULT = &H80000000         '表示位置をWindowsが決定する
Public Const WM_USER = &H400                    'ユーザーが定義できるメッセージの使用領域を表すだけでこれ自体に意味はない
'メッセージを送る
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageStr Lib "user32.dll" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageAny Lib "user32.dll" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, lParam As Any) As Long
'=================--ステータスバー関連===================-
Public Const STATUSCLASSNAME = "msctls_statusbar32"     'クラス名
'メッセージ
Public Const SB_SETTEXT = (WM_USER + 1)                 'テキストを設定する
Public Const SB_GETTEXT = (WM_USER + 2)                 'テキストを取得する
Public Const SB_GETTEXTLENGTH = (WM_USER + 3)           'テキストの長さを取得する
Public Const SB_SETPARTS = (WM_USER + 4)                'ペインを設定する
Public Const SB_GETPARTS = (WM_USER + 6)                'ペイン数を取得する
Public Const SB_GETBORDERS = (WM_USER + 7)              '境界線の幅を取得する
Public Const SB_SETMINHEIGHT = (WM_USER + 8)            'ウィンドウを最小化したときの
                                                        'ステータスウィンドウの最小の高さ
Public Const SB_SIMPLE = (WM_USER + 9)                  'シンプルなスタイルにする
Public Const SB_GETRECT = (WM_USER + 10)                '指定されたペインのサイズを取得する
'テキストの表示のスタイル
Public Const SBT_OWNERDRAW = &H1000                     'オーナー描画
Public Const SBT_NOBORDERS = &H100                      '境界線なし
Public Const SBT_POPOUT = &H200                         '凸形
Public Const SBT_RTLREADING = &H400                     '右から左へ(ヘブライ語・アラビア語のみ)
Public Const SBT_SUNKEN = &H0                           '凹形
'境界線のスタイル
Public Const SBB_HORIZONTAL = 0                         '水平境界線の幅
Public Const SBB_VERTICAL = 1                           '垂直境界線の幅
Public Const SBB_DIVIDER = 2                            'ペイン区切り線の幅
'COMCTL32.dllからコモン子トロールクラスを登録する
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Long
Type tagINITCOMMONCONTROLSEX
    dwSize As Long                              '構造体のバイト数
    dwICC As Long                               'ロードするクラスを指定する
End Type
Public Const ICC_ANIMATE_CLASS = &H80           'アニメーションコントロール
Public Const ICC_BAR_CLASSES = &H4              'ツールバー、ステータスバー、スライダーバー
Public Const ICC_COOL_CLASSES = &H400           'リバーコントロール
Public Const ICC_DATE_CLASSES = &H100           '日時ピックアップコントロール
Public Const ICC_HOTKEY_CLASS = &H40            'ホットキーコントロール
Public Const ICC_INTERNET_CLASSES = &H800       'IPアドレスクラス
Public Const ICC_LISTVIEW_CLASSES = &H1         'リストビュー、ヘッダーコントロール
Public Const ICC_PAGESCROLLER_CLASS = &H1000    'ページャコントロール
Public Const ICC_PROGRESS_CLASS = &H20          'プログレスバーコントロール
Public Const ICC_TABCLASSES = &H8               'タブコントロール
Public Const ICC_TREEVIEW_CLASSES = &H2         'ツリービューコントロール
Public Const ICC_UPDOWN_CLASS = &H10            'アップダウンコントロール
Public Const ICC_USEREX_CLASSES = &H200         '拡張コンボボックスクラス
Public Const ICC_WIN95_CLASSES = &HFF           'Windows95コモンコントロール
'指定の座標位置にあるウィンドウハンドルを取得する(スクリーン座標)
Public Declare Function WindowFromPoint Lib "user32.dll" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetCursorPos Lib "user32.dll" (lpPoint As PointApi) As Long
Type PointApi
    X As Long
    Y As Long
End Type
'ウィンドウテキストの取得
Public Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" _
    (ByVal hwnd As Long, ByVal lpString As String, ByVal nMaxCount As String) As Long
'ウィンドウテキストサイズ取得
Public Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" _
    (ByVal hwnd As Long) As Long




