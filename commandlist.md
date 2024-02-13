'win32api関数の宣言
Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare PtrSafe Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

'リストボックスのウィンドウクラス名
Const LISTBOX_CLASS = "LISTBOX"

'リストボックスのウィンドウスタイル
Const LBS_STANDARD = &H1A0000

'リストボックスのメッセージ
Const CB_ADDSTRING = &H143
Const CB_GETCURSEL = &H147

'メッセージボックスのタイプ
Const MB_OK = &H0

'リストボックスのハンドルを格納する変数
Dim hwndListBox As Long

Sub リストボックス作成()
    'リストボックスのウィンドウを作成
    hwndListBox = CreateWindowEx(0, LISTBOX_CLASS, "", LBS_STANDARD, 100, 100, 200, 200, 0, 0, 0, 0)
    
    'リストボックスに項目を追加
    SendMessage hwndListBox, CB_ADDSTRING, 0, "りんご"
    SendMessage hwndListBox, CB_ADDSTRING, 0, "みかん"
    SendMessage hwndListBox, CB_ADDSTRING, 0, "ぶどう"
End Sub

Sub リストボックス選択()
    'リストボックスで選択された項目のインデックスを取得
    Dim index As Long
    index = SendMessage(hwndListBox, CB_GETCURSEL, 0, 0)
    
    '選択された項目の名前を取得
    Dim name As String
    Select Case index
        Case 0
            name = "りんご"
        Case 1
            name = "みかん"
        Case 2
            name = "ぶどう"
        Case Else
            name = "なし"
    End Select
    
    'メッセージボックスに選択された項目を表示
    MessageBox 0, name & "が選択されました。", "リストボックス選択", MB_OK
End Sub

Sub リストボックス破棄()
    'リストボックスのウィンドウを破棄
    DestroyWindow hwndListBox
End Sub
