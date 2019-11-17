Public Sub ReplyWithQuoting()
    '
    ' 全員に返信（折り返し無し）
    '
    ' メール一覧から返信したいメール行を選択して実行すると、
    ' テキスト形式の場合引用符付き折り返し無しで返信するメールを
    ' 作成してエディタが開きます。
    '
    Dim msgReply As MailItem '返信メール
    Dim strBody As String '自動改行された引用符付き引用文
    Dim s As Integer '引用符付き引用文の先頭の引用符の直後
    Dim e As Integer 'メールヘッダの最後
    Dim strMode As String '引用符モード

    Dim msgCopy As MailItem '自動改行されいない元のメール
    Dim docBody As Object '自動改行されていない元のメール文
    Dim strSel As String '作成したい引用符付き引用文

    '返信メールを作成
    Set msgReply = ActiveExplorer.Selection(1).ReplyAll

    'TEXT 形式を設定(強制的にテキストにする場合)
    'msgReply.BodyFormat = olFormatPlain

    '返信用引用符を設定
    strMode = "Reply"

    If msgReply.BodyFormat = olFormatHTML Then
       'HTMLの場合何もしない
    Else
       Dim strPrefix As String
       strPrefix = vbCrLf
       strBody = msgReply.Body

       '引用文の先頭の引用符の直後の文字位置を取得
       s = InStr(strBody, "-----Original Message-----")

       '通常の返信ボタンで生成した自動改行された引用文に
       '引用符が使用されているかチェック
       If Mid(strBody, s - 1, 1) <> vbLf Then
          strPrefix = GetPrefixText(strMode)
       End If

       '自動改行がない引用文を作成する
       Set msgCopy = ActiveExplorer.Selection(1)

       '表示しているメールの WordEditor を取得
       Set docBody = msgCopy.GetInspector.WordEditor

       '全てを選択
       docBody.Range(0, 0).Select
       docBody.Application.Selection.WholeStory

       '通常の返信ボタンで生成した引用文から引用符を、
       '元の文の先頭に付ける
       strSel = strPrefix & docBody.Application.Selection.Text

       If strPrefix <> vbCrLf Then
          ' 選択範囲の行の頭に引用記号を追加
          strSel = Replace(strSel, vbCr, vbCr & strPrefix)

          ' 選択範囲の最後が改行の場合は最後の引用記号を削除
          If strSel Like "*" & strPrefix Then
             strSel = Left(strSel, Len(strSel) - Len(strPrefix))
          End If
       Else
          '引用符が空の場合は改行を先頭に追加
          strSel = strPrefix & strSel
       End If

       '通常の返信ボタンで生成した自動改行された引用文から、
       '引用符のみの最初の空行を見つけてメールヘッダの最後とする
       e = InStr(s, strBody, strPrefix & vbCrLf)

       '通常の返信ボタンで生成した自動改行された引用文を削除し、
       'メールヘッダのみ取り出す
       strBody = Left(strBody, e)

       'メールヘッダに作成した自動改行がない引用符付きの引用文を連結
       strBody = strBody & vbCrLf & strSel

       '返信メールに戻す
       msgReply.Body = strBody
    End If
    '返信メールを表示
    msgReply.Display
End Sub

Function GetPrefixText(strMode As String) As String
    On Error Resume Next
    Dim wshShell As Variant
    Dim iStyle As Integer
    Dim strPrefix As String
    strPrefix = ""
    Set wshShell = CreateObject("WScript.Shell")
    iStyle = wshShell.RegRead("HKCU\Software\Microsoft\Office\" & Left(Application.Version, 2) & _
       ".0\Outlook\Preferences\" & strMode & "Style")
    If iStyle = 1000 Then
       strPrefix = wshShell.RegRead("HKCU\Software\Microsoft\Office\" & Left(Application.Version, 2) & _
          ".0\Outlook\Preferences\PrefixText")
       If strPrefix = "" Then
          strPrefix = "> "
       End If
    End If
    GetPrefixText = strPrefix
End Function
