Public Sub ReplyWithQuoting()
    '
    ' 全員に返信（折り返し無し）
    '
    ' メール一覧から返信したいメール行を選択して実行すると、
    ' テキスト形式の場合引用符付き折り返し無しで返信するメールを
    ' 作成してエディタが開きます。
    '
    ' ・署名(Signature)は、後から署名ボタンで切り替えられます。
    ' ・引用文の前に署名が入ります。
    '
    '
    ' 2019.08.30 新規作成
    ' 2019.09.02 選択フォーカスちらつき防止処理追加
    '
    '
    ' Copy Right (C) Hiroyasu Watanabe 201909.02
    '
    Dim objWord As Variant
    Dim objSignature As Variant
    Dim strBody As String
    
    Dim s As Integer '引用符付き引用文の先頭の引用符の直後
    Dim e As Integer 'メールヘッダの最後
    Dim length As Integer
    
    Dim strMode As String '引用符モード

    Dim msgCopy As MailItem '自動改行されいない元のメール
    Dim docBody As Object '自動改行されていない元のメール文
    Dim strSel As String '作成したい引用符付き引用文
        
    Dim msgReply As MailItem '返信メール
        
    ' 全員に返信ボタン押下を実行
    If TypeName(Application.ActiveWindow) = "Inspector" Then
        ActiveInspector.CommandBars.ExecuteMso "ReplyAll"
    Else
        ActiveExplorer.CommandBars.ExecuteMso "ReplyAll"
    End If
    
 
    '元の引用メッセージを差し替えるために
    '元の引用メッセージの範囲を保存
    Set objWord = ActiveInspector.WordEditor
    If objWord.bookmarks.Exists("_MailOriginal") Then

       Set objBkm = objWord.bookmarks("_MailOriginal")
       Set objSel = objWord.Application.Selection
       
       '下記は選択フォーカスが表示されてしまうので、直前でセット
       'objSel.Start = objBkm.Start
       'objSel.End = objBkm.End
    
    End If
       
       
    '返信メールのMailItemオブジェクトを取得
    Set msgReply = ActiveInspector.CurrentItem

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
       '引用符が使用されているかチェック(引用符をstrPrefixに保存)
       If Mid(strBody, s - 1, 1) <> vbLf Then
          strPrefix = GetPrefixText(strMode)
       End If


       '以下、自動改行がない引用文を作成する処理
       Set msgCopy = ActiveExplorer.Selection(1)
       '表示しているメールの WordEditor を取得
       Set docBody = msgCopy.GetInspector.WordEditor

       '全てを選択
       docBody.Range(0, 0).Select
       docBody.Application.Selection.WholeStory

       '通常の返信ボタンで生成した引用文から取得保存した引用符を、
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


       '以下、引用符付きのヘッダ情報のみ取り出す処理
       '
       '通常の返信ボタンで生成した自動改行された引用文から、
       '引用符のみの最初の空行を見つけてメールヘッダの最後とする
       e = InStr(s, strBody, strPrefix & vbCrLf)

       '通常の返信ボタンで生成した自動改行された引用文を削除し、
       '文字列の署名＋メールヘッダを取り出す
       strBody = Left(strBody, e)
       
       '文字列の署名を削除してメールヘッダのみ取り出す
       length = Len(strBody)
       e = length - s
       strBody = strPrefix & Right(strBody, e)

       'メールヘッダに作成した自動改行がない引用符付きの引用文を連結
       strBody = strBody & vbCrLf & strSel
       
       
       '元の引用ブックマークがあるときのみ差し替え
       If objWord.bookmarks.Exists("_MailOriginal") Then
          
          '先に保存した範囲をセット
          objSel.Start = objBkm.Start
          objSel.End = objBkm.End

          '返信の"_MailOriginal"内にセット("_MailAutoSig"はいじらない)
          objSel.Text = strBody
    
       End If
    
    End If
    
    '元のメールのフォーカスを外す
    docBody.Range(0, 0).Select
    
    '返信メールのフォーカスを先頭に
    objWord.Range(0, 0).Select

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
