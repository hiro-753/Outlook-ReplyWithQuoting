Public Sub ReplyWithQuoting()
    '
    ' �S���ɕԐM�i�܂�Ԃ������j
    '
    ' ���[���ꗗ����ԐM���������[���s��I�����Ď��s����ƁA
    ' �e�L�X�g�`���̏ꍇ���p���t���܂�Ԃ������ŕԐM���郁�[����
    ' �쐬���ăG�f�B�^���J���܂��B
    '
    Dim msgReply As MailItem '�ԐM���[��
    Dim strBody As String '�������s���ꂽ���p���t�����p��
    Dim s As Integer '���p���t�����p���̐擪�̈��p���̒���
    Dim e As Integer '���[���w�b�_�̍Ō�
    Dim strMode As String '���p�����[�h

    Dim msgCopy As MailItem '�������s���ꂢ�Ȃ����̃��[��
    Dim docBody As Object '�������s����Ă��Ȃ����̃��[����
    Dim strSel As String '�쐬���������p���t�����p��

    '�ԐM���[�����쐬
    Set msgReply = ActiveExplorer.Selection(1).ReplyAll

    'TEXT �`����ݒ�(�����I�Ƀe�L�X�g�ɂ���ꍇ)
    'msgReply.BodyFormat = olFormatPlain

    '�ԐM�p���p����ݒ�
    strMode = "Reply"

    If msgReply.BodyFormat = olFormatHTML Then
       'HTML�̏ꍇ�������Ȃ�
    Else
       Dim strPrefix As String
       strPrefix = vbCrLf
       strBody = msgReply.Body

       '���p���̐擪�̈��p���̒���̕����ʒu���擾
       s = InStr(strBody, "-----Original Message-----")

       '�ʏ�̕ԐM�{�^���Ő��������������s���ꂽ���p����
       '���p�����g�p����Ă��邩�`�F�b�N
       If Mid(strBody, s - 1, 1) <> vbLf Then
          strPrefix = GetPrefixText(strMode)
       End If

       '�������s���Ȃ����p�����쐬����
       Set msgCopy = ActiveExplorer.Selection(1)

       '�\�����Ă��郁�[���� WordEditor ���擾
       Set docBody = msgCopy.GetInspector.WordEditor

       '�S�Ă�I��
       docBody.Range(0, 0).Select
       docBody.Application.Selection.WholeStory

       '�ʏ�̕ԐM�{�^���Ő����������p��������p�����A
       '���̕��̐擪�ɕt����
       strSel = strPrefix & docBody.Application.Selection.Text

       If strPrefix <> vbCrLf Then
          ' �I��͈͂̍s�̓��Ɉ��p�L����ǉ�
          strSel = Replace(strSel, vbCr, vbCr & strPrefix)

          ' �I��͈͂̍Ōオ���s�̏ꍇ�͍Ō�̈��p�L�����폜
          If strSel Like "*" & strPrefix Then
             strSel = Left(strSel, Len(strSel) - Len(strPrefix))
          End If
       Else
          '���p������̏ꍇ�͉��s��擪�ɒǉ�
          strSel = strPrefix & strSel
       End If

       '�ʏ�̕ԐM�{�^���Ő��������������s���ꂽ���p������A
       '���p���݂̂̍ŏ��̋�s�������ă��[���w�b�_�̍Ō�Ƃ���
       e = InStr(s, strBody, strPrefix & vbCrLf)

       '�ʏ�̕ԐM�{�^���Ő��������������s���ꂽ���p�����폜���A
       '���[���w�b�_�̂ݎ��o��
       strBody = Left(strBody, e)

       '���[���w�b�_�ɍ쐬�����������s���Ȃ����p���t���̈��p����A��
       strBody = strBody & vbCrLf & strSel

       '�ԐM���[���ɖ߂�
       msgReply.Body = strBody
    End If
    '�ԐM���[����\��
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
