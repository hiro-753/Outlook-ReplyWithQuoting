Public Sub ReplyWithQuoting()
    '
    ' �S���ɕԐM�i�܂�Ԃ������j
    '
    ' ���[���ꗗ����ԐM���������[���s��I�����Ď��s����ƁA
    ' �e�L�X�g�`���̏ꍇ���p���t���܂�Ԃ������ŕԐM���郁�[����
    ' �쐬���ăG�f�B�^���J���܂��B
    '
    ' �E����(Signature)�́A�ォ�珐���{�^���Ő؂�ւ����܂��B
    ' �E���p���̑O�ɏ���������܂��B
    '
    '
    ' 2019.08.30 �V�K�쐬
    ' 2019.09.02 �I���t�H�[�J�X������h�~�����ǉ�
    '
    '
    ' Copy Right (C) Hiroyasu Watanabe 201909.02
    '
    Dim objWord As Variant
    Dim objSignature As Variant
    Dim strBody As String
    
    Dim s As Integer '���p���t�����p���̐擪�̈��p���̒���
    Dim e As Integer '���[���w�b�_�̍Ō�
    Dim length As Integer
    
    Dim strMode As String '���p�����[�h

    Dim msgCopy As MailItem '�������s���ꂢ�Ȃ����̃��[��
    Dim docBody As Object '�������s����Ă��Ȃ����̃��[����
    Dim strSel As String '�쐬���������p���t�����p��
        
    Dim msgReply As MailItem '�ԐM���[��
        
    ' �S���ɕԐM�{�^�����������s
    If TypeName(Application.ActiveWindow) = "Inspector" Then
        ActiveInspector.CommandBars.ExecuteMso "ReplyAll"
    Else
        ActiveExplorer.CommandBars.ExecuteMso "ReplyAll"
    End If
    
 
    '���̈��p���b�Z�[�W�������ւ��邽�߂�
    '���̈��p���b�Z�[�W�͈̔͂�ۑ�
    Set objWord = ActiveInspector.WordEditor
    If objWord.bookmarks.Exists("_MailOriginal") Then

       Set objBkm = objWord.bookmarks("_MailOriginal")
       Set objSel = objWord.Application.Selection
       
       '���L�͑I���t�H�[�J�X���\������Ă��܂��̂ŁA���O�ŃZ�b�g
       'objSel.Start = objBkm.Start
       'objSel.End = objBkm.End
    
    End If
       
       
    '�ԐM���[����MailItem�I�u�W�F�N�g���擾
    Set msgReply = ActiveInspector.CurrentItem

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
       '���p�����g�p����Ă��邩�`�F�b�N(���p����strPrefix�ɕۑ�)
       If Mid(strBody, s - 1, 1) <> vbLf Then
          strPrefix = GetPrefixText(strMode)
       End If


       '�ȉ��A�������s���Ȃ����p�����쐬���鏈��
       Set msgCopy = ActiveExplorer.Selection(1)
       '�\�����Ă��郁�[���� WordEditor ���擾
       Set docBody = msgCopy.GetInspector.WordEditor

       '�S�Ă�I��
       docBody.Range(0, 0).Select
       docBody.Application.Selection.WholeStory

       '�ʏ�̕ԐM�{�^���Ő����������p������擾�ۑ��������p�����A
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


       '�ȉ��A���p���t���̃w�b�_���̂ݎ��o������
       '
       '�ʏ�̕ԐM�{�^���Ő��������������s���ꂽ���p������A
       '���p���݂̂̍ŏ��̋�s�������ă��[���w�b�_�̍Ō�Ƃ���
       e = InStr(s, strBody, strPrefix & vbCrLf)

       '�ʏ�̕ԐM�{�^���Ő��������������s���ꂽ���p�����폜���A
       '������̏����{���[���w�b�_�����o��
       strBody = Left(strBody, e)
       
       '������̏������폜���ă��[���w�b�_�̂ݎ��o��
       length = Len(strBody)
       e = length - s
       strBody = strPrefix & Right(strBody, e)

       '���[���w�b�_�ɍ쐬�����������s���Ȃ����p���t���̈��p����A��
       strBody = strBody & vbCrLf & strSel
       
       
       '���̈��p�u�b�N�}�[�N������Ƃ��̂ݍ����ւ�
       If objWord.bookmarks.Exists("_MailOriginal") Then
          
          '��ɕۑ������͈͂��Z�b�g
          objSel.Start = objBkm.Start
          objSel.End = objBkm.End

          '�ԐM��"_MailOriginal"���ɃZ�b�g("_MailAutoSig"�͂�����Ȃ�)
          objSel.Text = strBody
    
       End If
    
    End If
    
    '���̃��[���̃t�H�[�J�X���O��
    docBody.Range(0, 0).Select
    
    '�ԐM���[���̃t�H�[�J�X��擪��
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
