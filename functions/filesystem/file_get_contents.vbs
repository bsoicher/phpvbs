'=======================================================================
'�t�@�C���̓��e��S�ĕ�����ɓǂݍ���
'=======================================================================
'�y�����z
'  filename  = string �f�[�^��ǂݍ��݂����t�@�C���̖��O�B
'�y�߂�l�z
'  �ǂݍ��񂾃f�[�^��Ԃ��܂��B���s�����ꍇ�� FALSE ��Ԃ��܂��B
'�y�����z
'  �t�@�C���̓��e�𕶎���ɓǂݍ���
' File_System class(https://github.com/masayukiando/phpvbs/blob/master/functions/filesystem/FileSystem.class.vbs)���ł̏���
'=======================================================================
Public function file_get_contents(filename)

    Dim ts
    Dim contents

    if left(filename,7) <> "http://" and file_exists( filename ) then
        Set TS = fso.OpenTextFile( fileMapPath(filename),1)

        '��̃t�@�C���̏ꍇ�A�G���[�ɂȂ��Ă��܂�
        If TS.AtEndOfStream <> True Then
           contents = TS.ReadAll
        End If

        file_get_contents = contents
        Exit Function
    end if

    if left(filename,7) <> "http://" then
        file_get_contents = false
        Exit Function
    end if

    Dim objWinHttp
    'Set objWinHttp = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
    'Set objWinHttp = Server.CreateObject("MSXML2.XMLHTTP")
    Set objWinHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
'        on error resume next

    objWinHttp.Open "GET", filename, false
    objWinHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    objWinHttp.Send
    'Response.Write objWinHttp.Status & " " & objWinHttp.StatusText

    file_get_contents = objWinHttp.ResponseText

    Set objWinHttp = Nothing

    '�L���ɂ����G���[�����𖳌��ɂ���
'        on error goto 0

end function