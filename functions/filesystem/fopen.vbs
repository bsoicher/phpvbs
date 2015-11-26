'=======================================================================
' �t�@�C���܂��� URL ���I�[�v������
'=======================================================================
'�y�����z
'  filename  =  string �f�[�^��ǂݍ��݂����t�@�C���̖��O�B
'  mode      =  string �X�g���[���ɗv����A�N�Z�X�`�����w�肵�܂�
'�y�߂�l�z
'  ���������ꍇ�Ƀt�@�C���|�C���^���\�[�X��Ԃ��܂��B
'�y�����z
'  fopen() �́Afilename  �Ŏw�肳�ꂽ���\�[�X���X�g���[���Ɍ��ѕt���܂��B
' File_System class(https://github.com/masayukiando/phpvbs/blob/master/functions/filesystem/FileSystem.class.vbs)���ł̏���
'=======================================================================
Public function fopen(filename, mode)

    Dim filePath
    filePath = fileMapPath(filename)

    If left(filePath,len("http://")) = "http://" Then
        fopen = file_get_contents(filePath)
        Exit Function
    End If

    Select Case mode
    Case "r"
        '�ǂݍ��݂݂̂ŃI�[�v�����܂��B
        Set ts = fso.OpenTextFile(filePath,1,false)
    Case "w"
        '�������݂ŃI�[�v�����܂��B
        Set ts = fso.OpenTextFile(filePath,2,true)
    Case "a"
        '�ǋL�ŃI�[�v�����܂��B
        Set ts = fso.OpenTextFile(filePath,8,true)
    Case "x"
        '�������݂ŃI�[�v�����܂��B�t�@�C�������݂����ꍇ��false��Ԃ��܂��B
        If is_file(filePath) Then
            fopen = false
        Else
            Set ts = fso.OpenTextFile(filePath,2,true)
        End If
    Case Else
        'empty
        ts = false
    End Select

end function