Class File_System

    Private fso
    Private ts

    '
    'Initialize Class
    ' 
    '@access private
    '
    Private Sub Class_Initialize()
        Set fso = Server.CreateObject("Scripting.FileSystemObject")
    End Sub

    '
    'Terminate Class
    ' 
    '@access private
    '
    Private Sub Class_Terminate()
        Set fso = Nothing
    End Sub

    '=======================================================================
    ' �p�X���̃t�@�C�����̕�����Ԃ�
    '=======================================================================
    '�y�����z
    '  path      = string   �p�X�B
    '  suffix    = string   �t�@�C�������A suffix  �ŏI������ꍇ�A ���̕������J�b�g����܂��B
    '�y�߂�l�z
    '  �w�肵�� path  �̃x�[�X����Ԃ��܂��B
    '�y�����z
    '  �E���̊֐��́A�t�@�C���ւ̃p�X��L���镶����������Ƃ��A �t�@�C���̃x�[�X����Ԃ��܂��B
    '=======================================================================
    Function basename(path, suffix)

        Dim b
        b = preg_replace("/^.*[\/\\]/g","",path,null,null)

        If len(suffix) > 0 Then
            If Right(b,len(suffix)) = suffix Then
                b = Left(b,len(b) - len(suffix))
            End If
        End If

        basename = b

    End Function

    '=======================================================================
    ' �t�@�C�����R�s�[����
    '=======================================================================
    '�y�����z
    '  source  = string   �R�s�[���t�@�C���ւ̃p�X�B
    '  dest    = string   �R�s�[��̃p�X�B
    '�y�߂�l�z
    '  ���������ꍇ�� TRUE ���A���s�����ꍇ�� FALSE ��Ԃ��܂��B
    '�y�����z
    '  �E �t�@�C�� source  �� dest  �ɃR�s�[���܂��B
    '  �E �t�@�C�����ړ��������Ȃ�́Arename() �֐����g�p���Ă��������B 
    '=======================================================================
    Public Function copy(source,dest)
        fso.CopyFile source,dest
    End Function

    '=======================================================================
    ' �I�[�v�����ꂽ�t�@�C���|�C���^���N���[�Y����
    '=======================================================================
    '�y�����z
    '  
    '�y�߂�l�z
    '  
    '�y�����z
    '  �t�@�C�����N���[�Y���܂��B
    '=======================================================================
    Public function fclose
        ts.close
        Set ts = Nothing
    end function

    '=======================================================================
    ' �t�@�C���|�C���^���t�@�C���I�[�ɒB���Ă��邩�ǂ������ׂ�
    '=======================================================================
    '�y�����z
    '  
    '�y�߂�l�z
    '  �t�@�C���|�C���^�� EOF �ɒB���Ă��邩�܂��̓G���[ (�\�P�b�g�^�C���A�E�g���܂݂܂�) �̏ꍇ�� TRUE �A ���̑��̏ꍇ�� FALSE ��Ԃ��܂��B
    '�y�����z
    '  �t�@�C���|�C���^���t�@�C���I�[�ɒB���Ă��邩�ǂ����𒲂ׂ܂��B
    '=======================================================================
    Public function feof
        feof = ts.AtEndofStream
    end function

    '=======================================================================
    ' �t�@�C���|�C���^����1�������o��
    '=======================================================================
    '�y�����z
    '  
    '�y�߂�l�z
    '  �t�@�C���|�C���^���� 1 �����ǂݏo���A ���̕�������Ȃ镶�����Ԃ��܂��BEOF �̏ꍇ�� FALSE ��Ԃ��܂��B
    '�y�����z
    '  �w�肵���t�@�C���|�C���^���� 1 �����ǂݏo���܂��B
    '=======================================================================
    Public function fgetc
        If ts.AtEndofStream Then
            fgetc = false
        Else
            fgetc = ts.Read(1)
        End If
    end function

    '=======================================================================
    ' �t�@�C���|�C���^����s���擾���ACSV�t�B�[���h����������
    '=======================================================================
    '�y�����z
    '  
    '�y�߂�l�z
    '  �ǂݍ��񂾃t�B�[���h�̓��e���܂ސ��l�Y���z���Ԃ��܂��B
    '�y�����z
    '  fgets() �ɓ���͎��Ă��܂����A fgetcsv() �͍s�� CSV  �t�H�[�}�b�g�̃t�B�[���h�Ƃ��ēǍ��ݏ������s���A �ǂݍ��񂾃t�B�[���h���܂ޔz���Ԃ��Ƃ����Ⴂ������܂��B
    '=======================================================================
    Public Function fgetcsv(delim)

        Dim tmp,d
        If len(delim) > 0 Then d = delim Else d = ","
        tmp = ts.ReadLine
        fgetcsv = fgetcsv_helper(tmp,d)

    End Function

    '************************************
    Public Function fgetcsv_helper(str,d)

        Dim matchAll,key
        Dim data,field,record : field = 0 : record = 0
        ReDim data(0)

        If preg_match_all(_
        "/" & d & "|" & vbCr & "?" & vbLf & "|[^" & d & """" & vbCrLf & "][^" & d & "" & vbCrLf & "]*|""(?:[^""]|"""")*""/",_
        str, matchAll,PREG_PATTERN_ORDER,"") Then
            For Each key In matchAll(0)
                Select Case key
                Case d
                    field = field + 1
                Case vbCrLf
                    [] data , ""
                    record = record +1
                Case Else
                    If left(key,1) = """" Then
                        key = Replace(substr(key,2,-1),"""""","""")
                    End if
                    [] data(record), key
                End Select
            Next
        End If

        fgetcsv_helper = data

    End Function

    '=======================================================================
    ' �t�@�C���|�C���^���� 1 �s�擾����
    '=======================================================================
    '�y�����z
    '  
    '�y�߂�l�z
    '  �t�@�C���|�C���^���� 1 �s�擾����
    '�y�����z
    '  �t�@�C���|�C���^���� 1 �s�擾���܂��B
    '=======================================================================
    Public function fgets
        fgets = ts.ReadLine
    end function

    '=======================================================================
    ' �t�@�C���|�C���^���� 1 �s���o���AHTML �^�O����菜��
    '=======================================================================
    '�y�����z
    '  
    '�y�߂�l�z
    '  HTML �� PHP �R�[�h����菜�����������Ԃ��܂��B
    '�y�����z
    '  fgets() �Ɠ����ł����A fgetss() �͓ǂݍ��񂾃e�L�X�g���� HTML ����� PHP �̃^�O����菜�����Ƃ��邱�Ƃ��قȂ�܂��B
    '=======================================================================
    Public function fgetss
        fgets = strip_tags(ts.ReadLine)
    end function

    '=======================================================================
    ' �t�@�C���܂��̓f�B���N�g�������݂��邩�ǂ������ׂ�
    '=======================================================================
    '�y�����z
    '  path      = string   �t�@�C�����邢�̓f�B���N�g���ւ̃p�X�B
    '�y�߂�l�z
    '  �t�@�C�����邢�̓f�B���N�g�������݂��邩�ǂ����𒲂ׂ܂��B
    '�y�����z
    '  �t�@�C�����邢�̓f�B���N�g�������݂��邩�ǂ����𒲂ׂ܂��B
    '=======================================================================
    Public Function file_exists(ByVal filename)

        file_exists = false
        filename = fileMapPath(filename)
        If fso.FileExists(filename) Then file_exists = true
        If fso.FolderExists(filename) Then file_exists = true

    End Function

    '=======================================================================
    '�t�@�C���̓��e��S�ĕ�����ɓǂݍ���
    '=======================================================================
    '�y�����z
    '  filename  = string �f�[�^��ǂݍ��݂����t�@�C���̖��O�B
    '�y�߂�l�z
    '  �ǂݍ��񂾃f�[�^��Ԃ��܂��B���s�����ꍇ�� FALSE ��Ԃ��܂��B
    '�y�����z
    '  �t�@�C���̓��e�𕶎���ɓǂݍ���
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

    '=======================================================================
    ' �t�@�C���̍ŏI�A�N�Z�X�������擾����
    '=======================================================================
    '�y�����z
    '  filename = string   �t�@�C���ւ̃p�X�B
    '�y�߂�l�z
    '  �t�@�C���̍ŏI�A�N�Z�X������Ԃ��A�G���[�̏ꍇ�� FALSE ��Ԃ��܂��B
    '�y�����z
    '  �w�肵���t�@�C���̍ŏI�A�N�Z�X�������擾���܂��B
    '=======================================================================
    Public Function fileatime(filename)

        Dim f
        filename = fileMapPath(filename)
        set f = fso.GetFile(filename)
        fileatime = f.DateLastAccessed

    End Function

    '=======================================================================
    ' �t�@�C���̍X�V�������擾����
    '=======================================================================
    '�y�����z
    '  filename = string   �t�@�C���ւ̃p�X�B
    '�y�߂�l�z
    '  �t�@�C���̍ŏI�X�V������Ԃ��A�G���[�̏ꍇ�� FALSE  ��Ԃ��܂��B
    '�y�����z
    '  ���̊֐��́A�t�@�C���̃u���b�N�f�[�^���������܂ꂽ���Ԃ�Ԃ��܂��B ����́A�t�@�C���̓��e���ύX���ꂽ�ۂ̎��Ԃł��B
    '=======================================================================
    Public Function filemtime(filename)

        Dim f
        filename = fileMapPath(filename)
        set f = fso.GetFile(filename)
        filemtime = f.DateLastModified

    End Function

    '*************************************
    Private Function fileMapPath(filename)

        Dim tmp
        tmp = Left(filename,3)
        tmp = Lcase(tmp)
        If tmp <> "d:\" and tmp <> "c:\" and left(filename,7) <> "http://" then
                fileMapPath = Server.MapPath(filename)
        Else
            fileMapPath = filename
        End If

    End Function

    '=======================================================================
    ' �t�@�C���̃T�C�Y���擾����
    '=======================================================================
    '�y�����z
    '  filename = string   �t�@�C���ւ̃p�X�B
    '�y�߂�l�z
    '  �t�@�C���̃T�C�Y��Ԃ��A�G���[�̏ꍇ�� FALSE ��Ԃ��܂� (�܂� E_WARNING ���x���̃G���[�𔭐������܂�) �B
    '�y�����z
    '  �w�肵���t�@�C���̃T�C�Y���擾���܂��B
    '=======================================================================
    Public Function filesize(filename)

        Dim f
        filename = fileMapPath(filename)
        set f = fso.GetFile(filename)
        filesize = f.Size

    End Function

    '=======================================================================
    ' �t�@�C���^�C�v���擾����
    '=======================================================================
    '�y�����z
    '  filename = string   �t�@�C���ւ̃p�X�B
    '�y�߂�l�z
    '  �t�@�C���̃^�C�v��Ԃ��܂��B
    '�y�����z
    '  �w�肵���t�@�C���̃^�C�v��Ԃ��܂��B
    '=======================================================================
    Public Function filetype(filename)

        Dim f
        filename = fileMapPath(filename)
        set f = fso.GetFile(filename)
        filetype = f.Type

    End Function

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

    '=======================================================================
    ' �s�� CSV �`���Ƀt�H�[�}�b�g���A�t�@�C���|�C���^�ɏ�������
    '=======================================================================
    '�y�����z
    '  fields  =  string    �l�̔z��B
    '  delimiter  =  string �I�v�V������ delimiter  �̓t�B�[���h��؂蕶�� (�ꕶ������) ���w�肵�܂��B�f�t�H���g�̓J���} (,) �ł��B
    '  enclosure  =  string �I�v�V������ enclosure  �̓t�B�[���h���͂ޕ��� (�ꕶ������) ���w�肵�܂��B�f�t�H���g�͓�d���p�� (") �ł��B
    '�y�߂�l�z
    '  �������񂾕�����̒�����Ԃ��܂��B���s�����ꍇ�� FALSE ��Ԃ��܂��B
    '�y�����z
    '  fputcsv() �́A�s�ifields  �z��Ƃ��ēn���ꂽ���́j�� CSV �Ƃ��ăt�H�[�}�b�g���A����� �t�@�C���ɏ������݂܂� (�����΂�Ō�ɉ��s��ǉ����܂�)�B
    '=======================================================================
    Public function fputcsv(fields,delimiter,enclosure)

        fputcsv = false
        If len(delimiter) = 0 Then delimiter = ","
        If len(enclosure) = 0 Then enclosure = """"

        Dim key,replaced
        For key = 0 to uBound(fields)
            replaced = false
            If inStr(fields(key),delimiter) or inStr(fields(key),enclosure) or inStr(fields(key),vbCrLf) Then
                fields(key) = Replace(fields(key),enclosure,enclosure & enclosure)
                fields(key) = enclosure & fields(key) & enclosure
            End If
        Next

        Dim str : str = join(fields,delimiter)
        ts.WriteLine str
        fputcsv = len(str)
    end function

    '=======================================================================
    ' fwrite() �̃G�C���A�X
    '=======================================================================
    '�y�����z
    '  str  =  string �������ޕ�����B
    '�y�����z
    '  ���̊֐��͎��̊֐��̃G�C���A�X�ł��B fwrite().
    '=======================================================================
    Public function fputs(str)
        fputs = fwrite(str)
    end function

    '=======================================================================
    ' �o�C�i���Z�[�t�ȃt�@�C���������ݏ���
    '=======================================================================
    '�y�����z
    '  str     =  string   �������ޕ�����B
    '�y�߂�l�z
    '  
    '�y�����z
    '  string �̓��e�� �t�@�C���E�X�g���[���ɏ������݂܂��B
    '=======================================================================
    Public function fwrite(str)
        ts.WriteLine str
    end function

    '=======================================================================
    ' �t�@�C�����f�B���N�g�����ǂ����𒲂ׂ�
    '=======================================================================
    '�y�����z
    '  filename = string    �t�@�C���ւ̃p�X�Bfilename  �����΃p�X�̏ꍇ�́A���݂̍�ƃf�B���N�g������̑��΃p�X�Ƃ��ď������܂��B
    '�y�߂�l�z
    '  �t�@�C���������݂��āA�����ꂪ�f�B���N�g���ł���� TRUE�A����ȊO�̏ꍇ�� FALSE ��Ԃ��܂��B
    '�y�����z
    '  �w�肵���t�@�C�����f�B���N�g�����ǂ����𒲂ׂ܂��B
    '=======================================================================
    Public Function is_dir(filename)

        Dim fn
        is_dir = false
        fn = fileMapPath(filename)

        If fso.FolderExists(fn) Then is_dir = true

    End Function

    '=======================================================================
    ' �ʏ�t�@�C�����ǂ����𒲂ׂ�
    '=======================================================================
    '�y�����z
    '  filename = string    �t�@�C���ւ̃p�X�B
    '�y�߂�l�z
    '  �t�@�C�������݂��A�����ꂪ�ʏ�̃t�@�C���ł���ꍇ�� TRUE�A ����ȊO�̏ꍇ�� FALSE ��Ԃ��܂��B
    '�y�����z
    '  �w�肵���t�@�C�����ʏ�̃t�@�C�����ǂ����𒲂ׂ܂��B
    '=======================================================================
    Public Function is_file(ByVal filename)

        is_file = false
        filename = fileMapPath(filename)
        If fso.FileExists(filename) Then is_file = true

    End Function

    '=======================================================================
    ' �f�B���N�g�������
    '=======================================================================
    '�y�����z
    '  pathname = string    �f�B���N�g���̃p�X�B
    '�y�߂�l�z
    '  ���������ꍇ�� TRUE ���A���s�����ꍇ�� FALSE ��Ԃ��܂��B
    '�y�����z
    '  �w�肵���f�B���N�g�����쐬���܂��B
    '=======================================================================
    Public Function mkdir(ByVal pathname)

        mkdir = false
        pathname = fileMapPath(pathname)
        If not fso.FolderExists(pathname) Then
            mkdir = fso.CreateFolder(pathname)
        End If

    End Function

    '=======================================================================
    ' �t�@�C���p�X�Ɋւ������Ԃ�
    '=======================================================================
    '�y�����z
    '  path     = string    ���ׂ����p�X�B
    '  options  = string    �ǂ̗v�f��Ԃ��̂����I�v�V�����̃p�����[�^ options  �Ŏw�肵�܂��B����� PATHINFO_DIRNAME�A PATHINFO_BASENAME�A PATHINFO_EXTENSION ����� PATHINFO_FILENAME �̑g�ݍ��킹�ƂȂ�܂��B �f�t�H���g�ł͂��ׂĂ̗v�f��Ԃ��܂��B
    '�y�߂�l�z
    '   �ȉ��̗v�f���܂ޘA�z�z���Ԃ��܂��B dirname (�f�B���N�g����)�Abasename (�t�@�C����) �����āA�������݂���� extension (�g���q)�B
    '   options ���g�p����ƁA ���ׂĂ̗v�f��I�����Ȃ����肱�̊֐��̕Ԃ�l�͕�����ƂȂ�܂��B 
    '�y�����z
    '  pathinfo() �́Apath  �Ɋւ������L����A�z�z���Ԃ��܂��B
    '=======================================================================
    Public Function pathinfo(path,options)

        Dim obj : set obj = Server.CreateObject("Scripting.Dictionary")
        Dim tmp


        obj("dirname") = dirname(path)
        obj("basename") = basename(path,"")
        obj("extension") = obj("basename")

        If inStr(obj("basename"),".") Then
            tmp = Split(obj("basename"),".")
            obj("extension") = tmp( uBound(tmp) )
        End if

        obj("filename") = Replace(obj("basename"),"." & obj("extension"),"")

        If len(options) > 0 Then

            If options = PATHINFO_DIRNAME Then
                pathinfo = obj("dirname")
            ElseIf options = PATHINFO_BASENAME Then
                pathinfo = obj("basename")
            ElseIf options = PATHINFO_EXTENSION Then
                pathinfo = obj("extension")
            ElseIf options = PATHINFO_FILENAME Then
                pathinfo = obj("filename")
            End if
            Exit Function
        End If

        set pathinfo = obj
    End Function

    '=======================================================================
    ' �f�B���N�g�����폜����
    '=======================================================================
    '�y�����z
    '  dirname     = string    �f�B���N�g���ւ̃p�X�B
    '�y�߂�l�z
    '   ���������ꍇ�� TRUE ���A���s�����ꍇ�� FALSE ��Ԃ��܂��B
    '�y�����z
    '  dirname �Ŏw�肳�ꂽ�f�B���N�g���� �폜���悤�Ǝ��݂܂��B
    '  �f�B���N�g���͋�łȂ��Ă͂Ȃ炸�A�܂� �K�؂ȃp�[�~�b�V�������ݒ肳��Ă��Ȃ���΂Ȃ�܂���B
    '=======================================================================
    Public Function rmdir(ByVal dirname)

        dirname = fileMapPath(dirname)
        fso.DeleteFolder dirname
        rmdir = true

    End Function
    
    '=======================================================================
    ' �t�@�C�����폜����
    '=======================================================================
    '�y�����z
    '  filename     = string    �t�@�C���ւ̃p�X�B
    '�y�߂�l�z
    '   ���������ꍇ�� TRUE ���A���s�����ꍇ�� FALSE ��Ԃ��܂��B
    '�y�����z
    '  filename  ���폜���܂��B Unix C ����̊֐� unlink() �Ɠ���͓����ł��B
    '=======================================================================
    Public Function unlink(ByVal filename)

        filename = fileMapPath(filename)
        fso.DeleteFile filename
        unlink = true

    End Function

End Class