'=======================================================================
' �t�@�C���|�C���^����s���擾���ACSV�t�B�[���h����������
'=======================================================================
'�y�����z
'  
'�y�߂�l�z
'  �ǂݍ��񂾃t�B�[���h�̓��e���܂ސ��l�Y���z���Ԃ��܂��B
'�y�����z
'  fgets() �ɓ���͎��Ă��܂����A fgetcsv() �͍s�� CSV  �t�H�[�}�b�g�̃t�B�[���h�Ƃ��ēǍ��ݏ������s���A �ǂݍ��񂾃t�B�[���h���܂ޔz���Ԃ��Ƃ����Ⴂ������܂��B
' File_System class(https://github.com/masayukiando/phpvbs/blob/master/functions/filesystem/FileSystem.class.vbs)���ł̏���
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