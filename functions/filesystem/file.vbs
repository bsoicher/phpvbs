'=======================================================================
' �t�@�C���S�̂�ǂݍ���Ŕz��Ɋi�[����
'=======================================================================
'�y�����z
'  filename  = string �t�@�C���ւ̃p�X�B
'  flags     = int    �I�v�V�����̃p�����[�^ flags  �́A�ȉ��̒萔�̂����̂ЂƂA���邢�͕����̑g�ݍ��킹�ƂȂ�܂��B
'�y�߂�l�z
'  �t�@�C����z��ɓ���ĕԂ��܂��B �z��̊e�v�f�̓t�@�C���̊e�s�ɑΉ����܂��B
'  ���s�L���͂����܂܂ƂȂ�܂��B ���s����� file() �� FALSE ��Ԃ��܂��B
'�y�����z
'  �t�@�C���S�̂�z��ɓǂݍ��݂܂��B
' File_System class(http://phpvbs.verygoodtown.com/)�K�{
'=======================================================================

Const FILE_IGNORE_NEW_LINES = 2
Const FILE_SKIP_EMPTY_LINES = 4

Function file(filename,flags)

    Dim req,tmp,key
    req = file_get_contents(filename)

    If flags = FILE_SKIP_EMPTY_LINES Then
        var_dump req
        tmp = preg_replace("/^" & vbCrLf & "/is","",req,"","")
        tmp = Split(tmp,vbCrLf)
    Else
        tmp = Split(req,vbCrLf)

        If flags = FILE_IGNORE_NEW_LINES Then
            For key = 0 to uBound(tmp)
                tmp(key) = tmp(key) & vbCrLf
            Next
        End If
    End If

    file = tmp

End Function