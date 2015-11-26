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
' File_System class(https://github.com/masayukiando/phpvbs/blob/master/functions/filesystem/FileSystem.class.vbs)���ł̏���
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