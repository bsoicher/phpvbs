'=======================================================================
'�z����L�[�Ń\�[�g����
'=======================================================================
'�y�����z
'  ary        = Array   ���͂̔z��B
'  sort_flags = int     �I�v�V�����̃p�����[�^ sort_flags  �ɂ��\�[�g�̓�����C���\�ł��B�ڍׂɂ��ẮA sort() ���Q�Ƃ��������B
'�y�߂�l�z
'  ���������ꍇ�� TRUE ���A���s�����ꍇ�� FALSE ��Ԃ��܂��B
'�y�����z
'  �E�L�[�ƃf�[�^�̊֌W���ێ����A�z����L�[�Ń\�[�g���܂��B 
'  �E���̊֐��́A��Ƃ��ĘA�z�z��ɂ����ėL�p�ł��B
'=======================================================================
Function ksort(ByRef arr, sort_flags)

    ksort = false
    If Not IsObject(arr) Then  Exit Function

    Dim key,keys
    Dim new_arr : set new_arr = Server.CreateObject("Scripting.Dictionary")

    keys = array_keys(arr,"",false)
    sort keys,sort_flags

    For Each key In keys
        new_arr.Add key, arr(key)
    Next

    set arr = new_arr

    ksort = true

End Function