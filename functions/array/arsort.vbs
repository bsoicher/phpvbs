'=======================================================================
'�A�z�L�[�Ɨv�f�Ƃ̊֌W���ێ����z����t���Ƀ\�[�g����
'=======================================================================
'�y�����z
'  ary        = Array   ���͂̔z��B
'  sort_flags = int     �I�v�V�����̃p�����[�^ sort_flags  �ɂ��\�[�g�̓�����C���\�ł��B�ڍׂɂ��ẮA sort() ���Q�Ƃ��������B
'�y�߂�l�z
'  ���������ꍇ�� TRUE ���A���s�����ꍇ�� FALSE ��Ԃ��܂��B
'�y�����z
'  �E���̊֐��́A �A�z�z��ɂ����Ċe�z��̃L�[�Ɨv�f�Ƃ̊֌W���ێ����z����\�[�g���܂��B
'  �E���̊֐��́A ��Ɏ��ۂ̗v�f�̕��ѕ����d�v�ł���A�z�z����\�[�g���邽�߂Ɏg���܂��B
'=======================================================================
Function arsort(ByRef arr, sort_flags)

    arsort = false
    If Not IsObject(arr) Then  Exit Function

    Dim key,keys
    Dim new_arr : set new_arr = Server.CreateObject("Scripting.Dictionary")
    Dim found
    Dim cnt

    keys = array_values(arr)
    rsort keys,sort_flags

    For Each key In keys
        found = array_keys(arr,key,true)
        If isArray(found) Then
            For cnt = 0 to uBound(found)
                If Not new_arr.Exists(found(cnt)) Then
                    new_arr.Add found(cnt), arr(found(cnt))
                End If
            Next
        Else
            new_arr.Add found, arr(found)
        End If
    Next

    set arr = new_arr

    arsort = true

End Function
