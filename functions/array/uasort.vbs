'=======================================================================
'���[�U��`�̔�r�֐��Ŕz����\�[�g���A�A�z�C���f�b�N�X��ێ�����
'=======================================================================
'�y�����z
'  ary          = Array   ���͂̔z��B
'  cmp_function = int     ���[�U��`�̔�r�֐��̗�ɂ��ẮA usort() ����� uksort()  ���Q�Ƃ��������B
'�y�߂�l�z
'  ���������ꍇ�� TRUE ���A���s�����ꍇ�� FALSE ��Ԃ��܂��B
'�y�����z
'  �E ���̊֐��́A�z��C���f�b�N�X���֘A����z��v�f�Ƃ̊֌W��ێ�����悤�Ȕz����\�[�g���܂��B
'  �E ��Ɏ��ۂ̔z��̏����ɈӖ�������A�z�z����\�[�g���邽�߂ɂ��̊֐��͎g�p����܂��B 
'=======================================================================
Function uasort(ByRef arr, cmp_function)

    uasort = false
    If Not IsObject(arr) Then  Exit Function

    Dim key,keys
    Dim new_arr : set new_arr = Server.CreateObject("Scripting.Dictionary")
    Dim found
    Dim cnt

    keys = array_values(arr)
    usort keys,cmp_function

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

    uasort = true

End Function