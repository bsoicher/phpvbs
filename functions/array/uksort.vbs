'=======================================================================
'���[�U��`�̔�r�֐���p���āA�L�[�Ŕz����\�[�g����
'=======================================================================
'�y�����z
'  ary          = Array   ���͂̔z��B
'  cmp_function = int     ��r�p�̃R�[���o�b�N�֐��B�֐� cmp_function �́A array �̃L�[�y�A�ɂ���Ė�������� 2 �̃p�����[�^���󂯎��܂��B ���̔�r�֐����Ԃ��l�́A�ŏ��̈�������Ԗڂ�菬�����ꍇ�͕��̐��A �������ꍇ�̓[���A�����đ傫���ꍇ�͐��̐��łȂ���΂Ȃ�܂���B
'�y�߂�l�z
'  ���������ꍇ�� TRUE ���A���s�����ꍇ�� FALSE ��Ԃ��܂��B
'�y�����z
'  �Euksort() �́A ���[�U��`�̔�r�֐���p���Ĕz��̃L�[���\�[�g���܂��B
'  �E�\�[�g�������z��𕡎G�Ȋ�Ń\�[�g����K�v������ꍇ�ɂ́A ���̊֐����g���K�v������܂��B
'=======================================================================
Function uksort(ByRef arr, cmp_function)

    uksort = false
    If Not IsObject(arr) Then  Exit Function

    Dim key,keys
    Dim new_arr : set new_arr = Server.CreateObject("Scripting.Dictionary")

    keys = array_keys(arr,"",false)
    usort keys,cmp_function

    For Each key In keys
        new_arr.Add key, arr(key)
    Next

    set arr = new_arr

    uksort = true

End Function