'=======================================================================
'���[�U���w�肵���R�[���o�b�N�֐��𗘗p���A �ǉ����ꂽ�Y���̊m�F���܂߂Ĕz��̍����v�Z����
'=======================================================================
'�y�����z
'  mAry1                = array     ��r���̔z��B
'  mAry2                = array     ��r����ΏۂƂȂ�z��B
'  date_compare_func    = callback  �g�p����R�[���o�b�N�֐��B ���̊֐��́A1 �߂̈����� 2 �߂�菬���� / ������ / �傫�� �ꍇ�ɂ��ꂼ�� ���̐� / �[�� / ���̐� ��Ԃ��K�v������܂��B
'  key_compare_func     = callback  �L�[�i�Y���j�̔�r�́A�R�[���o�b�N�֐�
'�y�߂�l�z
'  mAry1  �̗v�f�̂����A ���̑��̔z��̂�����ɂ��܂܂�Ȃ����̂������c�����z���Ԃ��܂��B
'�y�����z
'  �EmAry1  �� mAry2 �Ɣ�r���A���̍���Ԃ��܂��B
'  �E���[�U���w�肵���R�[���o�b�N�֐���p���ēY�����r���܂��B
'=======================================================================
Function array_udiff_uassoc(ByVal mAry1,ByVal mAry2, data_compare_func,key_compare_func)

    Dim arr_dif
    set arr_dif = Server.CreateObject("Scripting.Dictionary")

    If Not isObject(mAry1) or Not isObject(mAry2) Then
        set array_diff_assoc = arr_dif : Exit Function
    End If

    Dim j,k,key_result,data_result,found
    For Each j in mAry1

        found = false
        For Each k In mAry2
            execute("key_result  = " & key_compare_func & "(j,k)")
            execute("data_result = " & data_compare_func & "(mAry1(j),mAry2(k))")

            If key_result = 0 and data_result = 0 Then
                found = true
                Exit For
            End If
        Next

        If Not found Then
             arr_dif.Add j , mAry1(j)
        End If
    Next

    set array_udiff_uassoc = arr_dif

End Function
