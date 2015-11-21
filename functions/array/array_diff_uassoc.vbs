'=======================================================================
'���[�U���w�肵���R�[���o�b�N�֐��𗘗p���A �ǉ����ꂽ�Y���̊m�F���܂߂Ĕz��̍����v�Z����
'=======================================================================
'�y�����z
'  mAry1            = array     ��r���̔z��B
'  mAry2            = array     ��r����ΏۂƂȂ�z��B
'  key_compare_func = callback  �g�p����R�[���o�b�N�֐��B ���̊֐��́A1 �߂̈����� 2 �߂�菬���� / ������ / �傫�� �ꍇ�ɂ��ꂼ�� ���̐� / �[�� / ���̐� ��Ԃ��K�v������܂��B
'�y�߂�l�z
'  mAry1  �̗v�f�̂����A ���̑��̔z��̂�����ɂ��܂܂�Ȃ����̂������c�����z���Ԃ��܂��B
'�y�����z
'  �EmAry1  �� mAry2 �Ɣ�r���A���̍���Ԃ��܂��B
'  �E���[�U���w�肵���R�[���o�b�N�֐���p���ēY�����r���܂��B
'=======================================================================
Function array_diff_uassoc(ByVal mAry1,ByVal mAry2,key_compare_func)

    Dim arr_dif
    set arr_dif = Server.CreateObject("Scripting.Dictionary")

    If Not isObject(mAry1) or Not isObject(mAry2) Then
        set array_diff_assoc = arr_dif : Exit Function
    End If

    Dim j,k,callback_ret
    For Each j in mAry1

        arr_dif.Add j, mAry1(j)

        For Each k In mAry2
            If mAry1(j) = mAry2(k) Then
                execute("callback_ret = " & key_compare_func & "(j,k)")
                If callback_ret = 0 Then
                    If arr_dif.Exists(j) Then arr_dif.Remove j
                ElseIf callback_ret < 0 Then
                    arr_dif.Remove j
                    If arr_dif.Exists(k) Then
                        arr_dif.Item( k ) = mAry2(k)
                    Else
                        arr_dif.Add k ,mAry2(k)
                    End If
                End If
            End If
        Next
    Next

    set array_diff_uassoc = arr_dif

End Function