'=======================================================================
'�L�[����ɂ��Ĕz��̍����v�Z����
'=======================================================================
'�y�����z
'  mAry1    = array  ��r���̔z��B
'  mAry2    = array  ��r����ΏۂƂȂ�z��B
'�y�߂�l�z
'  mAry1  �̗v�f�̂����A ���̑��̔z��̂�����ɂ��܂܂�Ȃ����̂������c�����z���Ԃ��܂��B
'�y�����z
'  �EmAry1  �� mAry2 �Ɣ�r���A���̍���Ԃ��܂��B
'  �E���̊֐��� array_diff() �Ɏ��Ă��܂����A �l�ł͂Ȃ��L�[��p���Ĕ�r����Ƃ����_���قȂ�܂��B
'=======================================================================
Function array_diff_key(ByVal mAry1,ByVal mAry2)

    Dim arr_dif
    set arr_dif = Server.CreateObject("Scripting.Dictionary")

    If Not isObject(mAry1) then
        set array_diff_uassoc = arr_dif
        Exit Function
    End If

    Dim key
    For Each key In mAry1
        arr_dif.Add key, mAry1(key)
    Next

    If isObject(mAry2) Then
        For Each key In mAry2
            If arr_dif.Exists( key ) Then arr_dif.Remove key
        Next
    End If

    set array_diff_key = arr_dif

End Function