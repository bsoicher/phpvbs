'=======================================================================
'�ǉ����ꂽ�Y���̊m�F���܂߂Ĕz��̍����v�Z����
'=======================================================================
'�y�����z
'  mAry1    = array  ��r���̔z��B
'  mAry2    = array  ��r����ΏۂƂȂ�z��B
'�y�߂�l�z
'  mAry1  �̗v�f�̂����A ���̑��̔z��̂�����ɂ��܂܂�Ȃ����̂������c�����z���Ԃ��܂��B
'�y�����z
'  �EmAry1  �� mAry2 �Ɣ�r���A���̍���Ԃ��܂��B
'  �Earray_diff() �Ƃ͈قȂ�A �z��̃L�[��p���Ĕ�r���s���܂��B
'=======================================================================
Function array_diff_assoc(ByVal mAry1,ByVal mAry2)

    Dim retAry
    set retAry = Server.CreateObject("Scripting.Dictionary")

    If Not isObject(mAry1) or Not isObject(mAry2) Then
        set array_diff_assoc = retAry : Exit Function
    End If

    Dim j,k
    For Each j in mAry1

        retAry.Add j, mAry1(j)

        For Each k In mAry2
            if j = k and mAry1(j) = mAry2(k) Then retAry.Remove k
        Next
    Next

    set array_diff_assoc = retAry

End Function