'=======================================================================
'�v�f���t���ɂ����z���Ԃ�
'=======================================================================
'�y�����z
'  ary              = Array ���͂̔z��B
'�y�߂�l�z
'  �E�t�]�������z���Ԃ��܂��B
'�y�����z
'  �E�z����󂯎���āA�v�f�̏��Ԃ��t�ɂ����V�����z���Ԃ��܂��B
'=======================================================================
Function array_reverse(mAry)

    Dim arr_len,i

    If isArray(mAry) Then

        Dim tmp_ar()
        Dim newkey

        arr_len = uBound(mAry)
        ReDim tmp_ar(arr_len)

        For i = 0 to arr_len
            newkey = arr_len -i
            tmp_ar(i) = mAry(newkey)
        Next

        array_reverse = tmp_ar

    ElseIf isObject(mAry) Then

        Dim tmpObj,j,cnt

        cnt = 0
        set tmpObj = Server.CreateObject("Scripting.Dictionary")
        arr_len = mAry.Count-1

        ReDim index_values(arr_len),index_keys(arr_len)

        For Each j In mAry
            index_values(cnt) = mAry(j)
            index_keys(cnt)   = j
            cnt = cnt + 1
        Next

        For i = cnt-1 To 0 Step -1

            If Not tmpObj.Exists(Cstr(index_keys(i))) Then
                tmpObj.add Cstr(index_keys(i)),index_values(i)
            End if
        Next

        set array_reverse = tmpObj
    End If
End Function