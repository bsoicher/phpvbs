'=======================================================================
'�z��̍����v�Z����
'=======================================================================
'�y�����z
'  mAry1    = array  ��r���̔z��B
'  mAry2    = array  ��r����ΏۂƂȂ�z��B
'�y�߂�l�z
'  mAry1  �̗v�f�̂����A ���̑��̔z��̂�����ɂ��܂܂�Ȃ����̂������c�����z���Ԃ��܂��B
'�y�����z
'  �EmAry1  �� mAry2 �Ɣ�r���A���̍���Ԃ��܂��B
'  �E��̗v�f�́A(string) elem1 = (string) elem2  �̏ꍇ�̂ݓ������ƌ�������܂��B
'  �E����������ƁA������\���������ꍇ�ƂȂ�܂��B 
'=======================================================================
Function array_diff(ByVal mAry1,ByVal mAry2)

    Dim arr_dif,key_c,key,found
    set arr_dif = Server.CreateObject("Scripting.Dictionary")

    If isArray(mAry1) Then
        set mAry1 = array2Dic(mAry1)
    End If

    If isArray(mAry2) Then
        set mAry2 = array2Dic(mAry2)
    End If


    For Each key In mAry1

        found = false
        For Each key_c In mAry2
            If mAry1(key) = mAry2(key_c) Then
                found = true
                Exit For
            End If
        Next

        If Not found Then
            arr_dif.add key, mAry1(key)
        End If
    Next


    set array_diff = arr_dif

End Function