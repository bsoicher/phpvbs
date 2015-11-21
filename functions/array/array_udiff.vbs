'=======================================================================
'�f�[�^�̔�r�ɃR�[���o�b�N�֐���p���A�z��̍����v�Z����
'=======================================================================
'�y�����z
'  mAry1     = Array �ŏ��̔z��B
'  mAry2     = Array 2 �Ԗڂ̔z��B
'  mAry1     = callback ��r�p�̃R�[���o�b�N�֐��B���[�U���w�肵���R�[���o�b�N�֐���p���ăf�[�^�̔�r���s���܂��B ���̊֐��́A1 �߂̈����� 2 �߂�菬���� / ������ / �傫�� �ꍇ�ɂ��ꂼ�� ���̐� / �[�� / ���̐� ��Ԃ��K�v������܂��B

'�y�߂�l�z
'  �E���̈����̂�����ɂ����݂��Ȃ� mAry1 �̒l�̑S�Ă�L����z���Ԃ��܂��B
'�y�����z
'  �E�f�[�^�̔�r�ɃR�[���o�b�N�֐���p���A�z��̍����v�Z���܂��B 
'  �E���̊֐��� array_diff() �ƈقȂ�A �f�[�^�̔�r�ɓ����֐��𗘗p���܂��B
'=======================================================================
Function array_udiff(mAry1,mAry2,data_compare_func)

    Dim arr_udif,key_c,key,found

    If Not isObject(mAry1) or Not isObject(mAry2) Then
        set array_diff_assoc = retAry : Exit Function
    End If

    If isArray(mAry1) and isArray(mAry2) Then

        For Each key In mAry1

            found = 0
            For Each key_c In mAry2
                execute("found = " & data_compare_func & "(key, key_c)")
                If found <> 0 Then
                    Exit For
                End If
            Next

            If found > 0 Then
                [] arr_udif, mAry1(key)
            ElseIf found < 0 Then
                [] arr_udif, mAry2(key_c)
            End If
        Next

        array_udiff = arr_udif

    ElseIf isObject(mAry1) and isObject(mAry2) Then

        set arr_udif = Server.CreateObject("Scripting.Dictionary")

        For Each key In mAry1

            found = 0
            For Each key_c In mAry2
                execute("found = " & data_compare_func & "(mAry1(key), mAry2(key_c))")
                If found <> 0 Then
                    Exit For
                End If
            Next

            If found > 0 Then
                If arr_udif.Exists(key) Then
                    arr_udif.Item(key) = mAry1(key)
                Else
                    arr_udif.Add key, mAry1(key)
                End If
            ElseIf found < 0 Then
                If arr_udif.Exists(key_c) Then
                    arr_udif.Item(key_c) = mAry2(key_c)
                Else
                    arr_udif.Add key_c, mAry2(key_c)
                End If
            End If

        Next

        set array_udiff = arr_udif

    End If

End Function

