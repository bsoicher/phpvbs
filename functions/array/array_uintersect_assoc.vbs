'=======================================================================
'�f�[�^�̔�r�ɃR�[���o�b�N�֐���p���A�z��̋��ʍ����v�Z����
'=======================================================================
'�y�����z
'  mAry1                = Array     �ŏ��̔z��B
'  mAry2                = Array     2 �Ԗڂ̔z��B
'  data_compare_func    = callback  ��r�p�̃R�[���o�b�N�֐��B��r�́A���[�U���w�肵���R�[���o�b�N�֐��𗘗p���čs���܂��B ���̊֐��́A1 �߂̈����� 2 �߂�菬���� / ������ / �傫�� �ꍇ�ɂ��ꂼ�� ���̐� / �[�� / ���̐� ��Ԃ��K�v������܂��B
'�y�߂�l�z
'  �E���̑S�Ă̈����ɑ��݂��� mAry1 �̒l��S�ėL����z���Ԃ��܂��B
'�y�����z
'  �E�f�[�^�̔�r�ɃR�[���o�b�N�֐���p���A�z��̋��ʍ����v�Z���܂��B
'=======================================================================
Function array_uintersect_assoc(mAry1,mAry2,data_compare_func)
'Callback�̗�
'function rmul(v, w)
'    rmul = 0
'    If isObject(v) or isArray(v) Then
'        rmul = 1
'    Elseif isObject(w) or isArray(w) Then
'        rmul = 1
'    End If
'    If rmul = 1 then Exit FUnction
'    If v = w Then
'        rmul = 0
'    Else
'        rmul = 1
'    End If
'End Function

    Dim key,key_c
    Dim found
    Dim output : set output = Server.CreateObject("Scripting.Dictionary")

    If isArray(mAry1) Then

        For key = 0 to uBound(mAry1)

            If array_key_exists(key, mAry2) Then
                execute("found = " & data_compare_func & "(mAry1(key), mAry2(key))")
                If found = 0 Then
                    If output.Exists(key) Then
                        output(key) = mAry1(key)
                    Else
                        output.Add key, mAry1(key)
                    End If
                End If
            End If
        Next

    ElseIf isObject(mAry1) Then

        For Each key In mAry1
            If array_key_exists(key, mAry2) Then
                execute("found = " & data_compare_func & "(mAry1(key), mAry2(key))")
                If found = 0 Then

                    If output.Exists(key) Then
                        output(key) = mAry1(key)
                    Else
                        output.Add key, mAry1(key)
                    End If
                End If
            End If
        Next

    End If

    set array_uintersect_assoc = output

End Function