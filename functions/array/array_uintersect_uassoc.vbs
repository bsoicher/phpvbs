'=======================================================================
'�f�[�^�ƓY���̔�r�ɃR�[���o�b�N�֐���p���A �ǉ����ꂽ�Y���̊m�F���܂߂Ĕz��̋��ʍ����v�Z����
'=======================================================================
'�y�����z
'  mAry1                = Array     �ŏ��̔z��B
'  mAry2                = Array     2 �Ԗڂ̔z��B
'  data_compare_func    = callback  ��r�p�̃R�[���o�b�N�֐��B��r�́A���[�U���w�肵���R�[���o�b�N�֐��𗘗p���čs���܂��B ���̊֐��́A1 �߂̈����� 2 �߂�菬���� / ������ / �傫�� �ꍇ�ɂ��ꂼ�� ���̐� / �[�� / ���̐� ��Ԃ��K�v������܂��B
'  key_compare_func     = callback  �L�[�̔�r�p�̃R�[���o�b�N�֐��B
'�y�߂�l�z
'  �E���̑S�Ă̈����Ɍ���� mAry1 �̒l���܂ޔz���Ԃ��܂��B
'�y�����z
'  �E�f�[�^�̔�r�ɃR�[���o�b�N�֐���p���A�z��̋��ʍ����v�Z���܂��B
'  �E�L�[����r�Ɏg�p����邱�Ƃɒ��ӂ��Ă��������B 
'  �E�f�[�^�ƓY���́A���ꂼ��ʂ̃R�[���o�b�N�֐���p���Ĕ�r����܂��B
'=======================================================================
Function array_uintersect_uassoc(mAry1,mAry2,data_compare_func,key_compare_func)

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
    Dim found,key_found
    Dim output : set output = Server.CreateObject("Scripting.Dictionary")

    If isArray(mAry1) Then

        For key = 0 to uBound(mAry1)

            If isArray(mAry2) Then
                For key_c = 0 to uBound(mAry2)
                    execute("found = " & data_compare_func & "(mAry1(key), mAry2(key_c))")
                    execute("key_found = " & key_compare_func & "(key, key_c)")
                    If found = 0 and key_found = 0 Then
                        If output.Exists(key) Then
                            output(key) = mAry1(key)
                        Else
                            output.Add key, mAry1(key)
                        End If
                    End If
                Next

            ElseIf isObject(mAry2) Then

                For Each key_c In mAry2
                    execute("found = " & data_compare_func & "(mAry1(key), mAry2(key_c))")
                    execute("key_found = " & key_compare_func & "(key, key_c)")
                    If found = 0 and key_found = 0 Then
                        If output.Exists(key) Then
                            output(key) = mAry1(key)
                        Else
                            output.Add key, mAry1(key)
                        End If
                    End If
                Next
            End If
        Next

    ElseIf isObject(mAry1) Then

        For Each key In mAry1
            If isArray(mAry2) Then
                For key_c = 0 to uBound(mAry2)
                    execute("found = " & data_compare_func & "(mAry1(key), mAry2(key_c))")
                    execute("key_found = " & key_compare_func & "(key, key_c)")
                    If found = 0 and key_found = 0 Then
                        If output.Exists(key) Then
                            output(key) = mAry1(key)
                        Else
                            output.Add key, mAry1(key)
                        End If
                    End If
                Next

            ElseIf isObject(mAry2) Then

                For Each key_c In mAry2
                    execute("found = " & data_compare_func & "(mAry1(key), mAry2(key_c))")
                    execute("key_found = " & key_compare_func & "(key, key_c)")
                    If found = 0 and key_found = 0 Then
                        If output.Exists(key) Then
                            output(key) = mAry1(key)
                        Else
                            output.Add key, mAry1(key)
                        End If
                    End If
                Next
            End If
        Next

    End If

    set array_uintersect_uassoc = output

End Function
