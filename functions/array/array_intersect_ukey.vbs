'=======================================================================
'�L�[����ɂ��A�R�[���o�b�N�֐���p���� �z��̋��ʍ����v�Z����
'=======================================================================
'�y�����z
'  mAry1    = array  ��r���ƂȂ�ŏ��̔z��B
'  mAry2    = array  �L�[���r����ΏۂƂȂ�ŏ��̔z��B
'  key_compare_func = callback  ��r�Ɏg�p����A���[�U��`�̃R�[���o�b�N�֐��B
'�y�߂�l�z
'  mAry1 �̒l�̂����A ���ׂĂ̈����ɑ��݂���L�[�̂��̂��܂ޘA�z�z���Ԃ��܂��B
'�y�����z
'  �E���̑S�Ă̈����ɑ��݂��� mAry1 �̒l��S�ėL����z���Ԃ��܂��B
'=======================================================================
Function array_intersect_ukey(mAry1,mAry2,key_compare_func)

    Dim result : set result = Server.CreateObject("Scripting.Dictionary")
    Dim key,k,compare

    If isArray(mAry1) Then
        For key = 0 to uBound(mAry1)

            If isArray(mAry2) Then
                For k = 0 to uBound(mAry2)
                    execute("compare = " & key_compare_func & "(key,k)")
                    If compare = 0 Then
                        result.Add key, mARy1(key)
                        Exit For
                    End If
                Next
            ElseIf isObject(mAry2) Then
                For Each k In mAry2
                    execute("compare = " & key_compare_func & "(key,k)")
                    If compare = 0 Then
                        result.Add key, mARy1(key)
                        Exit For
                    End If
                Next
            End If

        Next
    ElseIf isObject(mAry1) Then
        For Each key In mAry1

            If isArray(mAry2) Then
                For k = 0 to uBound(mAry2)
                    execute("compare = " & key_compare_func & "(key,k)")
                    If compare = 0 Then
                        result.Add key, mARy1(key)
                        Exit For
                    End If
                Next
            ElseIf isObject(mAry2) Then
                For Each k In mAry2
                    execute("compare = " & key_compare_func & "(key,k)")
                    If compare = 0 Then
                        result.Add key, mARy1(key)
                        Exit For
                    End If
                Next
            End If

        Next
    End If

    set array_intersect_ukey = result

End Function