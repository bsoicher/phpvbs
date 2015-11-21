'=======================================================================
'�ǉ����ꂽ�Y���̊m�F���܂߁A�R�[���o�b�N�֐���p���� �z��̋��ʍ����m�F����
'=======================================================================
'�y�����z
'  mAry1            = array     ��r���ƂȂ�ŏ��̔z��B
'  mAry2            = array     �L�[���r����ΏۂƂȂ�ŏ��̔z��B
'  key_compare_func = callback  ��r�Ɏg�p����A���[�U��`�̃R�[���o�b�N�֐��B
'�y�߂�l�z
'  mAry1 �̒l�̂����A ���ׂĂ̈����ɑ��݂�����݂̂̂�Ԃ��܂��B
'�y�����z
'  �E�S�Ă̈����Ɍ���� mAry1 �̑S�Ă̒l���܂ޔz���Ԃ��܂��B array_intersect() �ƈقȂ�A �L�[����r�Ɏg�p����邱�Ƃɒ��ӂ��Ă��������B
'  �E��r�́A���[�U���w�肵���R�[���o�b�N�֐��𗘗p���čs���܂��B 
'  �E���̊֐��́A1 �߂̈����� 2 �߂�菬���� / ������ / �傫�� �ꍇ�ɂ��ꂼ�� ���̐� / �[�� / ���̐� ��Ԃ��K�v������܂��B
'=======================================================================
Function array_intersect_uassoc(mAry1,mAry2,key_compare_func)

    Dim result : set result = Server.CreateObject("Scripting.Dictionary")
    Dim key,k,found,compare

    If isArray(mAry1) Then
        For key = 0 to uBound(mAry1)
            found = false

            If isArray(mAry2) Then
                For k = 0 to uBound(mAry2)
                    execute("compare = " & key_compare_func & "(key,k)")
                    If compare = 0 and mAry1(key) = mAry2(k) Then
                        found = true
                        Exit For
                    End If
                Next
            ElseIf isObject(mAry2) Then
                For Each k In mAry2
                    execute("compare = " & key_compare_func & "(key,k)")
                    If compare = 0 and mAry1(key) = mAry2(k) Then
                        found = true
                        Exit For
                    End If
                Next
            End If

            If found = true Then
                result.Add k, mAry1(key)
            End if
        Next
    ElseIf isObject(mAry1) Then
        For Each key In mAry1
            found = false

            If isArray(mAry2) Then
                For k = 0 to uBound(mAry2)
                    execute("compare = " & key_compare_func & "(key,k)")
                    If compare = 0 and mAry1(key) = mAry2(k) Then
                        found = true
                        Exit For
                    End If
                Next
            ElseIf isObject(mAry2) Then
                For Each k In mAry2
                    execute("compare = " & key_compare_func & "(key,k)")
                    If compare = 0 and mAry1(key) = mAry2(k) Then
                        found = true
                        Exit For
                    End If
                Next
            End If

            If found = true Then
                result.Add k, mAry1(key)
            End if
        Next
    End If

    set array_intersect_uassoc = result

End Function