'=======================================================================
'�z��̃L�[�����ׂĕԂ�
'=======================================================================
'�y�����z
'  mAry         = array  �Ԃ��L�[���܂ޔz��B
'  search_value = mixed  �w�肵���ꍇ�́A�����̒l���܂ރL�[�݂̂�Ԃ��܂��B
'  strict       = mixed  �������Ɍ^��r���s���܂��B
'�y�߂�l�z
'  mAry �̂��ׂẴL�[��z��ŕԂ��܂��B
'�y�����z
'  �E�z�� mAry ����S�ẴL�[ (���l����ѕ�����) ��Ԃ��܂��B
'  �E�I�v�V���� search_value  ���w�肳�ꂽ�ꍇ�A �w�肵���l�Ɋւ���L�[�݂̂��Ԃ���܂��B
'  �E�w�肳��Ȃ��ꍇ�́AmAry ����S�ẴL�[���Ԃ���܂��B
'  �Estrict  �p�����[�^���g���āA ��r�Ɍ^����r���邱�Ƃ��ł��܂��B
'=======================================================================
Function array_keys(mAry,search_value,strict)

    Dim tmp_arr
    Dim key
    Dim include
    Dim addArr
    Dim cnt : cnt = 0

    addArr = true
    If [==](search_value,empty) Then
        addArr = false
        ReDim tmp_arr( count(mAry,0)-1 )
    End If

    If isObject( mAry ) Then

        For Each key In mAry
            include = true
            If [!=](search_value,empty) Then
                If strict = true Then
                    If [!=](mAry(key) , search_value) or (varType(mAry(key)) <> varType(search_value)) Then
                        include = false
                    End If
                ElseIf [!=](mAry(key) , search_value) Then
                    include = false
                End If
            End If

            If include = true Then
                If addArr Then
                    [] tmp_arr, key
                Else
                    tmp_arr(cnt) = key
                    cnt = cnt + 1
                End If
            End If
        Next

    ElseIf isArray(mAry) Then

        For cnt = 0 to uBound(mAry)

            include = true
            If [!=](search_value,empty) Then

                If strict = true Then
                    If [!=](mAry(cnt) , search_value) or (varType(mAry(cnt)) <> varType(search_value)) Then
                        include = false
                    End If
                ElseIf [!=](mAry(cnt) , search_value) Then
                    include = false
                End If
            End If

            If include = true Then
                If addArr Then
                    [] tmp_arr, cnt
                Else
                    tmp_arr(cnt) = cnt
                End If
            End if
        Next
    End If

    array_keys = tmp_arr

End Function