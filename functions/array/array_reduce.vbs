'=======================================================================
'�R�[���o�b�N�֐���p���Ĕz��𕁒ʂ̒l�ɕύX���邱�Ƃɂ��A�z����ċA�I�Ɍ��炷
'=======================================================================
'�y�����z
'  mAry     = array     ���͂̔z��B
'  callback = callback  �R�[���o�b�N�֐��B
'  initial  = int       �I�v�V������ intial �����p�\�ȏꍇ�A�����̍ŏ��Ŏg�p���ꂽ��A �z�񂪋�̏ꍇ�̍ŏI���ʂƂ��Ďg�p����܂��B
'�y�߂�l�z
'   ���ʂ̒l��Ԃ��܂��B
'   �z�񂪋�� initial ���n����Ȃ������ꍇ�́A array_reduce() �� NULL ��Ԃ��܂��B 
'�y�����z
'  �E�z�� mAry �̊e�v�f�� callback �֐����J��Ԃ��K�p���A �z�����̒l�Ɍ��炵�܂��B
'=======================================================================
Function array_reduce(ByVal mAry, callback, ByVal initial)

    array_reduce = null
    If len( initial ) > 0 Then array_reduce = initial
    If not isArray( mAry ) and not isObject( mAry ) Then Exit Function

    Dim acc : acc = initial
    Dim key

    If isObject( mAry ) Then
        For Each key In mAry
            execute("acc = " & callback & "(acc, mAry(key))")
        Next

    ElseIf isArray( mAry ) Then

        Dim lon : lon = uBound( mAry )
        For key = 0 to lon
            execute("acc = " & callback & "(acc, mAry(key))")
        Next
    End If

    array_reduce = acc

End Function