'=======================================================================
'�z��̑S�Ă̗v�f�ɁA���[�U�[�֐����ċA�I�ɓK�p����
'=======================================================================
'�y�����z
'  arr      = array     ���͂̔z��B
'  callback = callback  �������Ƃ�܂��B array  �p�����[�^�̒l���ŏ��̈����A �L�[/�Y���͓�Ԗڂ̈����ƂȂ�܂��Bfuncname  �ɂ��z��̒l���̂��̂�ύX����K�v������ꍇ�A funcname  �̍ŏ��̈����� �Q��  �Ƃ��ēn���K�v������܂��B���̏ꍇ�A�z��̗v�f�ɉ������ύX�́A �z�񎩑̂ɑ΂��čs���܂��B 
'  userdata = array     userdata  �p�����[�^���w�肳�ꂽ�ꍇ�A �R�[���o�b�N�֐� funcname  �ւ̎O�Ԗڂ̈����Ƃ��ēn����܂��B
'�y�߂�l�z
'  ���������ꍇ�� TRUE ��Ԃ��܂��B
'�y�����z
'  �Earr �̊e�v�f�� callback  �֐���K�p���܂��B
'  �E���̊֐��͔z��̗v�f�����ċA�I�ɂ��ǂ��Ă����܂��B
'=======================================================================
Function array_walk_recursive(ByRef arr, callback, userdata)

    Dim key
    Dim thisCall

    If Len( callback ) = 0 Then Exit Function

    If isArray( arr ) Then

        For key = 0 to uBound( arr )
            If isArray(arr(key)) or isObject(arr(key)) Then
                thisCall = "array_walk_recursive"
            Else
                thisCall = callback
            End If

            execute("call " & callback & "(arr(key),key,userdata)")
        Next

    ElseIf isObject( arr ) Then

        For Each key In arr
            If isArray(arr(key)) or isObject(arr(key)) Then
                thisCall = "array_walk_recursive"
            Else
                thisCall = callback
            End If

            execute("call " & callback & "(arr(key),key,userdata)")
        Next

    End If

    array_walk_recursive = true

End Function