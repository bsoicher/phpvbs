'=======================================================================
'�z��̑S�Ă̗v�f�Ƀ��[�U�֐���K�p����
'=======================================================================
'�y�����z
'  arr      = array     ���͂̔z��B
'  callback = callback  �������Ƃ�܂��B array  �p�����[�^�̒l���ŏ��̈����A �L�[/�Y���͓�Ԗڂ̈����ƂȂ�܂��Bfuncname  �ɂ��z��̒l���̂��̂�ύX����K�v������ꍇ�A funcname  �̍ŏ��̈����� �Q��  �Ƃ��ēn���K�v������܂��B���̏ꍇ�A�z��̗v�f�ɉ������ύX�́A �z�񎩑̂ɑ΂��čs���܂��B 
'  userdata = array     userdata  �p�����[�^���w�肳�ꂽ�ꍇ�A �R�[���o�b�N�֐� funcname  �ւ̎O�Ԗڂ̈����Ƃ��ēn����܂��B
'�y�߂�l�z
'  ���������ꍇ�� TRUE ��Ԃ��܂��B
'�y�����z
'  �Earr �̊e�v�f�� callback  �֐���K�p���܂��B
'=======================================================================
Function array_walk(ByRef arr, callback, userdata)

    Dim key

    If Len( callback ) = 0 Then Exit Function

    If isArray( arr ) Then

        For key = 0 to uBound( arr )
            execute("call " & callback & "(arr(key),key,userdata)")
        Next

    ElseIf isObject( arr ) Then

        Dim return_val

        For Each key In arr
            execute("call " & callback & "(arr.Item(key),key,userdata)")
        Next

    End If

    array_walk = true

End Function
