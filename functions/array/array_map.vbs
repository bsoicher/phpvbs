'=======================================================================
'�w�肵���z��̗v�f�ɃR�[���o�b�N�֐���K�p����
'=======================================================================
'�y�����z
'  callback = callback  �z��̊e�v�f�ɓK�p����R�[���o�b�N�֐��B
'  arr      = array     �R�[���o�b�N�֐���K�p����z��B
'�y�߂�l�z
'  arr �̊e�v�f�� callback  �֐���K�p������A ���̑S�Ă̗v�f���܂ޔz���Ԃ��܂��B
'�y�����z
'  �Earr �̊e�v�f�� callback  �֐���K�p���܂��B
'=======================================================================
Function array_map(callback, arr)

    Dim key
    Dim tmp_ar

    If isArray( arr ) Then

        If Len( callback ) = 0 Then
            array_map = arr
            Exit Function
        End If

        ReDim tmp_ar( uBound(arr) )
        For key = 0 to uBound( arr )
            If isObject( arr(key) ) Then
                execute("set tmp_ar(key) = " & callback & "(arr(key))")
            Else
                execute("tmp_ar(key) = " & callback & "(arr(key))")
            End If
        Next

        array_map = tmp_ar

    ElseIf isObject( arr ) Then

        If Len( callback ) = 0 Then
            set array_map = arr
            Exit Function
        End If

        Dim return_val

        set tmp_ar = Server.CreateObject("Scripting.Dictionary")
        For Each key In arr
            return_val = ""
            If isObject( arr.Item(key) ) Then
                execute("set return_val = " & callback & "(arr.Item(key))")
            Else
                execute("return_val = " & callback & "(arr.Item(key))")
            End If
            tmp_ar.Add key, return_val
        Next

        set array_map = tmp_ar

    End If

End Function