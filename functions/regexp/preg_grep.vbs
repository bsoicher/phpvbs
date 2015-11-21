'=======================================================================
'�p�^�[���Ƀ}�b�`����z��̗v�f��Ԃ�
'=======================================================================
'�y�����z
'  pattern  = string    ��������p�^�[����\��������B
'  input    = array     ���͂̔z��B
'  flags    = array     PREG_GREP_INVERT  ��ݒ肷��ƁA���̊֐��� �^���� pattern  �Ƀ}�b�` ���Ȃ�  �v�f��Ԃ��܂��B
'�y�߂�l�z
'  input  �z��̃L�[���g�p�����z���Ԃ��܂��B
'�y�����z
'  �E input  �z��̗v�f�̂����A �w�肵�� pattern  �Ƀ}�b�`������̂�v�f�Ƃ���z���Ԃ��܂��B 
'=======================================================================
Const PREG_GREP_INVERT    = 1
Function preg_grep(pattern, input, flags)

    Dim obj
    set obj = Server.CreateObject("Scripting.Dictionary")

    If not isArray(input) and not isObject(input) Then
        set preg_grep =  obj
        Exit Function
    End If

    Dim key
    If isArray(input) Then
        For key = 0 to uBound(input)
            If flags = PREG_GREP_INVERT Then
                If Not preg_match(pattern, input(key),"","","") Then
                    obj.Add key, input(key)
                End If
            Else
                If preg_match(pattern, input(key),"","","") Then
                    obj.Add key, input(key)
                End If
            End If
        Next
    ElseIf isObject(input) Then
        For Each key In input
            If flags = PREG_GREP_INVERT Then
                If Not preg_match(pattern, input(key),"","","") Then
                    obj.Add key, input(key)
                End If
            Else
                If preg_match(pattern, input(key),"","","") Then
                    obj.Add key, input(key)
                End If
            End If
        Next
    End If

    set preg_grep =  obj

End Function