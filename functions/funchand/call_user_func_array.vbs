'=======================================================================
'�p�����[�^�̔z����w�肵�ă��[�U�֐����R�[������
'=======================================================================
'�y�����z
'  callback     = mixed  �R�[������֐��B���̃p�����[�^�� array($classname, $methodname) ���w�肷�邱�Ƃɂ��A �N���X���\�b�h���ÓI�ɃR�[�����邱�Ƃ��ł��܂��B
'  param_arr    = array  �֐��ɓn���p�����[�^���w�肷��z��B
'�y�߂�l�z
'  �֐��̌��ʁA���邢�̓G���[���� FALSE ��Ԃ��܂��B
'�y�����z
'  �Eparam_arr  �Ƀp�����[�^���w�肵�āA function  �Ŏw�肵�����[�U��`�֐����R�[�����܂��B
'=======================================================================
Function call_user_func_array(callback,param_arr)

    Dim thisFunc,thisParam,retval
    If isArray(callback) Then
        thisFunc  = callback(0) & "." & callback(1)
    Else
        thisFunc = callback
    End If

    If isArray(param_arr) Then
        Dim key
        For Each key In parameter
            If len( thisParam ) > 0 Then
                thisParam = thisParam & "," & key
            Else
                thisParam = key
            End IF
        Next
    Else
        thisParam = param_arr
    End If
    execute("retval = " & thisFunc & "(" & thisParam & ")")
    call_user_func_array = retval

End Function