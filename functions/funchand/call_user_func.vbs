'=======================================================================
'�ŏ��̈����Ŏw�肵�����[�U�֐����R�[������
'=======================================================================
'�y�����z
'  callback     = mixed  �R�[������֐��B���̃p�����[�^�� array(classname, methodname) ���w�肷�邱�Ƃɂ��A �N���X���\�b�h���ÓI�ɃR�[�����邱�Ƃ��ł��܂��B
'  parameter    = mixed  ���̊֐��ɓn���A�[���ȏ�̃p�����[�^�B
'�y�߂�l�z
'  �֐��̌��ʁA���邢�̓G���[���� FALSE ��Ԃ��܂��B
'�y�����z
'  �E�p�����[�^ callback �Ŏw�肵�� ���[�U��`�̃R�[���o�b�N�֐����R�[�����܂��B
'=======================================================================
Function call_user_func(callback,parameter)

    Dim thisFunc,retval
    If isArray(callback) Then
        thisFunc  = callback(0) & "." & callback(1)
    Else
        thisFunc = callback
    End If

    execute("retval = " & thisFunc & "(parameter)")
    call_user_func = retval
End Function