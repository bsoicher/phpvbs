'=======================================================================
' 1 �ȏ�̕�������o�͂���
'=======================================================================
'�y�����z
'  str  = string    �o�͂������p�����[�^�B
'�y�߂�l�z
'  �l��Ԃ��܂���B
'�y�����z
'  �E���ׂẴp�����[�^���o�͂��܂��B
'=======================================================================
Sub echo(str)

    If isObject(str) then
        Response.Write "Object"
    ElseIf IsArray(str) then
        Response.Write "Array"
    Else
        Response.Write str
    End if

End Sub