'=======================================================================
'�ϐ��Ɋ܂܂��v�f�A ���邢�̓I�u�W�F�N�g�Ɋ܂܂��v���p�e�B�̐��𐔂���
'=======================================================================
'�y�����z
'  var    = mixed   �z��B
'  mode   = int     �I�v�V������mode  ������ COUNT_RECURSIVE  (�܂��� 1) �ɃZ�b�g���ꂽ�ꍇ�Acount()  �͍ċA�I�ɔz����J�E���g���܂��B
'�y�߂�l�z
'   var �Ɋ܂܂��v�f�̐���Ԃ��܂��B ���̂��̂ɂ́A1�̗v�f��������܂���̂ŁA�ʏ� var  �͔z��ł��B
'   ���� var ���z��������� Countable �C���^�[�t�F�[�X�����������I�u�W�F�N�g�ł͂Ȃ��ꍇ�A 1 ���Ԃ���܂��B
'   �ЂƂ�O������Avar �� NULL �̏ꍇ�A 0 ���Ԃ���܂��B 
'�y�����z
'  �E�ϐ��Ɋ܂܂��v�f�A ���邢�̓I�u�W�F�N�g�Ɋ܂܂��v���p�e�B�̐��𐔂��܂��B
'=======================================================================

Const COUNT_RECURSIVE = 1

Function count(var,mode)

    If not isArray(var) and not isObject(var) Then
        If isNull(var) Then
            count = 0
        Else
            count = 1
        End If
        Exit Function
    End If

    If mode <> COUNT_RECURSIVE Then

        If isArray(var) Then
            count = uBound(var) + 1
        ElseIf isObject(var) Then
            count = var.Count
        End If
        Exit Function

    Else

        Dim key,output : output = 0
        If isArray(var) Then
            For key = 0 to uBound(var)
                If isArray(var(key)) or isObject(var(key)) Then output = output + 1
                output = output + count(var(key),mode)
            Next
        ElseIf isObject(var) Then
            For Each key In var
                If isArray(var(key)) or isObject(var(key)) Then output = output + 1
                output = output + count(var(key),mode)
            Next
        End If

        count = output
    End If

End Function