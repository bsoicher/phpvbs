'=======================================================================
'�f�[�^�̔�r�ɃR�[���o�b�N�֐���p���A �ǉ����ꂽ�Y���̊m�F���܂߂Ĕz��̍����v�Z����
'=======================================================================
'�y�����z
'  mAry1     = Array �ŏ��̔z��B
'  mAry2     = Array 2 �Ԗڂ̔z��B
'  mAry1     = callback ��r�p�̃R�[���o�b�N�֐��B���[�U���w�肵���R�[���o�b�N�֐���p���ăf�[�^�̔�r���s���܂��B ���̊֐��́A1 �߂̈����� 2 �߂�菬���� / ������ / �傫�� �ꍇ�ɂ��ꂼ�� ���̐� / �[�� / ���̐� ��Ԃ��K�v������܂��B

'�y�߂�l�z
'  �E���̈����̂�����ɂ����݂��Ȃ� mAry1 �̒l�̑S�Ă�L����z���Ԃ��܂��B
'�y�����z
'  �E�f�[�^�̔�r�ɃR�[���o�b�N�֐���p���A�z��̍����v�Z���܂��B 
'  �E���̊֐��� array_diff_assoc() �ƈقȂ�A �f�[�^�̔�r�ɓ����֐��𗘗p���܂��B
'=======================================================================
Function array_udiff_assoc(ByVal mAry1,ByVal mAry2, data_compare_func)

    Dim arr_udif
    set arr_udif = Server.CreateObject("Scripting.Dictionary")

    If Not isObject(mAry1) then
        set array_diff_uassoc = arr_udif
        Exit Function
    End If

    Dim key,found
    For Each key In mAry1
        arr_udif.Add key, mAry1(key)
    Next

    If Not isObject(mAry2) Then Exit Function

    For Each key In mAry2
        If arr_udif.Exists( key ) Then
            execute("found = " & data_compare_func & "(arr_udif(key), mAry2(key))")
            If found = 0 Then
                If arr_udif.Exists( key ) Then arr_udif.Remove key
            ElseIf found < 0 Then
                If isObject(mAry2(key)) Then
                    set arr_udif.Item( key ) = mAry2(key)
                Else
                    arr_udif.Item( key ) = mAry2(key)
                End If
            End if
        End If
    Next

    set array_udiff_assoc = arr_udif

End Function