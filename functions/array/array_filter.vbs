'=======================================================================
'�z��̗v�f���t�B���^�����O����
'=======================================================================
'�y�����z
'  mAry         = aarray    ��������z��B
'  callback     = callback  �g�p����R�[���o�b�N�֐��B�R�[���o�b�N�֐����^�����Ȃ������ꍇ�A input �̃G���g���̒��� FALSE �Ɠ��������� (boolean �ւ̕ϊ� ���Q�Ƃ�������) �����ׂč폜����܂�

'�y�߂�l�z
'  �t�B���^�����O���ꂽ���ʂ̔z���Ԃ��܂��B
'�y�����z
'  �EmAry �̃G���g���̒��� FALSE �Ɠ��������� �����ׂč폜����܂��B
'=======================================================================
Function array_filter(ByRef mAry,callback)

    If isArray(mAry) Then

        Dim intCounter,i,strType,callback_ret
        intCounter = uBound(mAry)

        For i = 0 to intCounter
            callback_ret = true
            If Len( callback ) > 0 Then _
                execute("callback_ret = " & callback & "(mAry(i))")

            If callback_ret = true and ( mAry(i) = empty or isNull(mAry(i)) ) Then
                mAry = array_remove(mAry,i)
                call array_filter(mAry,callback)
                Exit For
            End If
        Next

    ElseIf isObject(mAry) Then
        Dim j
        For Each j IN mAry
            callback_ret = true
            If Len( callback ) > 0 Then _
                execute("callback_ret = " & callback & "(mAry(i))")

            If callback_ret = true and ( mAry(j) = empty or isNull(mAry(j)) ) Then _
                mAry.Remove j
        Next

    End If

    array_filter = true

End Function