'=======================================================================
'�z��̑S�Ă̒l��Ԃ�
'=======================================================================
'�y�����z
'  mAry     = array �z��B
'�y�߂�l�z
'  ���l�Y���̒l�̔z���Ԃ��܂��B
'�y�����z
'  �E�z�񂩂�S�Ă̒l�����o���A���l�Y���������z���Ԃ��܂��B
'=======================================================================
Function array_values(mAry)

    Dim tmp_ar
    Dim key,counter : counter= 0

    If isObject(mAry) Then

        ReDim tmp_ar(mAry.Count -1)

        For Each key In mAry
            If isObject(mAry(key)) Then
                set tmp_ar(counter) = mAry(key)
            Else
                tmp_ar(counter) = mAry(key)
            End if
            counter = counter + 1
        Next

    ElseIf isArray(mAry) Then
        tmp_ar = mAry
    End If

    array_values = tmp_ar

End Function