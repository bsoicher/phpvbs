'=======================================================================
'�z��̈ꕔ��W�J����
'=======================================================================
'�y�����z
'  mAry     = Array ���͂̔z��B
'  offset   = int   offset  �����̒l�ł͂Ȃ��ꍇ�A�v�f�ʒu�̌v�Z�́A �z�� array  �� offset ����n�߂��܂��B offset  �����̏ꍇ�A�v�f�ʒu�̌v�Z�� array  �̍Ōォ��s���܂��B
'  level    = int  level ���w�肳��A���̏ꍇ�A �A�����镡���̗v�f���Ԃ���܂��Blevel ���w�肳��A���̏ꍇ�A�z��̖�������A�����镡���̗v�f���Ԃ���܂��B �ȗ����ꂽ�ꍇ�Aoffset  ����z��̍Ō�܂ł̑S�Ă̗v�f���Ԃ���܂��B
'�y�߂�l�z
'  �E�؂�����������Ԃ��܂��B
'�y�����z
'  �EmAry ������� offset  ����� level �Ŏw�肳�ꂽ�A������v�f��Ԃ��܂��B
'=======================================================================
Function array_slice(mAry,offset,level)

    array_slice = false

    If Not isArray(mAry) Then Exit Function
    If Not isNumeric(offset) Then Exit Function
    If Not isNumeric(level) Then level = uBound(mAry)

    Dim s,e,arynum
    arynum = uBound(mAry)

    If offset >= 0 Then _
        s = offset _
    Else _
        s = arynum + offset + 1

    If level >= 0 Then _
        e = s + level _
    Else _
        e = arynum + level

    If e > arynum Then e = arynum

    Dim i,counter
    counter = 0
    ReDim tmp_ar(e-s)
    For i = s to e
        tmp_ar(counter) = mAry(i)
        counter = counter +1
    Next

    array_slice = tmp_ar

End Function