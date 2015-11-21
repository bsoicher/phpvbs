'=======================================================================
'�z��𕪊�����
'=======================================================================
'�y�����z
'  mAry     = Array         �������s���z��B
'  size     = int           �e�����̃T�C�Y�B
'  preserve_keys = bool     TRUE �̏ꍇ�̓L�[�����̂܂ܕێ����܂��B �f�t�H���g�� FALSE �ŁA�e�����̃L�[�����炽�߂Đ����ŐU��Ȃ����܂��B
'�y�߂�l�z
'  ���l�Y���̑������z���Ԃ��܂��B�Y�����̓[������n�܂�A �e�����̗v�f���� size  �ƂȂ�܂��B
'�y�����z
'  �E�z����Asize  ���̗v�f�ɕ������܂��B 
'  �E�Ō�̕����̗v�f���� size  ��菬�����Ȃ邱�Ƃ�����܂��B
'=======================================================================
Function array_chunk(mAry,size)

    If not isNumeric(size) Then Exit Function
    If size < 1 then Exit Function

    Dim x,i,c : x = 0 : c = -1
    Dim l : l = uBound(mAry)
    Dim n : n = int(l / size)
    ReDim tmpAry(n)

    For i = 0 to l
        x = i Mod size

        If x >= 1 Then
            If isObject(mAry(i)) Then
                set tmpAry(c)(x) = mAry(i)
            Else
                tmpAry(c)(x) = mAry(i)
            End If
        Else
            c = c +1
            If n <> c Then
                toReDim tmpAry(c),size -1
            Else
                toReDim tmpAry(c),l -i
            End If

            If isObject(mAry(i)) Then
                set tmpAry(c)(0) = mAry(i)
            Else
                tmpAry(c)(0) = mAry(i)
            End If

        End If
    Next

    array_chunk = tmpAry

End Function