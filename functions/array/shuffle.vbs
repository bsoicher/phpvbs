'=======================================================================
'�z����V���b�t������
'=======================================================================
'�y�����z
'  arr        = Array   �z��B
'�y�߂�l�z
'  ���������ꍇ�� TRUE ���A���s�����ꍇ�� FALSE ��Ԃ��܂��B
'�y�����z
'  �E���̊֐��́A�z����V���b�t�� (�v�f�̏��Ԃ������_����) ���܂��B
'=======================================================================
Function shuffle(ByRef arr)

    shuffle = false
    If not isArray(arr) Then Exit Function

    Dim key,j,x,i : i = count(arr,0)

    Randomize

    For key = 0 to uBound(arr)
        i = i -1
        j = Round(Rnd * i)
        [=] x , arr(i)
        [=] arr(i) , arr(j)
        [=] arr(j) , x
    Next

    shuffle = true

End Function
