'=======================================================================
'�z��̒l�̐ς��v�Z����
'=======================================================================
'�y�����z
'  mAry     = array  �z��
'�y�߂�l�z
'  �ς��Ainteger ���邢�� float �ŕԂ��܂��B
'�y�����z
'  �E�z��̊e�v�f�̐ς��v�Z���܂��B
'=======================================================================
Function array_product(mAry)

    If Not isArray( mAry ) Then Exit Function

    Dim j,product
    product = 1

    For Each j In mAry
        If isNumeric(j) Then product = product * j
    Next

    array_product = product

End Function