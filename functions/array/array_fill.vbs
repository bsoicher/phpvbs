'=======================================================================
'�z����w�肵���l�Ŗ��߂�
'=======================================================================
'�y�����z
'  start_index  = int       �Ԃ����z��̍ŏ��̃C���f�b�N�X�B
'  num          = int       �}������v�f���B
'  val          = string    �v�f�Ɏg�p����l�B
'�y�߂�l�z
'  �l�𖄂߂��z���Ԃ��܂��B
'�y�����z
'  �E�p�����[�^ value  ��l�Ƃ��� num  �̃G���g������Ȃ�z��𖄂߂܂��B 
'  �E���̍ہA�L�[�́Astart_index  �p�����[�^����J�n���܂��B
'=======================================================================
Function array_fill(start_index, num, val)

    If Not isNumeric(num) or num < 1 then Exit Function

    Dim intCounter,ary()
    Dim i

    intCounter = start_index + num -1
    ReDim ary(intCounter)

    For i = start_index to intCounter
        ary(i) = val
    Next

    array_fill = ary

End Function