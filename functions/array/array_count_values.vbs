'=======================================================================
'�z��̒l�̐��𐔂���
'=======================================================================
'�y�����z
'  mAry     = array  �l�𐔂���z��B
'�y�߂�l�z
'   mAry �̃L�[�Ƃ��̓o��񐔂�g�ݍ��킹���A�z�z����쐬���܂��B
'�y�����z
'  �E�z�� mAry �̒l���L�[�Ƃ��AmAry �ɂ����邻�̒l�̏o���񐔂�l�Ƃ����z���Ԃ��܂��B
'�y�G���[ / ��O�z
'  �Estring ���邢�� integer �ȊO�̌^�̗v�f���o�ꂷ��ƒv���I�ȃG���[���������܂��B
'=======================================================================
Function array_count_values(mAry)

    Dim obj
    set obj = Server.CreateObject("Scripting.Dictionary")
    if Not isArray(mAry) and Not isObject(mAry) Then
        Set array_count_values = obj
        Exit Function
    End If

    Dim j,k
    Dim intCounter


    For Each j In mAry

        intCounter = 0

        For Each k In mAry
            If j = k Then intCounter = intCounter + 1
        Next

        If Not obj.Exists(j) Then obj.Add j, intCounter

    Next

    Set array_count_values = obj

End Function