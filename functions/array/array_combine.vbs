'=======================================================================
'����̔z����L�[�Ƃ��āA��������̔z���l�Ƃ��āA�ЂƂ̔z��𐶐�����
'=======================================================================
'�y�����z
'  keys     = array  �L�[�Ƃ��Ďg�p����z��B
'  values   = array  �l�Ƃ��Ďg�p����z��B
'�y�߂�l�z
'  �쐬�����z���Ԃ��܂��B
'  �݂��̔z��̗v�f�̐������v���Ȃ��ꍇ�� FALSE ��Ԃ��܂��B
'�y�����z
'  �Ekeys  �z��̒l���L�[�Ƃ��āA
'  �E�܂� values  �z��̒l��Ή�����l�Ƃ��Đ������� �z�� ���쐬���܂��B
'=======================================================================
Function array_combine(keys,values)

    Dim obj
    set obj = Server.CreateObject("Scripting.Dictionary")

    If uBound(keys) <> uBound(values) Then
        set array_combine = obj
        Exit Function
    End If

    Dim i
    For i = 0 to uBound(keys)
        If obj.Exists( keys(i) ) Then
            obj.Item( keys(i) ) = values(i)
        Else
            obj.Add keys(i) , values(i)
        End If
    Next

    set array_combine = obj

End Function