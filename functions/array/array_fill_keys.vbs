'=======================================================================
'�L�[���w�肵�āA�z���l�Ŗ��߂�
'=======================================================================
'�y�����z
'  keys     = array     �L�[�Ƃ��Ďg�p����l�̔z��B
'  val      = string    �����񂩁A���邢�͒l�̔z��B
'�y�߂�l�z
'  �l�𖄂߂��z���Ԃ��܂��B
'�y�����z
'  �E�p�����[�^ val  �Ŏw�肵���l�Ŕz��𖄂߂܂��B 
'  �E�L�[�Ƃ��āA�z�� keys  �Ŏw�肵���l���g�p���܂��B
'=======================================================================
Function array_fill_keys(keys, val)

    Dim ary_fill,i
    set ary_fill = Server.CreateObject("Scripting.Dictionary")
    set array_fill_keys = ary_fill
    if Not isArray(keys) then Exit Function
    If isArray(val) Then
        If uBound(val) > uBound(keys) then Exit Function
    End If

    If IsArray(val) Then
        For i = 0 to uBound(keys)
            If Not ary_fill.Exists( keys(i) ) Then ary_fill.Add keys(i), val(i)
        Next
    Else
        For i = 0 to uBound(keys)
            If Not ary_fill.Exists( keys(i) ) Then ary_fill.Add keys(i), val
        Next
    End If

    set array_fill_keys = ary_fill

End Function