'=======================================================================
'�z��̋��ʍ����v�Z����
'=======================================================================
'�y�����z
'  mAry1    = array  �l�𒲂ׂ���ƂƂȂ�z��B
'  mAry2    = array  �l���r����ΏۂƂȂ�z��B
'�y�߂�l�z
'  mAry1 �̒l�̂����A���ׂĂ̈����ɑ��݂���l�̂��̂��܂ޔz���Ԃ��܂��B
'�y�����z
'  �E���̑S�Ă̈����ɑ��݂��� mAry1 �̒l��S�ėL����z���Ԃ��܂��B
'=======================================================================
Function array_intersect(mAry1,mAry2)

    Dim key
    Dim output : set output = Server.CreateObject("Scripting.Dictionary")

    If isArray(mAry1) Then
        For key = 0 to uBound(mAry1)
            If len(array_search(mAry1(key),mAry2,false)) > 0 Then
                output.Add key, mAry1(key)
            End If
        Next
    ElseIf isObject(mAry1) Then
        For Each key In mAry1
            If len(array_search(mAry1(key),mAry2,false)) > 0 Then
                output.Add key, mAry1(key)
            End If
        Next
    End If

    set array_intersect = output

End Function