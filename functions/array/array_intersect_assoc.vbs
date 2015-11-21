'=======================================================================
'�ǉ����ꂽ�Y���̊m�F���܂߂Ĕz��̋��ʍ����m�F����
'=======================================================================
'�y�����z
'  mAry1    = array  �l�𒲂ׂ���ƂƂȂ�z��B
'  mAry2    = array  �l���r����ΏۂƂȂ�z��B
'�y�߂�l�z
'  mAry1 �̒l�̂����A���ׂĂ̈����ɑ��݂�����̂��܂ޘA�z�z���Ԃ��܂��B
'�y�����z
'  �E�S�Ă̈����Ɍ���� mAry1 �̑S�Ă̒l���܂ޔz���Ԃ��܂��B 
'  �Earray_intersect() �ƈقȂ�A �L�[����r�Ɏg�p����邱�Ƃɒ��ӂ��Ă��������B
'=======================================================================
Function array_intersect_assoc(mAry1,mAry2)

    Dim intersect : set intersect = Server.CreateObject("Scripting.Dictionary")
    Dim key,counter

    If isArray(mAry2) Then
        counter = uBound(mAry2)
    Else
        counter = null
    End If

    If isArray(mAry1) Then
        For key = 0 to uBound(mAry1)
            intersect.Add key, mAry1(key)
            If counter >= key or isNull(counter) Then
                If mAry2(key) <> mAry1(key) Then
                    intersect.Remove key
                End If
            End If
        Next
    ElseIf isObject(mAry1) Then
        For Each key In mAry1
            intersect.Add key, mAry1(key)
            If isNull(counter) or (isNumeric(key) and counter >= key) Then
                If mAry2(key) <> mAry1(key) Then
                    intersect.Remove key
                End If
            Else
               intersect.Remove key
            End If
        Next
    End If

    set array_intersect_assoc = intersect

End Function