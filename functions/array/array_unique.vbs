'=======================================================================
'�z�񂩂�d�������l���폜����
'=======================================================================
'�y�����z
'  mAry     = Array ���͂̔z��B
'�y�߂�l�z
'  �E�����ς̔z���Ԃ��܂��B
'�y�����z
'  �E�l�ɏd���̂Ȃ��V�K�z���Ԃ��܂��B
'=======================================================================
Function array_unique(arr)

    Dim key,key_c
    Dim found
    Dim output : set output = Server.CreateObject("Scripting.Dictionary")

    If isArray(arr) Then
        For key = 0 to uBound(arr)
            found = array_search(arr(key),output,false)
            If found = false and varType(found) = 11 Then
                output.Add key, arr(key)
            End If
        Next
    ElseIf isObject(arr) Then
        For Each key In arr
            found = array_search(arr(key),output,false)
            If found = false and varType(found) = 11 Then
                    output.Add key, arr(key)
            End If
        next
    End If

    set array_unique = output
end Function