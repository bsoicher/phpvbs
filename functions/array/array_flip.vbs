'=======================================================================
'�z��̃L�[�ƒl�𔽓]����
'=======================================================================
'�y�����z
'  trans    = array  ���]���s���L�[/�l�̑g�B
'�y�߂�l�z
'  ���������ꍇ�ɔ��]�����z��A���s�����ꍇ�� ��̃I�u�W�F�N�g ��Ԃ��܂��B
'�y�����z
'  �E�z��𔽓]���ĕԂ��܂��B
'  �E���Ȃ킿�Atrans  �̃L�[���l�ƂȂ�A trans  �̒l���L�[�ƂȂ�܂��B
'=======================================================================
Function array_flip(trans)

    Dim aryObj : set aryObj = Server.CreateObject("Scripting.Dictionary")

    If Not isArray(trans) and Not isObject(trans) Then
        set array_flip = aryObj
        Exit Function
    End If


    If isArray(trans) Then

        Dim i
        For i = 0 to uBound(trans)
            aryObj( trans(i) ) = i
        Next

    Elseif isObject(trans) Then

        Dim j
        For Each j In trans
            aryObj( trans(j) ) = j
        Next

    End If

    set array_flip = aryObj
End Function