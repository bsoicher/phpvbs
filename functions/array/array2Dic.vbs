'=======================================================================
'�z����f�B�N�V���i���ɕϊ�����
'=======================================================================
'�y�����z
'  arr  = array  �z��
'�y�߂�l�z
'  �f�B�N�V���i���I�u�W�F�N�g�B
'�y�����z
'  �E�n���ꂽ�z��� �f�B�N�V���i���I�u�W�F�N�g�ɕϊ����܂��B
'=======================================================================
Function array2Dic(ByVal myAry)

    Dim i,tmpObj
    set tmpObj = Server.CreateObject("Scripting.Dictionary")
    For i = 0 to uBound(myAry)
        tmpObj.add i, myAry(i)
    Next
    set array2Dic = tmpObj

End Function