'=======================================================================
'�ϐ��̌^���擾����
'=======================================================================
'�y�����z
'  str  = mixed  �^�𒲂ׂ����ϐ��B
'�y�߂�l�z
'  �^�̕������Ԃ��܂��B
'�y�����z
'  �E�ϐ� str �̌^��Ԃ��܂��B
'=======================================================================
Function gettype(s)

    Select Case VarType(s)
    Case 0
        gettype = "vbEnpty"
    Case 1
        gettype = "vbNull"
    Case 2
        gettype = "vbInteger"
    Case 3
        gettype = "vbLong"
    Case 4
        gettype = "vbSingle"
    Case 5
        gettype = "vbDouble"
    Case 6
        gettype = "vbCurrency"
    Case 7
        gettype = "vbDate"
    Case 8
        gettype = "vbString"
    Case 9
        gettype = "vbObject"
    Case 10
        gettype = "vbError"
    Case 11
        gettype = "vbBoolean"
    Case 12
        gettype = "vbVariant"
    Case 13
        gettype = "vbDataObject"
    Case 17
        gettype = "vbByte"
    Case 8192
        gettype = "vbArray"
    Case 8204
        gettype = "vbArray"
    Case 8209
        gettype = "vbBinary"
    End Select

End Function