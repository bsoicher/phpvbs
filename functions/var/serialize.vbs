'=======================================================================
'�l�̕ۑ��\�ȕ\���𐶐�����
'=======================================================================
'�y�����z
'  val   = mixed    �V���A��������l�B
'�y�߂�l�z
'  val  �̕ۑ��\�ȃo�C�g�X�g���[���\�����܂ޕ������Ԃ��܂��B 
'�y�����z
'  �E �l�̕ۑ��\�ȕ\���𐶐����܂��B
'  �E �^��\�������킸�� ASP �̒l��ۑ��܂��͓n���ۂɗL�p�ł��B
'  �E �V���A�������ꂽ������� ASP �̒l�ɖ߂��ɂ́A unserialize() ���g�p���Ă��������B 
'=======================================================================
Function serialize(ByVal val)

    Dim strstrType
    strType = getType(val)

    Dim str
    Dim cnt : cnt = 0
    Dim strs
    Dim key

    Select Case strType

        Case "vbEnpty","vbNull"
            str = "N"
        Case "vbBoolean"
            str = "b:" & [?](val,1,0)
        Case "vbInteger","vbLong","vbSingle","vbDouble","vbCurrency"
            str = [?]([==](int(val),val),"i","d") & ":" & val
        Case "vbDate","vbString","vbVariant"
            str = "s:" & len(val) & ":""" & val & """"
        Case "vbArray"
            str = "a"

            For key = 0 to uBound(val)
                strs = strs & serialize(key) & _
                        serialize(val(key))
                cnt = cnt + 1
            Next
            str = str & ":" & cnt & ":{" & strs & "}"
            str = str & ";"

        Case "vbObject"
            str = "O"

            For Each key In val
                strs = strs & serialize(key) & _
                        serialize(val(key))
                cnt = cnt + 1
            Next
            str = str & ":" & cnt & ":{" & strs & "}"

        Case Else
            'empty
    End Select

    If strType <> "vbArray" AND strType <> "vbObject" Then
        str = str & ";"
    End If

    serialize = str
End Function