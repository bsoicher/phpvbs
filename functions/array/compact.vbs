'=======================================================================
'�ϐ����Ƃ��̒l����z����쐬����
'=======================================================================
'�y�����z
'  varname    = mixed   �ϐ����̔z��Ƃ��邱�Ƃ��ł��܂��B
'�y�߂�l�z
'  �ǉ����ꂽ�S�Ă̕ϐ���l�Ƃ���o�͔z���Ԃ��܂��B
'�y�����z
'  �E�����Ƃ��̒l����z����쐬���܂��B
'  �E�e�����ɂ��āAcompact() �͌��݂̃V���{���e�[�u���ɂ����Ă��̖��O��L����ϐ���T���A �ϐ������L�[�A�ϐ��̒l�����̃L�[�Ɋւ���l�ƂȂ�悤�ɒǉ����܂��B
'=======================================================================
Function compact(varname)

    If Not isArray(varname) Then Exit Function

    Dim output : set output = Server.CreateObject("Scripting.Dictionary")
    Dim var,code

    For Each var In varname
        code = "If output.Exists(var) Then" & vbCrLf & _
                "   output.Item(var) = " & var & vbCrLf & _
                "Else" & vbCrLf & _
                "   output.Add var, " & var & vbCrLf & _
                "End If" & vbCrLf
        execute (code)

    Next

    set compact = output

End Function