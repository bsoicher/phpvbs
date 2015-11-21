'=======================================================================
' (IPv4) �C���^�[�l�b�g�A�h���X���C���^�[�l�b�g�W���h�b�g�\�L�ɕϊ�����
'=======================================================================
'�y�����z
'  proper_address = string   �������`���̃A�h���X�B
'�y�߂�l�z
'  �C���^�[�l�b�g�� IP �A�h���X��\���������Ԃ��܂��B
'�y�����z
'  �E�֐�long2ip() �́A�K�؂ȃA�h���X�\������h�b�g�\�L (��:aaa.bbb.ccc.ddd)�̃C���^�[�l�b�g�A�h���X�𐶐����܂��B
'=======================================================================
Function long2ip( proper_address )

    long2ip = false

    If isNumeric(proper_address) Then
        If proper_address >= 0 and proper_address <= 4294967295 Then

            long2ip = ( proper_address / pow(256,3) ) & "." & _
                      ( (proper_address Mod pow(256,3)) / pow(256,2) ) & "." & _
                      ( ((proper_address Mod pow(256,3)) / pow(256,2)) / pow(256,1) ) & "." & _
                      ( (((proper_address Mod pow(256,3)) / pow(256,2)) / pow(256,1)) / pow(256,0) )

        End If
    End If
End Function