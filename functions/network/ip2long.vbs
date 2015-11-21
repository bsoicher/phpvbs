'=======================================================================
' (IPv4) �C���^�[�l�b�g�v���g�R���h�b�g�\�L�̃A�h���X���A�K���ȃA�h���X��L���镶����ɕϊ�����
'=======================================================================
'�y�����z
'  ip_address = string   �W���`���̃A�h���X�B
'�y�߂�l�z
'  IPv4 �A�h���X�A���邢�� ip_address  ���s���Ȍ`���̏ꍇ�� FALSE ��Ԃ��܂��B
'�y�����z
'  �E�֐� ip2long() �́A�C���^�[�l�b�g�W���`�� (�h�b�g�\�L�̕�����) �\������ IPv4 �C���^�[�l�b�g�l�b�g�A�h���X�𐶐����܂��B
'=======================================================================
Function ip2long( ip_address )

    ip2long = false

    If preg_match("/^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$/",ip_address,"","","") Then
        Dim parts
        parts = Split(ip_address,".")
        ip2long = parts(0) * pow(256,3) + _
                  parts(1) * pow(256,2) + _
                  parts(2) * pow(256,1) + _
                  parts(3) * pow(256,0)
    End If

End Function