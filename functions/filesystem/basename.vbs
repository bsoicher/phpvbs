'=======================================================================
' �p�X���̃t�@�C�����̕�����Ԃ�
'=======================================================================
'�y�����z
'  path      = string   �p�X�B
'  suffix    = string   �t�@�C�������A suffix  �ŏI������ꍇ�A ���̕������J�b�g����܂��B
'�y�߂�l�z
'  �w�肵�� path  �̃x�[�X����Ԃ��܂��B
'�y�����z
'  �E���̊֐��́A�t�@�C���ւ̃p�X��L���镶����������Ƃ��A �t�@�C���̃x�[�X����Ԃ��܂��B
'=======================================================================
Function basename(path, suffix)

    Dim b
    b = preg_replace("/^.*[��/����]/g","",path,null,null)

    If len(suffix) > 0 Then
        If Right(b,len(suffix)) = suffix Then
            b = Left(b,len(b) - len(suffix))
        End If
    End If

    basename = b

End Function