'=======================================================================
' �t�@�C���p�X�Ɋւ������Ԃ�
'=======================================================================
'�y�����z
'  path     = string    ���ׂ����p�X�B
'  options  = string    �ǂ̗v�f��Ԃ��̂����I�v�V�����̃p�����[�^ options  �Ŏw�肵�܂��B����� PATHINFO_DIRNAME�A PATHINFO_BASENAME�A PATHINFO_EXTENSION ����� PATHINFO_FILENAME �̑g�ݍ��킹�ƂȂ�܂��B �f�t�H���g�ł͂��ׂĂ̗v�f��Ԃ��܂��B
'�y�߂�l�z
'   �ȉ��̗v�f���܂ޘA�z�z���Ԃ��܂��B dirname (�f�B���N�g����)�Abasename (�t�@�C����) �����āA�������݂���� extension (�g���q)�B
'   options ���g�p����ƁA ���ׂĂ̗v�f��I�����Ȃ����肱�̊֐��̕Ԃ�l�͕�����ƂȂ�܂��B 
'�y�����z
'  pathinfo() �́Apath  �Ɋւ������L����A�z�z���Ԃ��܂��B
' File_System class(http://phpvbs.verygoodtown.com/)���ł̏���
'=======================================================================

Const PATHINFO_DIRNAME = 1
Const PATHINFO_BASENAME = 2
Const PATHINFO_EXTENSION = 4
Const PATHINFO_FILENAME = 3

Public Function pathinfo(ByVal path,ByVal options)

    Dim obj : set obj = Server.CreateObject("Scripting.Dictionary")
    Dim tmp


    obj("dirname") = dirname(path)
    obj("basename") = basename(path,"")
    obj("extension") = obj("basename")

    If inStr(obj("basename"),".") Then
        tmp = Split(obj("basename"),".")
        obj("extension") = tmp( uBound(tmp) )
    End if

    obj("filename") = Replace(obj("basename"),"." & obj("extension"),"")

    If len(options) > 0 Then

        If options = PATHINFO_DIRNAME Then
            pathinfo = obj("dirname")
        ElseIf options = PATHINFO_BASENAME Then
            pathinfo = obj("basename")
        ElseIf options = PATHINFO_EXTENSION Then
            pathinfo = obj("extension")
        ElseIf options = PATHINFO_FILENAME Then
            pathinfo = obj("filename")
        End if
        Exit Function
    End If

    set pathinfo = obj
End Function