'=======================================================================
'�L�[����ɂ��Ĕz��̋��ʍ����v�Z����
'=======================================================================
'�y�����z
'  mAry1    = array  �l�𒲂ׂ���ƂƂȂ�z��B
'  mAry2    = array  �l���r����ΏۂƂȂ�z��B
'�y�߂�l�z
'  mAry1 �̒l�̂����A ���ׂĂ̈����ɑ��݂���L�[�̂��̂��܂ޘA�z�z���Ԃ��܂��B
'�y�����z
'  �E���̑S�Ă̈����ɑ��݂��� mAry1 �̒l��S�ėL����z���Ԃ��܂��B
'=======================================================================
Function array_intersect_key(mAry1,mAry2)

    Dim result : set result = Server.CreateObject("Scripting.Dictionary")
    Dim result_keys,key
    ReDim arg_keys(1)

    arg_keys(0) = array_keys(mAry1,"",false)
    arg_keys(1) = array_keys(mAry2,"",false)
    set result_keys = array_intersect(arg_keys(0),arg_keys(1))

    For Each key In result_keys
        result.Add result_keys(key) ,mAry1(result_keys(key))
    Next
    set array_intersect_key = result

End Function