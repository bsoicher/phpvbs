'=======================================================================
'��ȏ�̗v�f��z��̍ŏ��ɉ�����
'=======================================================================
'�y�����z
'  mAry     = Array �z��
'  mVal     = mixed �ǉ�����v�f
'�y�߂�l�z
'  �E������� mAry  �̗v�f�̐���Ԃ��܂��B
'�y�����z
'  �E���X�g�̗v�f�͑S�̂Ƃ��ĉ������邽�߁A ������ꂽ�v�f�̏��Ԃ͕ς��Ȃ����Ƃɒ��ӂ��Ă��������B 
'  �E�z��̐��l�Y���͂��ׂĐV���Ƀ[������U��Ȃ�����܂��B 
'  �E���e�̃L�[�ɂ��Ă͕ύX����܂���B
'=======================================================================
Function array_unshift(ByRef mAry, ByVal mVal)

    Dim intCounter
    Dim intElementCount

    If IsArray(mAry) Then
        If IsArray(mVal) Then

            ret = array_push(mVal,mAry)
            mAry = mVal

        Else

            ReDim Preserve mAry(UBound(mAry) + 1)

            For intCounter = UBound(mAry) to 1 Step -1
                mAry(intCounter) = mAry(intCounter -1)
            Next

            mAry(0) = mVal

        End If
    Else
        If IsArray(mVal) Then
            mAry = mVal
        Else
            mAry = Array(mVal)
        End If
    End If

    array_unshift = UBound(mAry)

End Function