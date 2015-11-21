'=======================================================================
'��ȏ�̔z����ċA�I�Ƀ}�[�W����
'=======================================================================
'�y�����z
'  mAry1    = array  �}�[�W������Ƃ̔z��B
'  mAry2    = array  �ċA�I�Ƀ}�[�W���Ă����z��B
'�y�߂�l�z
'  ���ׂĂ̈����̔z����}�[�W�������ʂ̔z���Ԃ��܂��B
'�y�����z
'  �E ��ȏ�̔z��̗v�f���}�[�W���A �O�̔z��̍Ō�ɂ�����̔z��̒l��ǉ����܂��B 
'  �E �}�[�W������̔z���Ԃ��܂��B
'  �E ���͔z�񂪓���������̃L�[��L���Ă���ꍇ�A �����̃L�[�̒l�͔z��Ɉ�̃}�[�W����܂��B
'  �E ����͍ċA�I�ɍs���܂��B 
'  �E ���̂��߁A�l�̈���z�񎩑̂��w���Ă���ꍇ�A ���̊֐��͕ʂ̔z��̑Ή�����G���g�����}�[�W���܂��B 
'  �E �������A�z�񂪓������l�L�[��L���Ă���ꍇ�A ��̒l�͌��̒l���㏑�����A�ǉ�����܂��B 
'=======================================================================
Function array_merge_recursive(mAry1,mAry2)

    Dim j
    Dim retAry : set retAry = Server.CreateObject("Scripting.Dictionary")

    If isObject( mAry1 ) Then
        For Each j In mAry1
            if Not retAry.Exists(j) then retAry.Add j, mAry1(j)
        Next

    ElseIf isArray( mAry1 ) Then
        For j = 0 to uBound( mAry1 )
            retAry.Add j, mAry1(j)
        Next

    End If

    If isObject( mAry2 ) Then
        For Each j In mAry2
            If isObject( mAry2(j) ) Then

                set retAry(j) = array_merge_recursive(retAry(j),mAry2(j))

            Elseif retAry.Exists(j) then
                retAry.Item(j) = array(retAry.Item(j) , mAry2(j))
            Else
                retAry.Add j, mAry2(j)
            End If
        Next

    ElseIf isArray( mAry2 ) Then
        For j = 0 to uBound( mAry2 )
            if retAry.Exists(j) then
                retAry.Item(j) = array(retAry.Item(j) , mAry2(j))
            Else
                retAry.Add j, mAry2(j)
            End If
        Next
    End If

    set array_merge_recursive = retAry

End Function