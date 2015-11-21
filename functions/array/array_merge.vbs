'=======================================================================
'�ЂƂ܂��͕����̔z����}�[�W����
'=======================================================================
'�y�����z
'  mAry1    = array  �ŏ��̔z��B
'  mAry2    = array  �ċA�I�Ƀ}�[�W���Ă����z��B
'�y�߂�l�z
'  ���ʂ̔z���Ԃ��܂��B
'�y�����z
'  �E�O�̔z��̌��ɔz���ǉ����邱�Ƃɂ��A �ЂƂ܂��͕����̔z��̗v�f���}�[�W���A����ꂽ�z���Ԃ��܂��B
'  �E���͔z�񂪓����L�[�������L���Ă����ꍇ�A���̃L�[�Ɋւ����Ɏw�肳�ꂽ�l���A �O�̒l���㏑�����܂��B
'  �E�������A�z�񂪓����Y���ԍ���L���Ă��Ă� �l�͒ǋL����邽�߁A���̂悤�Ȃ��Ƃ͋N���܂���B
'  �E�z�񂪈�����w�肳��A���̔z�񂪐����œY���w�肳��Ă����ꍇ�A �L�[�̓Y�����A���ƂȂ�悤�ɐU�蒼����܂��B 
'=======================================================================
Function array_merge(mAry1,mAry2)

    Dim j,k
    Dim ret,retAry

    If isArray(mAry1) AND isArray(mAry2) Then

        If is_empty(mAry1) Then
            array_merge = mAry2
            Exit Function
        End If

        If is_empty(mAry2) Then
            array_merge = mAry1
            Exit Function
        End If

        Dim cnt : cnt = 0
        Dim uBoundCnt : uBoundCnt = count(mAry1,0) + count(mAry2,0)
        ReDim retAry(uBoundCnt-1)

        For Each j In mAry1
            If isObject(j) Then
                set retAry(cnt) = j
            Else
                retAry(cnt) = j
            End If
            cnt = cnt + 1
        Next

        For Each j In mAry2
            If isObject(j) Then
                set retAry(cnt) = j
            Else
                retAry(cnt) = j
            End If
            cnt = cnt + 1
        Next

        array_merge = retAry

    Else
        If Not isObject(retAry) Then
            set retAry = Server.CreateObject("Scripting.Dictionary")
        End If
    
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

                    set retAry(j) = Server.CreateObject("Scripting.Dictionary")
                    set ret = array_merge(retAry(j),mAry2(j))

                    For Each k in ret
                        if retAry(j).Exists(k) then
                            retAry(j).Item(k) = ret(k)
                        Else
                            retAry(j).Add k, ret(k)
                        End If
                    Next
                Elseif retAry.Exists(j) then
                    retAry.Item(j) = mAry2(j)
                Else
                    retAry.Add j, mAry2(j)
                End If
            Next
    
        ElseIf isArray( mAry2 ) Then
            For j = 0 to uBound( mAry2 )
                if retAry.Exists(j) then
                    retAry.Item(j) = mAry2(j)
                Else
                    retAry.Add j, mAry2(j)
                End If
            Next
        End If
    
        set array_merge = retAry
    End If

End Function