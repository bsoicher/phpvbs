'=======================================================================
' ������� soundex �L�[���v�Z����
'=======================================================================
'�y�����z
'  format   = string ���͕�����B
'�y�߂�l�z
'  ���^�������N�H�[�g�����������Ԃ��܂��B
'�y�����z
'  �E str  �� soundex �L�[���v�Z���܂��B
'  �E soundex �L�[�ɂ́A�����悤�Ȕ����̒P��Ɋւ��ē��� soundex �L�[�����������Ƃ�������������܂��B ���̂��߁A�����͒m���Ă��邪�A�X�y�����킩��Ȃ��ꍇ�ɁA �f�[�^�x�[�X���������邱�Ƃ�e�Ղɂ��邱�Ƃ��ł��܂��B
'  �E soundex �֐��́A���镶������n�܂� 4 �����̕������Ԃ��܂��B
'  �E ���� soundex �֐��ɂ��Ă̐����́ADonald Knuth �� "The Art Of Computer Programming, vol. 3: Sorting And Searching", Addison-Wesley (1973), pp. 391-392 �ɂ���܂��B 
'=======================================================================
Function soundex(str)

    Dim i,j, l, r, p, m, s
    p = [?](Not isNumeric(p) , 4 , [?](p > 10 , 10 , [?](p < 4 , 4 , p) ) )

    set m = Server.CreateObject("Scripting.Dictionary")
    m.Add "BFPV", 1
    m.Add "CGJKQSXZ", 2
    m.add "DT", 3
    m.add "L", 4
    m.add "MN", 5
    m.add "R", 6

    s = Ucase( str )
    s = preg_replace("/[^A-Z]/","",s,"","")
    s = str_split(s,1)
    r = array( array_shift(s) )

    For i = 0 to uBound(s)
        For Each j In m
            if inStr(j,s(i)) and r( uBound(r) ) <> m.Item(j) Then
                array_push r,m(j)
                Exit For
            End If
        Next
    Next

    If uBound(r) + 1 > p Then
        r = array_slice(r,0,p-1)
    End If

    Dim newArray()
    ReDim newArray(p - (uBound(r)+1))

    soundex = join(r,"") & join( newArray, "0" )

End Function