'=======================================================================
' ��̕�����̃��[�x���V���^�C���������v�Z����
'=======================================================================
'�y�����z
'  str1 = string    ���[�x���V���^�C���������v�Z���镶����̂ЂƂB
'  str2 = string    ���[�x���V���^�C���������v�Z���镶����̂ЂƂB
'�y�߂�l�z
'  ���̊֐��́A�����Ŏw�肵����̕�����̃��[�x���V���^�C��������Ԃ��܂��B
'  ����������̈�� 255 �����̐�����蒷���ꍇ�� -1 ��Ԃ��܂��B
'�y�����z
'  �E���[�x���V���^�C�������́Astr1  �� str2  �ɕϊ����邽�߂ɒu���A�}���A�폜 ���Ȃ���΂Ȃ�Ȃ��ŏ��̕������Ƃ��Ē�`����܂��B
'  �E�A���S���Y���̕��G���́A O(m*n) �ł��B
'  �E�����ŁAn ����� m �͂��ꂼ�� str1  ����� str2  �̒����ł� (O(max(n,m)**3) �ƂȂ� similar_text() ���͗ǂ��ł����A �܂����Ȃ�̌v�Z�ʂł�)�B
'  �E��L�̍ł��ȒP�Ȍ`���ł́A���̊֐��̓p�����[�^�Ƃ��Ĉ����������Ƃ�A str1  ���� str2  �ɕϊ�����ۂɕK�v�� �}���A�u���A�폜���Z�̐��݂̂��v�Z���܂��B
'=======================================================================
Function levenshtein( str1, str2 )

    Dim s,l,t,i,j,m,n,u
    Dim a,tmp

    s = str_split(str1,1)
    u = str_split(str2,1)

    If isArray(s) Then l = count(s,"") Else l = 0
    If isArray(u) Then t = count(u,"") Else t = 0

    If is_empty(l) or is_empty(t) Then
        If [>](l , t) Then levenshtein = l
        If [>](t , l) Then levenshtein = t
        If isEmpty(levenshtein) Then levenshtein = 0
        Exit Function
    End If

    ReDim a(l)
    For i = 0 to l
        toReDim a(i),t
    Next

    For i = l to 0 Step -1
       a(i)(0) = i
    Next

    For i = t to 0 Step -1
       a(0)(i) = i
    Next

    i = 0
    m = l

    Do While(i < m)

        j = 0
        n = t

        Do While(j < n)
            tmp = a(i)(j + 1) + 1
            If tmp > a(i+1)(j) + 1 Then tmp = a(i+1)(j) + 1
            If tmp > a(i)(j) + intval([!=](s(i) ,u(j))) Then tmp = a(i)(j) + intval([!=](s(i) ,u(j)))
            a(i+1)(j+1) = tmp

            j = j + 1
        Loop

        i = i +1
    Loop

    levenshtein = a(l)(t)

End Function