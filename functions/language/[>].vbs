'=======================================================================
'a ‚ª b ‚æ‚è‘½‚¢ TRUEB
'=======================================================================
'yˆø”z
'  a    = mixed  ’l
'  b    = mixed  ”äŠr‚·‚é’l
'y–ß‚è’lz
'  a‚ªb‚æ‚è‘½‚¢ê‡‚ÉTRUE ‚ðA‘½‚­‚È‚¢ê‡‚É FALSE ‚ð•Ô‚µ‚Ü‚·B
'yˆ—z
'  E¶•Ó‚Æ‰E•Ó‚ð”äŠr‚µ‚Ü‚·B
'=======================================================================
Function [>](a, b)

    [>] = false

    Dim tmp_a,tmp_b
    Dim key

    If (isArray(a) or isArray(b)) or (isObject(a) or isObject(b)) Then

        If isObject(a) and isObject(b) Then
            If a.count < b.count Then Exit Function

            tmp_a = a.keys : tmp_b = b.keys
            If Not [>](uBound(tmp_a),uBound(tmp_b)) Then Exit Function
            [>] = true
        End If

        If isObject(a) and not isObject(b) Then
            [>] = true
        End If

        If isArray(a) and isArray(b) Then
            If uBound(a) < uBound(b) Then Exit Function

            For key = 0 to uBound(a)
                If Not [>](a(key),b(key) ) Then Exit Function
            Next

            [>] = true
        End If

        If isArray(a) and not isArray(b) Then
            [>] = true
        End If
    Else
        [>] = eval(a > b)
    End If

End Function