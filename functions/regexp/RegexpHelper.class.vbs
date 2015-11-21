Class Regexp_Helper

    Private strPattern
    Private boolIgnoreCase
    Private boolMultiLine

    'IgnoreCase = �啶������������ʂ��Ȃ��悤�ݒ肵�܂��B
    'Global     = ������S�̂���������悤�ݒ肵�܂��B
    'pattern    = ���K�\���p�^�[����ݒ肵�܂��B
    'MultiLine  = ������𕡐��s�Ƃ��Ĉ���Ȃ��B

    Private Sub Class_Initialize()
        'empty
    End Sub

    Private Sub Class_Terminate()
        'empty
    End Sub

    Private Property Let withPattern(str)
        strPattern = str
    End Property

    Public Property Get withPattern
        withPattern = strPattern
    End Property

    Private Property Let withIgnoreCase(bool)
        boolIgnoreCase = bool
    End Property

    Public Property Get withIgnoreCase
        withIgnoreCase = boolIgnoreCase
    End Property

    Private Property Let withMultiLine(bool)
        boolMultiLine = bool
    End Property

    Public Property Get withMultiLine
        withMultiLine =boolMultiLine
    End Property

    Public Function parseOption(str)

        If left(str,1) <> "/" Then Exit Function

        Dim tmp,options
        tmp = Split(str,"/")
        withPattern = tmp(1)

        If uBound( tmp ) > 2 Then
            Dim key
            For key = 2 to uBound( tmp ) -1
                withPattern = withPattern & "/" & tmp(key)
            Next
        End If

        withMultiLine = false
        withIgnoreCase = false

        options = tmp( uBound(tmp) )
        If inStr(options,"s") > 0 Then withMultiLine = true
        If inStr(options,"i") > 0 Then withIgnoreCase = true

    End Function

End Class