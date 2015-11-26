'=======================================================================
' ファイルポインタから行を取得し、CSVフィールドを処理する
'=======================================================================
'【引数】
'  
'【戻り値】
'  読み込んだフィールドの内容を含む数値添字配列を返します。
'【処理】
'  fgets() に動作は似ていますが、 fgetcsv() は行を CSV  フォーマットのフィールドとして読込み処理を行い、 読み込んだフィールドを含む配列を返すという違いがあります。
' File_System class(https://github.com/masayukiando/phpvbs/blob/master/functions/filesystem/FileSystem.class.vbs)内での処理
'=======================================================================
Public Function fgetcsv(delim)

    Dim tmp,d
    If len(delim) > 0 Then d = delim Else d = ","
    tmp = ts.ReadLine
    fgetcsv = fgetcsv_helper(tmp,d)

End Function


'************************************
Public Function fgetcsv_helper(str,d)

    Dim matchAll,key
    Dim data,field,record : field = 0 : record = 0
    ReDim data(0)

    If preg_match_all(_
    "/" & d & "|" & vbCr & "?" & vbLf & "|[^" & d & """" & vbCrLf & "][^" & d & "" & vbCrLf & "]*|""(?:[^""]|"""")*""/",_
    str, matchAll,PREG_PATTERN_ORDER,"") Then
        For Each key In matchAll(0)
            Select Case key
            Case d
                field = field + 1
            Case vbCrLf
                [] data , ""
                record = record +1
            Case Else
                If left(key,1) = """" Then
                    key = Replace(substr(key,2,-1),"""""","""")
                End if
                [] data(record), key
            End Select
        Next
    End If

    fgetcsv_helper = data

End Function