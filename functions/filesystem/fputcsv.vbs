'=======================================================================
' 行を CSV 形式にフォーマットし、ファイルポインタに書き込む
'=======================================================================
'【引数】
'  fields  =  string    値の配列。
'  delimiter  =  string オプションの delimiter  はフィールド区切り文字 (一文字だけ) を指定します。デフォルトはカンマ (,) です。
'  enclosure  =  string オプションの enclosure  はフィールドを囲む文字 (一文字だけ) を指定します。デフォルトは二重引用符 (") です。
'【戻り値】
'  書き込んだ文字列の長さを返します。失敗した場合は FALSE を返します。
'【処理】
'  fputcsv() は、行（fields  配列として渡されたもの）を CSV としてフォーマットし、それを ファイルに書き込みます (いちばん最後に改行を追加します)。
' File_System class(https://github.com/masayukiando/phpvbs/blob/master/functions/filesystem/FileSystem.class.vbs)内での処理
'=======================================================================
Public function fputcsv(fields,delimiter,enclosure)

    fputcsv = false
    If len(delimiter) = 0 Then delimiter = ","
    If len(enclosure) = 0 Then enclosure = """"

    Dim key,replaced
    For key = 0 to uBound(fields)
        replaced = false
        If inStr(fields(key),delimiter) or inStr(fields(key),enclosure) or inStr(fields(key),vbCrLf) Then
            fields(key) = Replace(fields(key),enclosure,enclosure & enclosure)
            fields(key) = enclosure & fields(key) & enclosure
        End If
    Next

    Dim str : str = join(fields,delimiter)
    ts.WriteLine str
    fputcsv = len(str)
end function