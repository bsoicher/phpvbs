'=======================================================================
'文字列を ASP コードとして評価する
'=======================================================================
'【引数】
'  code_str = string   文字列
'【戻り値】
'  NULL を返します。
'【処理】
'  ・code_str  で与えられた文字列を PHP コードとして評価します。 
'  ・中でも、データベースのテキストフィールドにコードを保存し、 後で実行するためには便利です。
'=======================================================================
sub eval(code_str)
    execute code_str
end sub