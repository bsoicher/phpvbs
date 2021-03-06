'=======================================================================
' ファイルをコピーする
'=======================================================================
'【引数】
'  source  = string   コピー元ファイルへのパス。
'  dest    = string   コピー先のパス。
'【戻り値】
'  成功した場合に TRUE を、失敗した場合に FALSE を返します。
'【処理】
'  ・ ファイル source  を dest  にコピーします。
'  ・ ファイルを移動したいならは、rename() 関数を使用してください。 
' File_System class(https://github.com/masayukiando/phpvbs/blob/master/functions/filesystem/FileSystem.class.vbs)内での処理
'=======================================================================
Public Function copy(source,dest)
    fso.CopyFile source,dest
End Function