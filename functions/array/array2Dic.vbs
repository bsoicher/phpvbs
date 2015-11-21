'=======================================================================
'配列をディクショナリに変換する
'=======================================================================
'【引数】
'  arr  = array  配列
'【戻り値】
'  ディクショナリオブジェクト。
'【処理】
'  ・渡された配列を ディクショナリオブジェクトに変換します。
'=======================================================================
Function array2Dic(ByVal myAry)

    Dim i,tmpObj
    set tmpObj = Server.CreateObject("Scripting.Dictionary")
    For i = 0 to uBound(myAry)
        tmpObj.add i, myAry(i)
    Next
    set array2Dic = tmpObj

End Function