'=======================================================================
'  MIME base64 �����Ńf�[�^���G���R�[�h����
'=======================================================================
'�y�����z
'  data = mixed  �G���R�[�h����f�[�^�B
'�y�߂�l�z
'  �G���R�[�h���ꂽ�f�[�^�𕶎���ŕԂ��܂��B
'�y�����z
'  �E�w�肵�� data  �� base64 �ŃG���R�[�h���܂��B
'=======================================================================
Function base64_encode(data)

    Dim obj
    set obj=server.createobject("basp21")
    base64_encode = obj.Base64(data,0)
    set obj = nothing

    'basp21���g�p���Ȃ��ꍇ
'    Dim ST, DM, EL, bin
'  
'    Set ST = CreateObject("ADODB.Stream")
'    ST.Type = adTypeText
'    ST.Charset = "Shift-JIS"
'    ST.Open
'    ST.WriteText PlainText
'    ST.Position = 0
'    ST.Type = adTypeBinary
'    bin = ST.Read
'    ST.Close
' 
'    Set DM = CreateObject("Microsoft.XMLDOM")
'    Set EL = DM.CreateElement("tmp")
'    EL.DataType = "bin.base64"
'    EL.NodeTypedValue = bin
'    base64_encode = EL.Text

End Function