'=======================================================================
'  URL�G���R�[�h���ꂽ��������f�R�[�h����
'=======================================================================
'�y�����z
'  sText   = string �f�R�[�h���镶����B
'�y�߂�l�z
'  �f�R�[�h�����������Ԃ��܂��B
'�y�����z
'  �E�^����ꂽ�����񒆂̂�����G���R�[�h���� %## ���f�R�[�h���܂��B
'=======================================================================
function urldecode(sText)

    Dim obj
    Dim strDecode
    Dim strOutput

    set obj=server.createobject("basp21")
    strDecode = obj.Base64(sText,5)
    strOutput = Server.HTMLEncode(strDecode)
    set obj=Nothing

    URLDecode = strOutput

    'BASP21���g�p�ł��Ȃ��ꍇ
    '-----------------------
'    On Error Resume Next
'    sTmp=""
'    iCount = 1
'    lSrcLen=Len(Source)
'    Do Until iCount > lSrcLen
'        sChr = Mid(Source,iCount,1)
'        iCount = iCount+1
'        If sChr="+" Then
'            sChr = " "
'        ElseIf sChr="%" Then
'            sHex = Mid(Source,iCount,2)
'            iCount = iCount + 2
'            iAsc = CByte("&H" & sHex)
'            If (&H00 <= iAsc And iAsc <= &H80) Or _
'               (&HA0 <= iAsc And iAsc <= &HDF) Then
'                '1�o�C�g����
'                sChr=Chr(iAsc)
'            ElseIf (&H81 <= iAsc And iAsc <= &H9F) Or _
'               (&HE0 <= iAsc And iAsc <= &HFF) Then
'                '2�o�C�g����
'                sChr = Mid(Source,iCount,1)
'                iCount = iCount + 1
'                If sChr="%" Then
'                    sHex2 = Mid(Source,iCount,2)
'                    iCount = iCount + 2
'                Else
'                    sHex2 = Hex(Asc(sChr))
'                    If Len(sHex2) = 1 Then
'                        sHex2 = "0" & sHex2
'                    End If
'                End If
'                sChr=Chr(CInt("&H" & sHex & sHex2))
'            End If
'        End If
'        sTmp=sTmp & sChr
'    Loop
'    urldecode = sTmp
End function