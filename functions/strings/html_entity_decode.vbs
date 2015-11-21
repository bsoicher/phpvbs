'=======================================================================
'����� HTML �G���e�B�e�B�𕶎��ɖ߂�
'=======================================================================
'�y�����z
'  str         = string �f�R�[�h���镶����B
'  quote_style = int    �N�H�[�g�̌`���B�ȉ��̒萔�̂����ꂩ�ł��B
'�y�߂�l�z
'  �f�R�[�h���ꂽ�������Ԃ��܂��B
'�y�����z
'  �E����� HTML �G���e�B�e�B�𕶎��ɖ߂��܂��B
'=======================================================================
function htmlspecialchars_decode(str,quote_style)

    Dim I
    Dim sText

    if empty_(quote_style) then quote_style = ENT_COMPAT
    sText = str

    if quote_style <> ENT_NOQUOTES then
        sText = Replace(sText, "&quot;", Chr(34))
    end if

    sText = Replace(sText, "&lt;"  , Chr(60))
    sText = Replace(sText, "&gt;"  , Chr(62))
    sText = Replace(sText, "&amp;" , Chr(38))
    sText = Replace(sText, "&nbsp;", Chr(32))

    For I = 1 to 255
        if I = 39 then
            if quote_style <> ENT_COMPAT then
                sText = Replace(sText, "&#" & I & ";", Chr(I))
            end if
        else
            sText = Replace(sText, "&#" & I & ";", Chr(I))
        end if
    Next

    htmlspecialchars_decode = sText

end function