'=======================================================================
' ���[���𑗐M����
'=======================================================================
'�y�����z
'  to       = string    ������
'  subject  = string    ����
'  message  = string    �{��
'  file     = string    �Y�t�t�@�C��
'�y�߂�l�z
'  ���������ꍇ�� TRUE ���A���s�����ꍇ�� FALSE ��Ԃ��܂��B
'�y�����z
'  �Eemail �𑗐M���܂��B
'=======================================================================
function mb_send_mail( strto, subject, message, from, strFile )

    Dim basp
    Dim msg

    if inStr(subject,vbCrLf) then
        Response.Write "�薼�ɉ��s���܂߂邱�Ƃ͂ł��܂���B"
        Response.End
    end if

	Set basp = Server.CreateObject("basp21")
	msg = basp.SendMail(SMTP_SERVER, strto, from, subject, message, strFile)

    mb_send_mail = msg

end function