#INCLUDE "TOPCONN.CH"
#include "protheus.ch"
#include "totvs.ch"
#include "parmtype.ch"

/*/{Protheus.doc} VO001
@author ERP-Tools
@since 20/07/2018
@version 1.1
@description Workflow de recebimento de materiais por produto
@type Function
/*/

User Function VO028()

Local cSQL   := ""
Local cQry   := GetNextAlias()
Local cMail  := ""
Local cHtml  := ""
Local cNovoMail := ""
Local cDias  := ""

cSql := " SELECT  "
cSql += "     SZB.ZB_CODIGO AS ID, "
cSql += "     SZB.ZB_NOME AS NOME, "
cSql += "     SZB.ZB_EMAIL AS EMAIL, "
cSql += "     SZB.ZB_CODTAB AS CODIGO_TABELA, "
cSql += "     SZB.ZB_DESCTAB AS DESCRICAO, "
cSql += "     SB1.B1_YCA AS CA, "
cSql += "     CONVERT(VARCHAR(10), CAST(SB1.B1_YDTCA AS DATETIME), 103) AS DATA_CA, "
cSql += "     SB1.B1_YCODFAB AS COD_FABRICANTE, "
cSql += "     CASE  "
cSql += "         WHEN DATEDIFF(day, SB1.B1_YDTCA, GETDATE()) > 0 THEN 'Vencido'  "
cSql += "         WHEN DATEDIFF(day, SB1.B1_YDTCA, GETDATE()) < 0 THEN 'A Vencer'  "
cSql += "         ELSE 'Vence Hoje'  "
cSql += "     END AS STATUS,  "
cSql += "     ABS(DATEDIFF(day, SB1.B1_YDTCA, GETDATE())) AS DIAS "
cSql += " FROM SB1990 SB1 "
cSql += " INNER JOIN SZB990 SZB "
cSql += "     ON SZB.ZB_CODIGO = SB1.B1_COD "
cSql += "     AND SZB.D_E_L_E_T_ = '' "
cSql += "     AND SZB.ZB_EMAIL <> '' "
cSql += " INNER JOIN DA0990 DA0 "
cSql += "     ON DA0.DA0_CODTAB = SZB.ZB_CODTAB "
cSql += "     AND DA0.D_E_L_E_T_ = '' "
cSql += " WHERE SB1.D_E_L_E_T_ = '' "
cSql += " AND SB1.B1_YCA <> '' "
cSql += " AND SB1.B1_YCA > '0' "


TcQuery(cSQL) New Alias(cQry)

If (cQry)->(Eof())
    MsgStop("Query sem dados!")
    Return
EndIf

While !(cQry)->(Eof())

    cNovoMail := AllTrim((cQry)->EMAIL)
    
    If Empty(cNovoMail)
        ConOut("Produto sem e-mail : " + AllTrim((cQry)->ID) + " - " + AllTrim((cQry)->DESCRICAO))
        (cQry)->(DbSkip())
        Loop
    EndIf

    If AllTrim((cQry)->STATUS) == "Vencido"
        cDias := "Vencido há " +AllTrim(Str((cQry)->DIAS)) + " dias "
    ElseIf AllTrim((cQry)->STATUS) == "A Vencer"
        cDias := "Vence em " + AllTrim(Str((cQry)->DIAS)) + " dias"
    Else 
        cDias := "Vence hoje"
    EndIf
  

    If AllTrim((cQry)->STATUS) == "A Vencer" .And. ((cQry)->DIAS = 30 .OR. (cQry)->DIAS = 15 .OR. (cQry)->DIAS = 7 .OR. (cQry)->DIAS = 1)
        
        If cMail != cNovoMail 
            If !Empty(cMail)
                cHtml += fGetFooter()
                fSendMail(cMail, cHtml)
            EndIf
        
            cMail := cNovoMail
            cHtml := fGetHeader()
        EndIf

        cHtml += fGetBody(cQry, cDias)
    EndIf   

    (cQry)->(DbSkip())
EndDo

If !Empty(cMail)
    cHtml += fGetFooter()
    fSendMail(cMail, cHtml)
EndIf

(cQry)->(DbCloseArea())

Return

Static Function fGetHeader()
    Local cRet := ""

	cRet := '<!DOCTYPE HTML PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"> '
	cRet += '<html xmlns="http://www.w3.org/1999/xhtml"> '
	cRet += '<head> '
	cRet += '    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" /> '
	cRet += '    <title>Workflow</title> '
	cRet += '    <style type="text/css"> '
	cRet += '        <!--  '
	cRet += '		.styleTable{ '
	cRet += '			border:0; '
	cRet += '			cellpadding:3; '
	cRet += '			cellspacing:2; '
	cRet += '			width:100%; '
	cRet += '		} '
	cRet += '		.styleTableCabecalho{ '
	cRet += '            background: #fff; '
	cRet += '            color: #ffffff; '
	cRet += '            font: 14px Arial, Helvetica, sans-serif; '
	cRet += '			font-weight: bold; '
	cRet += '		} '
	cRet += '        .styleCabecalho{ '
	cRet += '            background: #0c2c65; '
	cRet += '            color: #ffffff; '
	cRet += '            font: 12px Arial, Helvetica, sans-serif; '
	cRet += '			font-weight: bold; '
	cRet += '			padding: 5px; '
	cRet += '        } '
	cRet += '		.styleLinha{ '
	cRet += '            background: #f6f6f6; '
	cRet += '            color: #747474; '
	cRet += '            font: 11px Arial, Helvetica, sans-serif; '
	cRet += '			padding: 5px; '
	cRet += '        } '
	cRet += '        .styleRodape{ '
	cRet += '            background: #0c2c65; '
	cRet += '            color: #ffffff; '
	cRet += '            font: 12px Arial, Helvetica, sans-serif; '
	cRet += '			font-weight: bold; '
	cRet += '			text-align: center; '
	cRet += '			padding: 5px; '
	cRet += '        } '
	cRet += '		.styleLabel{ '
	cRet += '			color:#0c2c65; '
	cRet += '		} '
	cRet += '		.styleValor{ '
	cRet += '			color:#747474; '
	cRet += '		} '
    cRet += '       #status{ '
    cRet += '       background: #9c1717; '
    cRet += '       } '
	cRet += '        --> '
	cRet += '    </style> '
	cRet += '</head> '
	cRet += '<body> '
	cRet += '    <table class="styleTable" align="center">	 '
	cRet += '    <tr align=center> '
    cRet += '       <th class="styleCabecalho"> ID     </th>'
    cRet += '       <th class="styleCabecalho"> Nome   </th>'
    cRet += '       <th class="styleCabecalho"> Email  </th>'
    cRet += '       <th class="styleCabecalho"> Codigo da Tabela </th>'
    cRet += '       <th class="styleCabecalho"> Descricao </th>'
    cRet += '       <th class="styleCabecalho"> Data CA </th>'
    cRet += '       <th class="styleCabecalho" > Status</th>'
    cRet += '       <th class="styleCabecalho" id="status"> Dias </th>'
    cRet += '    </tr>'
    Return cRet

Static Function fGetBody(cQry, cDias)
    Local cRet := ""
    cRet += '<tr>'
    cRet += '<th class="styleLinha" width="60" scope="col">'   +  AllTrim((cQry)->ID)            + '</th>'
    cRet += '<th class="styleLinha" width="100" scope="col">'  +  AllTrim((cQry)->NOME)          + '</th>'
    cRet += '<th class="styleLinha" width="60" scope="col">'   +  AllTrim((cQry)->EMAIL)         + '</th>'
    cRet += '<th class="styleLinha" width="60" scope="col">'   +  AllTrim((cQry)->CODIGO_TABELA) + '</th>'
    cRet += '<th class="styleLinha" width="200" scope="col">'  +  AllTrim((cQry)->DESCRICAO)     + '</th>'
    cRet += '<th class="styleLinha" width="60" scope="col">'   +  AllTrim((cQry)->DATA_CA)       + '</th>'
    cRet += '<th class="styleLinha" width="60" scope="col">'   +  AllTrim((cQry)->STATUS)        + '</th>'
    cRet += '<th class="styleLinha" width="60" scope="col">'   +  AllTrim(cDias)                 + '</th>'
    cRet += '</tr>'
Return cRet

Static Function fGetFooter()
    Local cRet := ""

	cRet += '        </tr> '
	cRet += '        <tr> '
	cRet += '            <td class="styleRodape" width="60" scope="col" colspan="13"> '
	cRet += '                E-mail enviado automaticamente pelo sistema Protheus - VO001 '
	cRet += '            </td> '
	cRet += '        </tr> '
	cRet += '	</table> '
	cRet += '</body> '
	cRet += '</html> '

Return cRet

Static Function fSendMail(cMail, cHtml)

   Local oServer  := TMailManager():New()
   Local oMessage := TMailMessage():New()
   Local nErro    := 0

   oServer:Init( "", "mail.alphapcm.com.br", "maria.costa@alphapcm.com.br", "Alpha@2026", 0, 587 )

   If ( nErro := oServer:SmtpConnect() ) != 0
       MsgStop("Erro SMTP: " + oServer:GetErrorString(nErro))
       Return
   EndIf

   oMessage:Clear()
   oMessage:cFrom    := "maria.costa@alphapcm.com.br"
   oMessage:cTo      := cMail
   oMessage:cSubject := "Recebimento de Material"
   oMessage:cBody    := cHtml
   oMessage:MsgBodyType("text/html")


   nErro := oMessage:Send(oServer)
   If nErro != 0
       MsgStop("Erro ao enviar para: " + cMail + " - " + oServer:GetErrorString(nErro))
   Else
       ConOut("Enviado com sucesso para: " + cMail)
   EndIf

   oServer:SmtpDisconnect()
Return
