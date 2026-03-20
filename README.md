#  Workflow de Certificados de Aprovação (CA)

> **Protheus TOTVS** · AdvPL · Versão 1.1  
> Autor: ERP-Tools · Atualizado: 2026

---

## Índice

- [Visão Geral](#visão-geral)
- [Pré-requisitos](#pré-requisitos)
- [Configuração SMTP](#configuração-smtp)
- [Estrutura do Código](#estrutura-do-código)
- [Fluxo Principal](#fluxo-principal)
- [Fluxo de Envio de E-mail](#fluxo-de-envio-de-e-mail)
- [Funções](#funções)
- [Query SQL](#query-sql)
- [Montagem do HTML](#montagem-do-html)
- [Erros Comuns](#erros-comuns)

---

## Visão Geral

A função `VO028` é um **workflow automático** que varre o cadastro de produtos (`SB1`) em busca de itens com **Certificado de Aprovação (CA)** vinculado, verifica o status de vencimento de cada CA e **envia um e-mail HTML** para o responsável cadastrado na tabela `SZB` (Cadastro de Emails).

Cada responsável recebe **um único e-mail consolidado** com todos os seus produtos listados em uma tabela HTML formatada, incluindo a coluna de **dias restantes ou vencidos**.

O envio é disparado apenas quando faltam **30, 15, 7 ou 1 dias** para o vencimento do CA.

```mermaid
graph LR
    A[("SB1 Produtos")] --> B["Filtra CAs vinculados"]
    B --> C["Cruza com SZB (Emails)"]
    C --> D["Calcula dias para vencer"]
    D --> E{"30, 15, 7 ou 1 dias?"}
    E -->|"Sim"| F["Envia e-mail HTML por responsável"]
    E -->|"Não"| G["Ignora registro"]
    F --> H(["✅ Concluído"])

    style A fill:#0c2c65,color:#fff
    style E fill:#f0a500,color:#fff
    style H fill:#198754,color:#fff
```

---

## Pré-requisitos

| Requisito | Detalhe |
|-----------|---------|
| **Protheus** | P12 ou superior |
| **Módulo** | Estoque/Custos (SB1 disponível) |
| **Tabela SZB** | Cadastro de e-mails preenchido (`ZB_EMAIL`) |
| **Tabela DA0** | Tabela de grupos vinculada ao `ZB_CODTAB` da SZB |
| **CA do Produto** | Campos `B1_YCA` e `B1_YDTCA` preenchidos em SB1 |
| **SMTP** | Configurado no `appserver.ini` |

---

## Configuração SMTP

No arquivo `appserver.ini` do servidor Protheus, configure a seção `[MAIL]`:

```ini
[MAIL]
User=seu_email@exemplo
Pass=sua_senha
Auth=1
TLS=1
PROTOCOL=POP3
TLSVERSION=3
SSLVERSION=3
TRYPROTOCOLS=0
AUTHLOGIN=1
AUTHPLAIN=1
AUTHNTLM=1

[SSLConfigure]
CertificateServer=totvs_certificate.crt
KeyServer=totvs_certificate_key.pem
```

> [!IMPORTANT]
>‼️O host e a porta usados em `oServer:Init()` no código devem ser idênticos ao `appserver.ini`. Divergências causam o erro *The HELLO command failed*.

---

## Estrutura do Código

```mermaid
graph LR
    A["VO028.prw"]
    A --> B["User Function VO028() Função Principal"]
    A --> C["fGetHeader() Cabeçalho HTML"]
    A --> D["fGetBody(cQry, cDias) Linha HTML por produto"]
    A --> E["fGetFooter() Rodapé HTML"]
    A --> F["fSendMail() Envio via SMTP"]

    B --> G["Monta Query SQL"]
    B --> H["Executa TcQuery"]
    B --> I["Calcula cDias por STATUS"]
    B --> J["Filtra 30/15/7/1 dias antes de enviar"]

    F --> L[Configura o envio do email via SMTP]

    style A fill:#0c2c65,color:#fff
    style B fill:#1a3a7a,color:#fff
    style C fill:#2d5fa3,color:#fff
    style D fill:#2d5fa3,color:#fff
    style E fill:#2d5fa3,color:#fff
    style F fill:#2d5fa3,color:#fff
    style L fill:#f0a503,color:#fff
```

---

## Fluxo Principal

```mermaid
flowchart TD
    START(["▶️ Início: VO028()"])
    SQL["Monta e executa Query SQL"]
    CHKEOF{{"Query retornou  dados?"}}
    MSGSTOP["MsgStop() 'Query sem dados!'"]
    LOOP(["While não Eof"])
    MAIL["cNovoMail :=AllTrim(EMAIL)"]
    CHKEMPTY{{"E-mail vazio?"}}
    CONOUT["ConOut()'Produto sem email'"]
    SKIP1["DbSkip() + Loop"]
    CALCDAYS["Calcula cDias por STATUS"]
    CHKDAYS{{"30, 15, 7 ou 1 dias?"}}
    SKIPDAYS["DbSkip() (ignora registro)"]
    CHKCHANGE{{"Mudou o responsável?"}}
    CHKPREV{{"!Empty (cMail)?"}}
    SEND1["fSendMail() responsável anterior"]
    NEWMAIL["cMail := cNovoMail cHtml := fGetHeader()"]
    BODY["cHtml += fGetBody(cQry, cDias)"]
    SKIP2["DbSkip()"]
    EOF{{"EOF?"}}
    CHKLAST{{"!Empty (cMail)?"}}
    FOOTER["cHtml += fGetFooter()"]
    SEND2["fSendMail() último responsável"]
    CLOSE["DbCloseArea()"]
    FINISH(["⏹️ Return"])

    START --> SQL --> CHKEOF
    CHKEOF -->|"Não"| MSGSTOP --> FINISH
    CHKEOF -->|"Sim"| LOOP
    LOOP --> MAIL --> CHKEMPTY
    CHKEMPTY -->|"Sim"| CONOUT --> SKIP1 --> LOOP
    CHKEMPTY -->|"Não"| CALCDAYS --> CHKDAYS
    CHKDAYS -->|"Não"| SKIPDAYS --> LOOP
    CHKDAYS -->|"Sim"| CHKCHANGE
    CHKCHANGE -->|"Sim"| CHKPREV
    CHKPREV -->|"Sim"| SEND1 --> NEWMAIL
    CHKPREV -->|"Não"| NEWMAIL
    CHKCHANGE -->|"Não"| BODY
    NEWMAIL --> BODY --> SKIP2 --> EOF
    EOF -->|"Não"| LOOP
    EOF -->|"Sim"| CHKLAST
    CHKLAST -->|"Sim"| FOOTER --> SEND2 --> CLOSE --> FINISH
    CHKLAST -->|"Não"| CLOSE

    style START fill:#198754,color:#fff
    style FINISH fill:#6c757d,color:#fff
    style MSGSTOP fill:#dc3545,color:#fff
    style SKIPDAYS fill:#f0a500,color:#fff
    style SEND1 fill:#0c2c65,color:#fff
    style SEND2 fill:#0c2c65,color:#fff
```

---

## Fluxo de Envio de E-mail

```mermaid
flowchart TD
    START(["fSendMail(cMail, cHtml)"])
    INIT["TMailManager():New() TMailMessage():New()"]
    CFG["oServer:Init() exemplo_empresa : mail.empresa.com.br:587"]
    CONNECT["oServer:SmtpConnect()"]
    CHKCONN{{"nErro != 0?"}}
    ERRMSG1["MsgStop (GetErrorString())"]
    MSG["oMessage:Clear() cFrom / cTo / cSubject cBody / MsgBodyType(html)"]
    SEND["oMessage:Send(oServer)"]
    CHKSEND{{"nErro != 0?"}}
    ERRMSG2["MsgStop (Erro ao enviar)"]
    LOG["ConOut (Enviado com sucesso)"]
    DISC["oServer:SmtpDisconnect()"]
    FINISH(["Return"])

    START --> INIT --> CFG --> CONNECT --> CHKCONN
    CHKCONN -->|"Sim"| ERRMSG1 --> FINISH
    CHKCONN -->|"Não"| MSG --> SEND --> CHKSEND
    CHKSEND -->|"Sim"| ERRMSG2 --> DISC
    CHKSEND -->|"Não"| LOG --> DISC --> FINISH

    style START fill:#0c2c65,color:#fff
    style FINISH fill:#6c757d,color:#fff
    style ERRMSG1 fill:#dc3545,color:#fff
    style ERRMSG2 fill:#dc3545,color:#fff
    style LOG fill:#198754,color:#fff
```

---

## Funções

### `User Function VO028()`

Função principal. Executa a query, calcula os dias de vencimento, filtra pelos marcos de 30/15/7/1 dias e dispara os envios agrupados por e-mail.

```advpl
User Function VO028()

Local cSQL      := ""
Local cQry      := GetNextAlias()
Local cMail     := ""   // E-mail do responsável anterior
Local cHtml     := ""   // HTML acumulado
Local cNovoMail := ""   // E-mail do registro atual
Local cDias     := ""   // Texto descritivo de dias
```

---

### `Static Function fGetHeader()`

Retorna a abertura do HTML com estilos CSS e o cabeçalho da tabela, incluindo a coluna **Dias**.

```advpl
Static Function fGetHeader()
Local cRet := ""
// ...estilos CSS...
cRet += '<th class="styleCabecalho"> ID              </th>'
cRet += '<th class="styleCabecalho"> Nome            </th>'
cRet += '<th class="styleCabecalho"> Email           </th>'
cRet += '<th class="styleCabecalho"> Codigo Tabela   </th>'
cRet += '<th class="styleCabecalho"> Descricao       </th>'
cRet += '<th class="styleCabecalho"> Data CA         </th>'
cRet += '<th class="styleCabecalho"> Status          </th>'
cRet += '<th class="styleCabecalho"> Dias            </th>'
Return cRet
```

---

### `Static Function fGetBody(cQry, cDias)`

Recebe o alias da query e o texto de dias calculado na função principal, retornando uma linha HTML com todos os dados do produto.

```advpl
Static Function fGetBody(cQry, cDias)
Local cRet := ""
cRet += '<tr>'
cRet += '<th class="styleLinha">' + AllTrim((cQry)->ID)            + '</th>'
cRet += '<th class="styleLinha">' + AllTrim((cQry)->NOME)          + '</th>'
cRet += '<th class="styleLinha">' + AllTrim((cQry)->EMAIL)         + '</th>'
cRet += '<th class="styleLinha">' + AllTrim((cQry)->CODIGO_TABELA) + '</th>'
cRet += '<th class="styleLinha">' + AllTrim((cQry)->DESCRICAO)     + '</th>'
cRet += '<th class="styleLinha">' + AllTrim((cQry)->DATA_CA)       + '</th>'
cRet += '<th class="styleLinha">' + AllTrim((cQry)->STATUS)        + '</th>'
cRet += '<th class="styleLinha">' + AllTrim(cDias)                 + '</th>'
cRet += '</tr>'
Return cRet
```

---

### `Static Function fGetFooter()`

Fecha a tabela e o HTML com uma barra de rodapé.

```advpl
Static Function fGetFooter()
Local cRet := ""
cRet += '<td class="styleRodape" colspan="13">'
cRet += 'E-mail enviado automaticamente pelo sistema Protheus - VO028'
cRet += '</td>'
// ...fecha table, body, html...
Return cRet
```

---

### `Static Function fSendMail(cMail, cHtml)`

Realiza a conexão SMTP e envia o e-mail HTML para o destinatário.

```advpl
Static Function fSendMail(cMail, cHtml)
   Local oServer  := TMailManager():New()
   Local oMessage := TMailMessage():New()
   Local nErro    := 0

   oServer:Init( "", "mail.empresa.com.br", "seu_email@exemplo", "sua_senha", 0, 587 )

   If ( nErro := oServer:SmtpConnect() ) != 0
       MsgStop("Erro SMTP: " + oServer:GetErrorString(nErro))
       Return
   EndIf

   oMessage:Clear()
   oMessage:cFrom    := "seu_email@exemplo"
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
```

---

## Query SQL

A query usa `SB1` como tabela principal, cruzando com `SZB` (cadastro de emails) e `DA0` (grupos), calculando status e dias de vencimento do CA:

```sql
SELECT
    SZB.ZB_CODIGO  AS ID,
    SZB.ZB_NOME    AS NOME,
    SZB.ZB_EMAIL   AS EMAIL,
    SZB.ZB_CODTAB  AS CODIGO_TABELA,
    SZB.ZB_DESCTAB AS DESCRICAO,
    SB1.B1_YCA     AS CA,
    CONVERT(VARCHAR(10), CAST(SB1.B1_YDTCA AS DATETIME), 103) AS DATA_CA,
    SB1.B1_YCODFAB AS COD_FABRICANTE,
    CASE
        WHEN DATEDIFF(day, SB1.B1_YDTCA, GETDATE()) > 0 THEN 'Vencido'
        WHEN DATEDIFF(day, SB1.B1_YDTCA, GETDATE()) < 0 THEN 'A Vencer'
        ELSE 'Vence Hoje'
    END AS STATUS,
    ABS(DATEDIFF(day, SB1.B1_YDTCA, GETDATE())) AS DIAS
FROM SB1990 SB1
INNER JOIN SZB990 SZB
    ON  SZB.ZB_CODIGO = SB1.B1_COD
    AND SZB.D_E_L_E_T_ = ''
    AND SZB.ZB_EMAIL  <> ''
INNER JOIN DA0990 DA0
    ON  DA0.DA0_CODTAB = SZB.ZB_CODTAB
    AND DA0.D_E_L_E_T_ = ''
WHERE SB1.D_E_L_E_T_ = ''
  AND SB1.B1_YCA     <> ''
  AND SB1.B1_YCA      > '0'
```

### Relacionamento entre tabelas

```mermaid
erDiagram
    SB1 {
        string B1_COD PK
        string B1_YCA
        date   B1_YDTCA
        string B1_YCODFAB
    }
    SZB {
        string ZB_CODIGO FK
        string ZB_NOME
        string ZB_EMAIL
        string ZB_CODTAB FK
        string ZB_DESCTAB
    }
    DA0 {
        string DA0_CODTAB PK
    }

    SB1 ||--o{ SZB : "B1_COD = ZB_CODIGO"
    SZB }o--|| DA0 : "ZB_CODTAB = DA0_CODTAB"
```

### Lógica de Status e Dias

| Condição `DATEDIFF` | STATUS | Coluna DIAS |
|--------------------|--------|-------------|
| `> 0` (data CA < hoje) | 🔴 **Vencido** | `Vencido há X dias ` |
| `< 0` (data CA > hoje) | 🟡 **Preste a Vencer** | `Vence em X dias` |
| `= 0` (data CA = hoje) | 🟠 **Vence Hoje** | `Vence hoje` |

### Regra de disparo do e-mail

O e-mail **só é enviado** quando o STATUS for `A Vencer` e a quantidade de dias for exatamente:

| Dias restantes | Envia? |
|---------------|--------|
| 30 dias | ✅ Sim |
| 15 dias | ✅ Sim |
| 7 dias  | ✅ Sim |
| 1 dias  | ✅ Sim |
| Qualquer outro valor | ❌ Não |

```advpl
If AllTrim((cQry)->STATUS) == "A Vencer" .And. ;
((cQry)->DIAS = 30 .OR. (cQry)->DIAS = 15 .OR. (cQry)->DIAS = 7 .OR. (cQry)->DIAS = 1)
    // monta e envia o e-mail
EndIf
```

---

## Montagem do HTML

O e-mail é construído em três partes que formam uma tabela HTML:

```mermaid
graph LR
    H["fGetHeader() ─────────── <html> + <style> <table> <tr> cabeçalho </tr>"]
    B["fGetBody(cQry, cDias) × N ─────────────── <tr>   <td> dados + dias</td> </tr>"]
    F["fGetFooter() ─────────── <tr> rodapé </table> </html>"]

    H --> B --> F

    style H fill:#0c2c65,color:#fff
    style B fill:#2d5fa3,color:#fff
    style F fill:#0c2c65,color:#fff
```

### Estilos CSS do e-mail

| Classe | Aplicação | Cor de fundo |
|--------|-----------|-------------|
| `.styleCabecalho` | Cabeçalho da tabela | `#0c2c65` (azul escuro) |
| `.styleLinha` | Linhas de dados | `#f6f6f6` (cinza claro) |
| `.styleRodape` | Rodapé | `#0c2c65` (azul escuro) |
| `#status` | Coluna Dias (destaque) | `#9c1717` (vermelho) |

---

## Erros Comuns

### `The HELLO command failed`

```mermaid
flowchart LR
    ERR["The HELLO command failed"]
    C1["Host/porta divergem do appserver.ini"]
    C2["Porta 465 sem flag SSL"]
    C3["DLL de SSL ausente"]
    S1["Alinhar oServer:Init() com o .ini"]
    S2["Usar porta 587 ou adicionar .T."]
    S3["Verificar tlppcore.tlpp e libssl na pasta"]

    ERR --> C1 --> S1
    ERR --> C2 --> S2
    ERR --> C3 --> S3

    style ERR fill:#dc3545,color:#fff
    style S1 fill:#198754,color:#fff
    style S2 fill:#198754,color:#fff
    style S3 fill:#198754,color:#fff
```

### `Fail to get 'TlppData'`

Adicionar no `appserver.ini`:

```ini
[TLS]
TlsEnable=1
TlppData=C:\TOTVS\Protheus\bin\AppServer\tlpp\tlppcore.tlpp
```

### Último responsável não recebe e-mail

Garantido pelo bloco após o loop — **não remover**:

```advpl
// Este bloco envia o último responsável que não é disparado dentro do While
If !Empty(cMail)
    cHtml += fGetFooter()
    fSendMail(cMail, cHtml)
EndIf
```

### Campo DIAS vazio no e-mail

Verificar se o `cDias` está sendo passado corretamente na chamada da função:

```advpl
// ❌ Errado — cDias não chega na função
cHtml += fGetBody(cQry)

// ✅ Correto
cHtml += fGetBody(cQry, cDias)
```

---

## Observações Finais

- Os campos `B1_YCA` e `B1_YDTCA` são **campos customizados** (prefixo `Y`), específicos do dicionário desta empresa.
- A tabela `SZB` é um cadastro customizado de e-mails — verifique os campos no SX3 do seu ambiente.
- O filtro de 30/15/7/1 dias pode ser ajustado conforme a necessidade do negócio.
- O `ABS()` no `DATEDIFF` garante que a coluna `DIAS` sempre exiba valor positivo, independente do status.

---

*Documentação gerada · VO028 v1.1 · ERP-Tools · 2026*
