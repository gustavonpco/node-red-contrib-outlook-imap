# node-red-contrib-outlook-imap

### Node-RED custom node para captura de e-mails do Outlook utilizando IMAP e OAuth2 token.

Este nó personalizado captura e-mails não lidos da caixa de entrada do Outlook via IMAP, utilizando um **OAuth2 token** para autenticação.

## Instalação

3. Instalação de pacote no Node-RED:

```bash
npm i node-red-contrib-outlook-imap
```

## Configuração do Nó

No editor do Node-RED, este nó aparecerá na categoria de **input**.

### Campos de Configuração

- **Name**: Um nome descritivo para o nó, usado dentro do fluxo.
- **Email**: O endereço de e-mail do Outlook que será utilizado para capturar os e-mails.
- **OAuth2 Token**: O token OAuth2 será utilizado para autenticação com o Outlook. Este campo deve ser enviado no corpo do msg.token.

## Exemplo de Uso

### Fluxo Simples

1. Arraste o nó **outlook-imap** para o seu fluxo no Node-RED.
2. Configure o e-mail da conta do Outlook e o token OAuth2.
3. Conecte um nó de **debug** ou outro nó de processamento para visualizar ou manipular os e-mails capturados.

```plaintext
[{"id":"node-id","type":"outlook-imap","z":"1d9ab690","name":"MeuNóOutlook","email":"meuemail@outlook.com","token":"seu_token_oauth2_aqui","x":260,"y":240,"wires":[["debug-node"]]}]
```

Este nó vai capturar e-mails não lidos da sua caixa de entrada e enviá-los pela saída do nó em formato JSON. A estrutura da mensagem será algo como:

```json
{
    "from": "Remetente <remetente@exemplo.com>",
    "subject": "Assunto do E-mail",
    "date": "Data do E-mail",
    "body": "Corpo do E-mail em texto"
}
```

## Dependências

- [imap-simple](https://github.com/chadxz/imap-simple) - Biblioteca para conexão IMAP simplificada.
- [mailparser](https://nodemailer.com/extras/mailparser/) - Parser para corpo de e-mails.

Certifique-se de que essas bibliotecas estão corretamente instaladas no diretório do seu nó customizado.

## Problemas Conhecidos

- Este nó foi projetado para capturar apenas e-mails não lidos (flag **UNSEEN**). Caso queira modificar os critérios de busca, você pode alterar a função de busca no arquivo `outlook-imap.js`.
- O campo **OAuth2 Token** é desabilitado na interface para evitar edições manuais. No entanto, o token pode ser atualizado dinamicamente através do backend ou de mensagens Node-RED.

## Contribuindo

Contribuições são bem-vindas! Se você encontrar um problema ou tiver uma sugestão de melhoria, sinta-se à vontade para abrir uma issue ou enviar um pull request.

## Licença

Este projeto é licenciado sob a [MIT License](LICENSE).