# Relatórios do Centro de Operações

Automação Node.js responsável por consolidar dados operacionais diariamente, semanalmente e mensalmente, gerar planilhas em Excel, enviar resumos por e-mail e notificar integrações via webhook para o Centro de Operações da Rede Avantar.

## Visão Geral
- Coleta dados no SGCOR (produção, transmissões, apólices e sinistros) autenticando com credenciais de serviço.
- Consulta assistências urgentes a partir da API do SULTS.
- Gera planilhas Excel com os registros detalhados do período processado.
- Dispara e-mails transacionais via SMTP DreamHost com resumos e anexos.
- Envia notificações JSON para um webhook corporativo.
- Executa automaticamente rotinas diárias, semanais e mensais com `node-cron`.
- Emite alertas por e-mail em caso de falha ou dados inconsistentes.

## Estrutura do Projeto
- `relatorio.js`: ponto de entrada da aplicação e orquestrador das rotinas (diária, semanal, mensal), incluindo agendamento com `node-cron`.
- `getAuth.js`: autenticação no SGCOR e definição da URL base da API.
- `utils/formatDate.js`: utilitário para trabalhar com datas relativas.

Dependências principais: `node-cron`, `nodemailer`, `exceljs`, `dotenv`, `pm2`.

## Pré-requisitos
- Node.js 18 LTS ou superior.
- Acesso às APIs SGCOR e SULTS com credenciais válidas.
- Conta SMTP DreamHost (ou compatível) para envio dos relatórios.
- Servidor Linux (a automação oficial roda em uma VPS Hostinger usando PM2 e cron).

## Variáveis de Ambiente
Crie um arquivo `.env` na raiz contendo:

```
SGCOR_USERNAME=usuario.suporte@dominio
SGCOR_PASSWORD=senhaSGCOR
SGCOR_API_URL=https://apirest.gruposgcor.com.br/api

MAIL_EMAIL=suporte@avantar.com
MAIL_PASSWORD=senhaSMTP
DIRETOR_EMAIL=diretor@avantar.com
DIRETOR_NOME=Nome Diretor
DIRETOR_TELEFONE=559999999999
EMAIL_ADM=tecnologia@avantar.com

SULTS_ACCESS_TOKEN=token_sults
WEBHOOK_URL=https://exemplo.com/webhook
```

> Todos os campos são obrigatórios, exceto `WEBHOOK_URL` (o envio é pulado se não informado) e `DIRETOR_*` (usados para personalização de mensagens).

## Instalação Local
1. Instale dependências: `npm install`.
2. Configure o `.env` conforme descrito acima.
3. Execute para teste manual: `node relatorio.js`. O processo permanecerá ativo executando os agendamentos configurados; use `Ctrl+C` para encerrar.

## Operação na VPS Hostinger
- O código está implantado em uma VPS Hostinger (Ubuntu) dentro do diretório `~/relatorios_do_centro_de_operacoes`.
- O processo é gerenciado pelo `pm2`, garantindo autorestart e logs centralizados.
- Cron jobs do servidor garantem que o PM2 suba após reinicializações e executam health checks programados.

### Fluxo de Deploy
1. Acesse a VPS via SSH.
2. Atualize o repositório: `git pull`.
3. Instale dependências (se necessário): `npm install`.
4. Reinicie a aplicação com PM2:
   ```
   pm2 restart relatorios-operacoes
   pm2 save
   ```

### Provisionamento PM2
Inicialização (executada uma única vez na VPS):
```
pm2 start relatorio.js --name relatorios-operacoes
pm2 startup                # gera comando para o serviço do sistema
sudo env PATH=$PATH pm2 startup systemd -u $USER --hp $HOME
pm2 save
```

### Cron Jobs do Servidor
No Hostinger, os seguintes jobs estão configurados no `crontab`:
```
@reboot /usr/bin/pm2 resurrect
0 5 * * * /usr/bin/pm2 restart relatorios-operacoes >/dev/null 2>&1
```

> As rotinas internas (diária, semanal e mensal) são disparadas automaticamente pelo `node-cron` dentro da aplicação (`relatorio.js`). O cron do sistema é usado apenas para garantir a disponibilidade do processo controlado pelo PM2.

### Logs e Monitoramento
- `pm2 logs relatorios-operacoes` para acompanhar a execução em tempo real.
- `pm2 status` exibe métricas de uptime e reinicializações.
- E-mails de alerta são enviados para `EMAIL_ADM` quando uma exceção não tratada ocorre.

## Rotinas Automatizadas
- **Diária**: Executa de terça a sábado às 06h00 (horário `America/Sao_Paulo`). Gera relatório do dia útil anterior.
- **Semanal**: Executa aos sábados às 06h15, cobrindo segunda a sexta da semana anterior.
- **Mensal**: Executa no dia 1º de cada mês às 06h00, consolidando todo o mês anterior.

Cada rotina:
1. Autentica no SGCOR e coleta dados paginados.
2. Gera planilha Excel com a produção detalhada.
3. Cria resumo numérico e envia por e-mail.
4. Envia o mesmo resumo para o webhook configurado.
5. Em caso de erro, dispara e-mail de alerta para a equipe.

## Execução Manual
Para rodar uma rotina específica durante testes, descomente o bloco de execução imediata ao final de `relatorio.js` e execute `node relatorio.js`. Relembre de comentar novamente após o teste.

## Solução de Problemas
- **Falha no envio de e-mail**: verifique credenciais SMTP e se a porta 587 está liberada na VPS.
- **Timeout nas APIs**: confirme conectividade com os endpoints SGCOR e SULTS e revise o token/credenciais.
- **Processo não inicia após reboot**: confirme os cron jobs `@reboot`/`pm2 resurrect` e se `pm2 save` foi executado após a última alteração.
- **Horários incorretos**: a aplicação usa `America/Sao_Paulo`; confira o timezone do sistema operacional da VPS.

## Contribuições
1. Crie um branch com sua alteração.
2. Atualize o README sempre que novas integrações ou variáveis de ambiente forem incluídas.
3. Abra PR ou mergeie conforme o fluxo interno da equipe.

---
Em caso de dúvidas operacionais, contate o time de Tecnologia Avantar.