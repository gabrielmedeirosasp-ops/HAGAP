# HAGAP — Controle de Obras

Projeto **independente** do site HAGAP atual.

O site existente não é alterado. Este módulo cria uma nova página focada em rastrear, por projeto:

**Programação → Execução/Documentos → Medição → FFO**

## Fontes automáticas

A rotina lê o Gmail em modo **somente leitura** e classifica:

- **DOCUMENTOS <projeto>** → documentos finais enviados / obra concluída documentalmente.
- **DOCUMENTOS PARCIAL <projeto>** → fechamento parcial.
- **PDE** → programação, projeto, município, datas e situação.
- **PLV** → programação de linha viva e projeto.
- **OMB** → OMB → PDE → projeto.
- **BMD** → medição da obra.
- **FFO** → fechamento físico/final.

Um e-mail com mais de um projeto gera vínculos para todos os projetos encontrados. O e-mail original é preservado como origem e pode ser aberto pela tela.

## Relação OMB → PDE → Projeto

OMB não precisa conter projeto.

A rotina grava o PDE citado na OMB e cruza com os eventos PDE já recebidos:

**OMB → número PDE → PDE → nº projeto**

Se o PDE ainda não estiver no banco, a OMB fica pendente e é relacionada automaticamente em uma sincronização futura.

## Deduplicação

O Gmail message_id é a chave do e-mail.

A mesma mensagem não é processada duas vezes. PDE/PLV/OMB repetidos continuam no histórico, mas a tela soma referências distintas para não duplicar a obra.

## Histórico inicial

Padrão:

**01/01/2025 → hoje**

Depois da primeira carga, a rotina reconsulta apenas uma pequena janela recente e ignora message_id já processado. Isso captura reenvios e atualizações sem reler todo o histórico a cada ciclo.

## Automação

config.json:

- sync_interval_minutes: 10
- history_start: 2025-01-01
- port: 5055

Quando app.py está em execução, o Gmail é consultado automaticamente no intervalo definido.

Para iniciar junto com o Windows:

**ativar_rotina_automatica.bat**

A página fica disponível em:

**http://127.0.0.1:5055**

## Primeira instalação

1. Execute **instalar.bat**.
2. Coloque o arquivo OAuth do Gmail com o nome **credentials.json** nesta pasta.
3. Execute **conectar_gmail.bat** uma única vez.
4. Faça o login Google e autorize **somente leitura**.
5. Execute **iniciar.bat**.
6. Quando estiver validado, execute **ativar_rotina_automatica.bat**.

credentials.json, token.json e o banco local não são enviados ao GitHub.

## Segurança e preservação

- Escopo Gmail: gmail.readonly.
- A rotina não envia, apaga, move ou marca e-mails.
- O site HAGAP atual não é modificado.
- /boletim não é modificado.
- A nova base usa SQLite independente: controle_obras.db.

## Situações exibidas

A tela mantém os conceitos separados:

- **Documentos enviados**
- **Documentos parcial**
- **Programada** (PDE / PLV / OMB)
- **BMD**
- **FFO**

Não transforma BMD parcial em FFO e não considera OMB projeto sem antes relacionar seu PDE.

## Rastreabilidade

Ao abrir uma obra, a tela mostra a linha do tempo dos eventos com:

- tipo;
- data;
- assunto;
- referência PDE/PLV/OMB;
- município quando localizado;
- anexo;
- link para abrir o e-mail original.
