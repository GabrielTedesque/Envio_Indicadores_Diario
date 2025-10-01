# Envio diário de indicadores (Power BI → PDF → E-mail)

## Introdução rápida
Este projeto é uma automação em **Python** que acessa **dashboards publicados** no Power BI Service, **exporta** cada um em **PDF**, aplica **regras de páginas**, **mescla** tudo em **um único arquivo** e **envia** automaticamente por **e-mail** com corpo em **HTML**.  
O objetivo é **eliminar tarefas manuais repetitivas**, padronizar a entrega e garantir que os **indicadores** cheguem **todos os dias**, no **mesmo horário**, em **um único e-mail**.

---

## Explicação detalhada

> **Por que Microsoft Edge?**  
> Utilizo o **Microsoft Edge** por causa do fluxo de **autenticação corporativa/SSO** do Power BI. Eu faço a autenticação **1 vez por semana** manualmente (no navegador controlado pelo Selenium). Com isso, a sessão permanece ativa e as execuções diárias acontecem **sem pedir login**, evitando falhas no momento do agendamento.

### ✨ O que esta automação resolve
- Substitui a rotina de abrir cada dashboard, exportar, salvar, juntar e enviar arquivos.
- Padroniza a entrega: **um único PDF** com apenas as páginas relevantes.
- Garante **consistência** (mesmo horário, todos os dias) e reduz erros humanos.
- Libera tempo para **análise** ao invés de tarefas operacionais.

### 🧠 Como funciona (passo a passo)
1. **Carregamento de configurações**  
   O script lê variáveis de ambiente (via `.env`, se existir) para definir pastas, horário de execução e caminho do `msedgedriver`. Se não houver `.env`, usa valores padrão embutidos no código.

2. **Inicialização do navegador (Edge + Selenium)**  
   O Edge é aberto com preferências para **download direto** de PDF (sem prompts).  
   A sessão autenticada (SSO) — feita manualmente 1x/semana — é reaproveitada nas execuções diárias.

3. **Abertura de cada dashboard**  
   Para cada URL configurada, o script acessa o dashboard publicado no Power BI Service.

4. **Exportação para PDF**  
   Tenta a **exportação imediata** (botão “Exportar”), com **retries**.  
   Se não estiver disponível (mudança de layout, lentidão), usa **fallback**: espera o relatório ficar **estável** e então exporta.

5. **Tratamento do download**  
   Monitora a pasta do dia e só considera o arquivo quando o **PDF estiver completo** (sem extensões temporárias e com tamanho estabilizado).

6. **Regras de páginas (por dashboard)** — opcionais  
   - `extract_page`: extrai **apenas** uma página específica;  
   - `drop_pages`: **remove** páginas específicas;  
   - `drop_last_pages`: remove as **n últimas** páginas.  
   Assim, o consolidado fica **enxuto** e **focado**.

7. **Mesclagem**  
   Todos os PDFs tratados são **unidos em um único arquivo** consolidado do dia.

8. **Composição do e-mail em HTML**  
   O corpo do e-mail lista os dashboards processados (sem citar nomes no README) e anexa o **PDF único**.

9. **Envio via Outlook**  
   Integração com **Microsoft Outlook** (pywin32) para disparo do e-mail.

10. **Agendamento diário**  
    Usa `schedule` para rodar **todo dia** no horário configurado. Pode ser mantido em loop ou disparado via **Task Scheduler** do Windows.

---

## 📨 Por que existe um arquivo de e-mails (destinatários)?
Para **não “engessar” os destinatários no código**.  
- Mantemos um `.txt` ou `.csv` com **um e-mail por linha**.  
- O script **lê automaticamente** esse arquivo e envia para todos os endereços listados.

**Vantagens:**
- Qualquer pessoa pode **atualizar a lista** sem tocar no código (governança simples).
- Reduz riscos (evita recompilar/editar script para mudar destinatários).
- Facilita auditoria de **quem recebe** os indicadores.

> **Recomendação:** manter esse arquivo **fora do versionamento** (listado no `.gitignore`) para não expor e-mails publicamente.

---

## 📦 Requisitos
- **Python 3.8+**
- Bibliotecas:
  - `selenium`, `PyPDF2`, `python-dotenv` (opcional), `schedule`, `pywin32`
- **Microsoft Edge** instalado + `msedgedriver` compatível com a versão do Edge
- **Microsoft Outlook** instalado e logado

---

## ⚙️ Configuração (.env opcional)
Crie um arquivo **`.env`** (opcional). Se não existir, o script usa defaults do código.

```ini
# Pastas
DOWNLOAD_DIR=C:\Caminho\Para\Saida\Apresentacao
EMAIL_LIST_DIR=C:\Caminho\Para\ListaEmails

# Driver do Edge
MSEDGEDRIVER_PATH=C:\Caminho\Para\msedgedriver.exe

# Horário diário (HH:MM)
RUN_HOUR_MINUTE=15:50

# Ajustes finos (opcionais)
PRE_EXPORT_COOLDOWN_SEC=1
IDLE_STABLE_SECONDS=5
IDLE_TIMEOUT_SEC=90
DOWNLOAD_TIMEOUT_SEC=600
FORCE_EXPORT_IMMEDIATE=true
IMMEDIATE_TRIES=3
