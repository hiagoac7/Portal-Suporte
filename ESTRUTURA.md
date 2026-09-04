# PortalSuporte — mapa da pasta

Reorganizado em 02–03/09/2026. A raiz agora só tem **um atalho, este arquivo e
as pastas**. Tudo mais está separado por função.

```
PortalSuporte\
│  LIGAR PORTAL.bat      ← clique aqui para subir o portal
│  ESTRUTURA.md          ← este arquivo
│
├─ app\                  o portal em si (código e imagens)
├─ dados\                banco, anexos, exportações e backups
├─ windows\              scripts de iniciar junto com o Windows
├─ documentos\           RH e laudos — não são do portal
└─ venv\                 ambiente Python (não mexer)
```

Ainda aparecem `.git`, `.gitignore`, `.streamlit`, `.claude` e `desktop.ini` —
são **arquivos ocultos** que as ferramentas exigem na raiz, e o Explorer não
mostra com as opções padrão.

## app\ — o portal

| Arquivo | Para que serve |
|---|---|
| `app.py` | O portal inteiro. É o que o Streamlit executa. |
| `marca_streamlit.py` | Escreve o título e o favicon TLM no HTML servido pelo Streamlit, antes de o servidor subir. Sem isso a aba pisca o nome antigo por ~2s. |
| `sanitize_app.py` | Remove bytes nulos do `app.py` antes de iniciar. A sincronização às vezes cola lixo no fim do arquivo e o Python recusa. |
| `tlm_icon.png` | Selo TLM usado como favicon. |
| `portal_logo.png` | Logo antigo — reserva. |

## dados\ — tudo que é informação

| Item | Conteúdo |
|---|---|
| `servicedesk.db` | Banco de produção: chamados e compras. |
| `anexos\` | Prints e arquivos enviados nos chamados. |
| `exports\` | Relatórios exportados pelo painel. |
| `backups\` | Uma cópia do banco por dia. O próprio app cria e apaga as antigas. |
| `backups\app\` | Instantâneos do `app.py` antes de cada alteração grande. |
| `backups\marcos\` | Cópias do banco de antes de migrações pontuais. Ficam guardadas. |

## windows\ — iniciar junto com o Windows

| Arquivo | Para que serve |
|---|---|
| `INSTALAR INICIO AUTOMATICO.bat` | **Comece por este.** Cria um atalho na pasta Inicializar. Não pede administrador. |
| `DESINSTALAR INICIO AUTOMATICO.bat` | Remove os dois métodos de uma vez. |
| `ALTERNATIVO - instalar via Tarefa Agendada.bat` | Só se o atalho não funcionar. Pede administrador e libera a porta 8501 no Firewall. |
| `iniciar_portal_oculto.pyw` | O que os dois métodos executam: saneia o `app.py`, aplica a marca e sobe o Streamlit sem janela. |
| `criar_atalho_inicio.ps1`, `instalar_portal_startup.ps1` | O trabalho de verdade por trás dos .bat. |
| `obsoletos\` | Os dois VBScript do método antigo. O instalador atual os apaga da Inicialização — ficam aqui só por histórico. |

## ⚠ Depois desta mudança, rode UMA VEZ

O início automático estava apontando para o caminho **antigo** do
`iniciar_portal_oculto.pyw`. Enquanto não for refeito, o portal **não sobe
sozinho ao ligar o PC** (ligar pelo `LIGAR PORTAL.bat` funciona normalmente).

Para consertar: abra `windows\` e dê dois cliques em
**`INSTALAR INICIO AUTOMATICO.bat`**. Ele limpa o registro antigo e aponta para
o novo lugar.

## Onde ficam os caminhos, no código

O `app.py` não depende mais do diretório de trabalho. Ele se localiza sozinho:

```python
_DIR_APP = os.path.dirname(os.path.abspath(__file__))   # ...\PortalSuporte\app
_RAIZ    = os.path.dirname(_DIR_APP)                    # ...\PortalSuporte
_DADOS   = os.path.join(_RAIZ, "dados")
```

E daí saem `DATABASE_SQLITE`, `ANEXO_DIR`, `BACKUP_DIR` e `EXPORT_DIR`. Por isso
o portal funciona igual iniciado pelo .bat, pela Inicialização do Windows ou
pelo VS Code — antes, iniciar de um lugar diferente criava um banco vazio.

Há uma rede de segurança: se o banco ainda estiver na raiz (mudança
interrompida, cópia antiga restaurada), o app usa **ele** em vez de criar um
vazio em `dados\`.

## Por que os backups do banco ficam soltos em `dados\backups\`

O `app.py` grava um `servicedesk_AAAAMMDD.db` por dia e apaga os mais antigos,
listando os arquivos que começam com `servicedesk_` **direto nessa pasta**. Se
essas cópias forem para uma subpasta, a rotina para de encontrá-las e nunca
mais poda nada.

Foi por isso que as três cópias `servicedesk_pre_*.db` saíram para
`backups\marcos\`: elas casavam com o mesmo padrão e, por ordem alfabética,
ocupavam as vagas das cópias que deviam ser mantidas — na prática o portal
apagava backup diário recente e guardava snapshot de junho.

## Rotina sugerida

- **Antes de mexer no `app.py`**: copie para `dados\backups\app\app_pre_<motivo>_<data>.py`.
- **Uma vez por mês**: se `backups\app\` passar de ~50 arquivos, apague os mais antigos.
- **Parar o portal**: Gerenciador de Tarefas → aba Detalhes → finalizar `pythonw.exe`.
