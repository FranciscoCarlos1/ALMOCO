# Sistema de Intenção de Almoço - IFC SBS

Aplicação web simples para substituir o formulário impresso de intenção de almoço.

## Turmas cadastradas
- TIN I, II e III
- TAI I, II e III
- TST I, II e III

## O que a aplicação faz
- Formulário para alunos informarem se vão almoçar (`SIM` ou `NAO`)
- Registro por data
- Atualização da resposta (se o mesmo aluno enviar de novo no mesmo dia)
- Busca automática de aluno por matrícula (quando a lista estiver importada)
- Painel administrativo com resumo por turma
- Exportação em CSV para planilha
- Importação de lista de alunos por CSV
- Planilha semanal no formato do ODS (`N`, `Nome`, `Seg`, `Ter`, `Qua`, `Qui`, `Sex`)

## 1) Instalação
No terminal, dentro da pasta do projeto:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## 2) Execução

```powershell
python app.py
```

Acesse:
- Formulário do aluno: `http://localhost:5000/`
- Painel: `http://localhost:5000/admin?token=ifc-sbs`
- Planilha semanal (modelo ODS): `http://localhost:5000/admin/planilha?token=ifc-sbs&turma=TIN%20I`

## Importar lista de alunos
No painel (`/admin`), use o bloco **Importar alunos** e envie um CSV com cabeçalho:

```csv
nome,matricula,turma
Aluno Exemplo,2026001,TIN I
```

Aceita também CSV separado por `;`.

## 3) Segurança básica
Troque o token padrão do painel:

```powershell
$env:ALMOCO_ADMIN_TOKEN="SEU_TOKEN_FORTE"
python app.py
```

Depois use:
`http://localhost:5000/admin?token=SEU_TOKEN_FORTE`

## 4) Compartilhar com os alunos
Se o servidor estiver ligado em um computador da escola e liberado na rede local, os alunos podem responder pelo celular usando o IP da máquina, por exemplo:

`http://192.168.0.10:5000/`

## 5) Banco de dados
Os dados ficam em:
- `data/almoco.db`

Você pode exportar diariamente no painel para gerar arquivos CSV e abrir no LibreOffice/Excel.

### Usar PostgreSQL (recomendado para produção)
O sistema agora aceita PostgreSQL via variável de ambiente `DATABASE_URL`.

Exemplo local (PowerShell):

```powershell
$env:DATABASE_URL="postgresql://USUARIO:SENHA@HOST:5432/NOME_DO_BANCO"
python app.py
```

Sem `DATABASE_URL`, o sistema continua usando SQLite (`data/almoco.db`).

#### Migrar dados atuais (SQLite -> PostgreSQL)

```powershell
$env:DATABASE_URL="postgresql://USUARIO:SENHA@HOST:5432/NOME_DO_BANCO?sslmode=require"
python scripts/migrate_sqlite_to_postgres.py
```

Script utilizado: `scripts/migrate_sqlite_to_postgres.py`

## 6) Publicar no Render
Este projeto já está preparado com `render.yaml`.

### Passos
1. Suba este projeto para um repositório no GitHub.
2. No Render, clique em **New +** -> **Blueprint**.
3. Conecte seu repositório e confirme o deploy.
4. O Render vai criar o serviço web com:
	- `gunicorn app:app`
	- disco persistente em `/var/data`
	- token admin automático (`ALMOCO_ADMIN_TOKEN`)

### Variáveis de ambiente no Render
- `ALMOCO_ADMIN_TOKEN` (já gerada automaticamente)
- `ALMOCO_DATA_DIR=/var/data`
- `DATABASE_URL=postgresql://...` (recomendado; banco gerenciado externo)

### URLs após deploy
- Formulário: `https://SEU-APP.onrender.com/`
- Painel: `https://SEU-APP.onrender.com/admin?token=SEU_TOKEN`
- Planilha semanal: `https://SEU-APP.onrender.com/admin/planilha?token=SEU_TOKEN&turma=TIN%20I`
