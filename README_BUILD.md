# Analise de atendimentos

## Configuracao segura

Nunca coloque as credenciais no codigo. Crie um arquivo `.env` ao lado do `analise_atendimentos.py` ou ao lado do `.exe`:

```env
AUVO_API_KEY=sua_chave
AUVO_API_TOKEN=seu_token
```

O arquivo `.env` fica fora do Git pelo `.gitignore`. Para publicar no GitHub, publique apenas `.env.example`.

## Modelo Excel

Coloque `Relatorio_Atendimentos_Mai2026.xlsx` ao lado do `.py` ou do `.exe`. Se preferir outro local, defina:

```env
ANALISE_TEMPLATE_EXCEL=C:\caminho\Relatorio_Atendimentos_Mai2026.xlsx
```

## Gerar executavel

```powershell
pip install -r requirements.txt
playwright install chromium
pyinstaller --onefile --name analise_atendimentos analise_atendimentos.py
```

Para distribuir para outro computador, envie:

- `dist\analise_atendimentos.exe`
- `.env` com as credenciais daquele ambiente
- `Relatorio_Atendimentos_Mai2026.xlsx`

O login do Gemini e salvo em `%APPDATA%\AnaliseAtendimentos\gemini_session.json`, fora da pasta do Git.

## Dados brutos

Por padrao, o TXT com prompt completo e JSON bruto nao e salvo. Para ativar localmente:

```env
ANALISE_SALVAR_DADOS_BRUTOS=1
```
