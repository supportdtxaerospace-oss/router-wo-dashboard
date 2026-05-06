# Router WO Dashboard

Dashboard Streamlit que replica os gráficos da aba `Data` do arquivo Excel de acompanhamento de WOs.

## Requisitos
- Python 3.11+
- Dependências em `requirements.txt`

## Instalação local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Atualizar os dados

1. Substitua o arquivo `data/Router Graphic Report MK 0.xlsx` pela versão mais recente.
2. Faça commit e push para o GitHub:
   ```bash
   git add data/
   git commit -m "atualiza dados - <data>"
   git push
   ```
3. O Streamlit Community Cloud detecta o push e atualiza o dashboard automaticamente em ~1 minuto.
4. Acesse o link do app e confirme que os números batem com a aba `Contagem` do Excel.

## Estrutura do projeto

```
.
├── app.py                  # Dashboard principal
├── requirements.txt        # Dependências Python
├── .streamlit/
│   └── config.toml         # Configurações do servidor e tema
├── data/
│   └── Router Graphic Report MK 0.xlsx  # Arquivo fonte de dados
└── README.md               # Este arquivo
```

## Mapeamento de dados

| WO   | Status (col F) | Quantidade (col H) | Título (col B) |
|------|---------------|-------------------|---------------|
| WO #1 | F30:F33      | H30:H33           | B30           |
| WO #2 | F38:F41      | H38:H41           | B38           |
| WO #3 | F46:F49      | H46:H49           | B46           |

## Publicação no Streamlit Community Cloud

1. Suba o repositório no GitHub (público ou privado com acesso).
2. Acesse [share.streamlit.io](https://share.streamlit.io) e conecte o repo.
3. Defina `app.py` como entry point e faça o deploy.
4. Compartilhe o link gerado com o time.
