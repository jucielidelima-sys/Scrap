# Dashboard de Sucata e Retrabalho | GNO

Dashboard em Streamlit com base interna em Excel.

## Como executar

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Arquivos obrigatórios na mesma pasta

- app.py
- requirements.txt
- Sucata Retrabalho.xlsx
- NEWORDER IMAGEM(1).png

## Ajustes desta versão

- Cabeçalho reduzido e alinhado.
- Removido aviso incorreto de "Nenhum dado encontrado" que bloqueava os gráficos.
- Mantidos filtros completos.
- Pareto usando Descrição do Produto no eixo principal.
- KPIs formatados em padrão brasileiro.
