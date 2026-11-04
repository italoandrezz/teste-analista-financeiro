# Desafio de analista financeiro

Neste projeto, trato os dados do desafio de analista financeiro e apuro os resultados do terceiro trimestre de 2026.

Uso o script para classificar os gastos, calcular o resultado por unidade, consolidar a receita por produto, medir a inadimplência e conciliar os honorários por CPF.

Organizei a entrega em três arquivos:

- `Desafio_ITALORAMOS.xlsx`: planilha final com os resultados do desafio.
- `resolver_desafio.py`: script de tratamento, cálculo e validação dos dados.
- `README.md`: apresentação do projeto e instruções de execução.

Para gerar uma nova planilha a partir da base original do desafio, uso Python 3.10 ou superior e instalo a dependência com:

```powershell
python -m pip install openpyxl==3.1.5
```

Depois, informo o caminho da base original e o nome do arquivo de saída:

```powershell
python resolver_desafio.py "Desafio_Controller_Italo (2).xlsx" "resultado.xlsx"
```

No próprio script, valido a preservação das bases, os caches das fórmulas e as reconciliações dos resultados. Também conferi que os valores e as fórmulas financeiros gerados coincidem com a planilha final.
