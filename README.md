# Desafio de analista financeiro

Neste projeto, trato os dados do desafio de analista financeiro e apuro os resultados do terceiro trimestre de 2026.

Preservo as bases originais e uso o script para classificar os gastos, calcular o resultado por unidade, consolidar a receita por produto, medir a inadimplência e conciliar os honorários por CPF. Mantenho as fórmulas e seus valores em cache na planilha para permitir a consulta sem recálculo imediato.

Para executar o projeto, uso Python 3.10 ou superior e instalo a dependência com:

```powershell
python -m pip install -r requirements.txt
python resolver_desafio.py
```

Quando preciso usar outros arquivos, informo os caminhos de entrada e saída:

```powershell
python resolver_desafio.py "entrada.xlsx" "saida.xlsx"
```

Uso `Desafio_Controller_Italo (2).xlsx` como entrada padrão e salvo o resultado em `Desafio_Italo_Andre_Mendes_da_Silva_Ramos_Completo.xlsx`.

No próprio script, valido a preservação das bases, os caches das fórmulas e as reconciliações dos resultados. Reservo o preenchimento da declaração de autoria em `RESPOSTAS!D33` para depois da leitura do seu conteúdo.

Registrei a comparação com a planilha de referência em [VALIDACAO.md](VALIDACAO.md).
