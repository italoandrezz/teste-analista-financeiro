# Desafio de analista financeiro

Projeto de Italo Ramos para tratamento dos dados e apuração financeira do terceiro trimestre de 2026.

O script preserva as bases originais e gera classificação de gastos, resultado por unidade, receita por produto, inadimplência e conciliação de honorários por CPF. As planilhas contêm fórmulas e valores em cache para consulta sem recálculo imediato.

Requer Python 3.10 ou superior.

```powershell
python -m pip install -r requirements.txt
python resolver_desafio.py
```

Também é possível informar os arquivos de entrada e saída:

```powershell
python resolver_desafio.py "entrada.xlsx" "saida.xlsx"
```

A entrada padrão é `Desafio_Controller_Italo (2).xlsx`. A saída padrão é `Desafio_Italo_Andre_Mendes_da_Silva_Ramos_Completo.xlsx`.

O script valida a preservação das bases, os caches das fórmulas e as reconciliações dos resultados. A declaração de autoria deve ser preenchida pessoalmente pelo candidato.

Veja `VALIDACAO.md` para a comparação com a planilha de referência.
