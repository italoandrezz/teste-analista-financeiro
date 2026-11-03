# Validação da entrega

Usei `Desafio_ITALORAMOS.xlsx` como referência para conferir a entrega.

Em 03/11/2026, executei o script com a base original e comparei todas as células das dez abas diferentes de PYTHON em dois modos: fórmulas e valores armazenados em cache. Usei tolerância numérica absoluta de 0,0000001.

Confirmei que os valores e as fórmulas financeiros coincidem com a referência. Identifiquei três diferenças de conteúdo, em ambos os modos:

| Célula | Referência | Saída do script |
| --- | --- | --- |
| RESPOSTAS!D26 | Vazia | Indicação do arquivo Python e da aba com o código |
| RESPOSTAS!D33 | Nome completo do candidato | Vazia, para preenchimento pessoal |
| MEMORIA_CALCULO!B2 | Desafio_Controller_Italo.xlsx | Desafio_Controller_Italo (2).xlsx |

Incluí a versão atualizada do script na aba PYTHON, por isso seu conteúdo difere da referência. Conferi as expressões e seus caches, sem executar as fórmulas em um aplicativo de planilha nem comparar a aparência ou a estrutura binária dos arquivos. Também executei as validações internas do script, que passaram.

Conferi os seguintes resultados: receita de R$ 1.413.947,23; custos de R$ 146.846,07; despesas de R$ 914.533,58; resultado de R$ 352.567,58; inadimplência de R$ 41.061,99; três processos sem honorário válido, totalizando R$ 23.393,67.
