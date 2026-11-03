# Validação da entrega

Referência: `Desafio_ITALORAMOS.xlsx`, fornecida pelo usuário.

Comparação realizada em 03/11/2026 após executar o script com a base original. Todas as células das dez abas diferentes de PYTHON foram comparadas em dois modos: fórmulas e valores armazenados em cache, com tolerância numérica absoluta de 0,0000001.

Os valores e as fórmulas financeiros coincidiram com a referência. Foram encontradas três diferenças de conteúdo, em ambos os modos:

| Célula | Referência | Saída do script |
| --- | --- | --- |
| RESPOSTAS!D26 | Vazia | Indicação do arquivo Python e da aba com o código |
| RESPOSTAS!D33 | Nome completo do candidato | Vazia, para preenchimento pessoal |
| MEMORIA_CALCULO!B2 | Desafio_Controller_Italo.xlsx | Desafio_Controller_Italo (2).xlsx |

A aba PYTHON contém a versão atualizada do script; seu conteúdo difere da referência. A comparação não certifica identidade visual ou binária dos arquivos e não executa as fórmulas em um aplicativo de planilha: confere as expressões e seus caches. As validações internas do script também passaram.

Resultados conferidos: receita de R$ 1.413.947,23; custos de R$ 146.846,07; despesas de R$ 914.533,58; resultado de R$ 352.567,58; inadimplência de R$ 41.061,99; três processos sem honorário válido, totalizando R$ 23.393,67.
