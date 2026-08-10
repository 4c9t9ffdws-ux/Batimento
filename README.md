# Batimento Operacional - refatoracao em andamento

## Fluxo seguro

1. Execute uma macro `GerarPrevia...`.
2. Revise os movimentos na aba `BASE`, colunas `AJ:AO`.
3. Execute a macro `Executar...Aprovado` correspondente e confirme a mensagem.

`MOD_PLANO_LOTES.bas` contem a primeira rotina refatorada: `GerarPreviaBatimentoInventario`.
`MOD_EXECUCAO_APROVADA.bas` bloqueia as quatro rotinas SAP existentes quando nao houver previa ou quando a confirmacao for recusada.

## Estrutura

- `legacy/`: exportacao preservada dos 13 modulos originais.
- `MOD_PLANO_LOTES.bas`: calculo de transferencias de lote sem acesso ao SAP.
- `MOD_EXECUCAO_APROVADA.bas`: confirmacao central antes do envio SAP.

Os modulos de status e diario ainda usam a logica SAP legada depois da confirmacao. Eles devem ser migrados para gerar suas proprias matrizes de previa antes de o envio direto da matriz substituir as rotinas legadas.

