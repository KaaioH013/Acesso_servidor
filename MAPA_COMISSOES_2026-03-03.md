# Mapa de Comissões (base recebida em 03/03/2026)

Fonte: Tabela de vendedores e comissões.xlsx (aba Planilha1)

## Estrutura identificada

- Cabeçalho real na linha 2 da planilha.
- Colunas de comissão por tipo de produto/serviço:
  - Bombas
  - Peças
  - Anfíbias / Submerças
  - Aeradores
  - Serviço
- Coluna Região define territorialidade (ex.: SP, RS/SC, GO/MT, etc).
- Valor `-` foi interpretado como “não comissiona” para aquela combinação.
- Há linha de continuação para Fernando Guideli em região MS (mesmas regras do vendedor, sem repetir nome/tipo).

## Regras por vendedor (extração)

| Vendedor | Tipo | Região | Bombas | Peças | Anfíbias/Submerças | Aeradores | Serviço | Observação |
|---|---|---|---:|---:|---:|---:|---:|---|
| Alex Mendonça | Interno | Nordeste | 1,50% | 2,00% | 1,50% | 1,50% | 2,00% | - |
| Marcos Sachetto | Interno | SP | 2,00% | - | 1,50% | 1,50% | 2,00% | - |
| Fernando Guideli | Interno | SP | 2,00% | - | 1,50% | 1,50% | 2,00% | + regra adicional para MS |
| Fernando Guideli | Interno | MS | 2,00% | 2,00% | 1,50% | 1,50% | 2,00% | linha de continuação na planilha |
| Leandro | Interno | SC | 2,00% | - | 1,50% | 1,50% | 2,00% | - |
| Marcelo Bento | Interno | RS | 2,00% | 1,50% | 1,50% | 1,50% | 1,50% | - |
| Marco Baldo | Interno | RS/SC | 1,00% | 0,75% | 1,00% | 1,00% | 1,00% | - |
| Lucas Serra | Interno | SP/MG (Triângulo Mineiro) | 1,50% | - | 1,50% | 1,50% | 2,00% | - |
| Edelsio | Interno | Todas | - | - | 1,00% | 1,00% | 1,00% | Só na indústria |
| Wagner | Interno | Todas Industria | - | 0,35% | - | - | 0,35% | Só na indústria |
| Eduardo | Interno | Todas | 0,10% | 0,10% | 0,33% | 0,33% | 0,33% | Inside sobre todas as vendas (inclui Sanepumps; exclui exportação e intercompany) |
| Alexandre* | Rep. | MG (sem Triângulo Mineiro) | 5,00% | 2,00% | 5,00% | 5,00% | 2,00% | * ver observações gerais |
| Sebastião* | Rep. | GO/MT | 5,00% | 2,50% | 5,00% | 5,00% | 2,50% | peças/serviços 2,5% somente em 2026 |
| Rafael* | Rep. | PR | 5,00% | 2,00% | 5,00% | 5,00% | 2,00% | * ver observações gerais |
| Mario* | Rep. | SP | 5,00% | - | 5,00% | 5,00% | 2,00% | * ver observações gerais |
| Zé Luiz* | Rep. | SP | 5,00% | - | 5,00% | 5,00% | 2,00% | * ver observações gerais |

## Observações gerais da própria planilha

- “Vendas para Mineração e Petróleo e Gás: consultar comissão com a Diretoria.”
- “Licitações de órgãos públicos para representantes: comissão reduzida para 3%.”
- “Vendas de vendedores fora de sua região: consultar Diretoria.”

## Implicações para implementação no cálculo

1. Precisamos de uma tabela de regras parametrizada por:
   - vendedor
   - tipo (interno/rep)
   - região
   - tipo de produto
   - vigência (para tratar exceções por ano, como Sebastião em 2026)

2. Regras de exceção devem ficar separadas da regra base:
   - licitação (3% para representante)
   - mineração/petróleo e gás (alçada diretoria)
   - fora da região (alçada diretoria)

3. Para o cálculo final, o fluxo ideal será:
   - identificar vendedor responsável e região da venda
   - classificar tipo de produto (peça/bomba/anfíbia/aerador/serviço)
   - aplicar regra base
   - aplicar exceção (se houver)
   - auditar resultado com trilha no relatório

## Status

- Entendimento da planilha: concluído.
- Próximo passo sugerido: transformar este mapa em tabela de regras (CSV/aba de configuração) e plugar no relatório de comissões.
