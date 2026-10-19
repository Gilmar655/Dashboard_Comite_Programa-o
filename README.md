# Dashboard — Comitê de Programação

**Site público:** https://gilmar655.github.io/Dashboard_Comite_Programa-o/

**Repositório:** https://github.com/Gilmar655/Dashboard_Comite_Programa-o

## Base publicada em 19/10/2026

- **Origem:** arquivo `Comite_19_10_2026_periodo__05_22_novembro.xlsx`, aba `Base_Sigeo`.
- **Programação:** 05 a 22 de novembro de 2026.
- **71 programações** (70 códigos de projeto distintos).
- **25 programações encaminhadas ao comitê**: coluna `Decisão Comitê` marcada como `SIM - COMITÊ`, correspondente ao critério de CHI > 1.500.
- **46 liberadas para documentação**, **25 pendentes de aprovação**.
- **CHI somado por linha:** 153.183. O CHI pode se repetir entre registros ligados ao mesmo equipamento, não representando necessariamente clientes únicos.
- **18 linhas** com solicitação explícita de viabilidade de gerador.
- Todos os campos `Numero PowerON` estão vazios na planilha de origem.

## Como o site funciona

O dashboard é um único arquivo `index.html`, com gráficos, estilos, logomarca e base de dados incorporados. Não necessita de conexão com serviços externos para visualizar os registros. Mantém filtros por programação, status, tipo de intervenção, contratada, região, comitê e solicitação de gerador; tabela detalhada, busca e exportação Excel.

A importação de planilha Excel é **local ao navegador**: importar um arquivo não publica automaticamente a alteração para outras pessoas. Para publicar uma nova base, é necessário atualizar `index.html` na branch `main` deste repositório. O importador ignora linhas formatadas sem valores, importantes em arquivos Sigeo extensos.

## Publicação e cautelas

No GitHub, configure **Settings > Pages > Deploy from branch > main > / (root)**, caso necessário. Após uma atualização, use **Ctrl + F5** para evitar o cache anterior.

Este repositório é público e seus dados são acessíveis por terceiros. A planilha e o painel retratam o status fornecido e **não comprovam a conclusão das intervenções**.
