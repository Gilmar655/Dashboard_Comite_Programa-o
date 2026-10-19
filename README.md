# Dashboard Comitê de Programação — Enel Brasil

**Site:** https://gilmar655.github.io/Dashboard_Comite_Programa-o/

**Repositório:** https://github.com/Gilmar655/Dashboard_Comite_Programa-o

## Publicar esta versão no GitHub

**IMPORTANTE: extraia a pasta ZIP antes de enviar ao GitHub.** Publique todos os arquivos desta pasta diretamente na raiz (`/`) da branch `main`, **não dentro de uma subpasta**. Substitua `index.html` e os demais arquivos de nomes iguais, adicionando os que não existirem.

Arquivos:

- `index.html` — página inicial e painéis.
- `styles.css` — layout profissional responsivo.
- `app.js` — indicadores, gráficos SVG, filtros, alertas e tabela.
- `excel.js` — leitor/escritor de Excel e CSV sem CDN.
- `programacoes.js` — base incorporada no site; **obrigatório para mostrar as informações para todos os visitantes**.
- `programacoes.csv` — base para conferência e download direto.
- `.nojekyll` — impede processamento Jekyll desnecessário.

No GitHub: `Settings > Pages > Build and deployment > Deploy from a branch > main > / (root)`.

Depois da publicação, acesse a URL e pressione Ctrl + F5 se visualizar dados em cache.

## Dados e indicadores

- Fonte: `Comite_19_10_2026_periodo__05_22_novembro.xlsx`, aba `Base_Sigeo`.
- Data da base: **19/10/2026**. Programação: **05 a 22 de novembro de 2026**.
- **71 programações**, 25 comitê, 46 liberadas para documentação, 25 pendentes de aprovação.
- **153.183 CHI acumulado por registro** e 18 solicitações textuais de viabilidade de gerador.
- Critério comitê: CHI acima de 1.500 e sinalização explícita na coluna do Excel.
- Gráficos SVG nativos, relógio de Brasília, filtros por datas, empresa, tipo, região, status, gerador e comitê, busca livre, tabela completa e exportação/importação Excel e CSV.

**Nota sobre a importação:** quando alguém importa um Excel no navegador, os dados só mudam para aquele visitante. Para atualizar a base visível a todos, substitua a base `programacoes.js` (e opcionalmente `programacoes.csv`) no repositório GitHub.

**Cuidado com dados públicos:** como o repositório é público, qualquer visitante pode acessar os registros. Valide autorização institucional para divulgação. CHI pode estar repetido entre linhas associadas ao mesmo equipamento, não devendo ser interpretado automaticamente como número de clientes únicos. Os status são os da fonte e não comprovam execução.
