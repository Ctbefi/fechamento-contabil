# HANDOFF.md — Sistema Web de Fechamento Contábil (Efí)

---

## 1. Resumo Executivo

**O que é:** Um sistema web (dashboard) de acompanhamento de demandas do fechamento contábil mensal da Efí S.A., cobrindo 5 empresas (IP, SCFI, EVA, Lesta, Holding). Lê os dados de uma planilha Excel e apresenta KPIs, gráficos e tabelas com visual moderno (tema escuro, "cara de sistema").

**Objetivo final:** Substituir a navegação por arquivos Excel soltos no SharePoint por um sistema visual único, acessível pela equipe de Contabilidade, com proteção por senha.

**Estado atual:** Sistema HTML funcional e completo, hospedado no **GitHub Pages** em `https://ctbefi.github.io/fechamento-contabil/`. Protegido por senha. Está em fase de ajuste final de publicação (problema de cache do navegador sendo resolvido). O dashboard já lê o Excel corretamente (upload manual, URL do SharePoint ou auto-fetch quando hospedado no próprio SharePoint).

**Últimas melhorias implementadas (ver Seção 8):**
1. Correção do truncamento do quadro "Demandas pendentes com prazo vencido" no Dashboard (mostrava só 10 registros; agora mostra todos, ordenados por prazo mais antigo primeiro).
2. Nova coluna **Observações** na aba Operacional, lida diretamente da planilha (detecção tolerante a espaços extras no cabeçalho).
3. Novo card **"Indicadores de performance"** no Dashboard, ao lado de "Distribuição de status" (agora 3 cards lado a lado), com Total de demandas / Concluídas no prazo / Entregues fora do prazo e um gráfico de pizza (No Prazo x Atrasado x Em andamento), respeitando o filtro de empresa. Detalhes: Seção 8.5.

---

## 2. Contexto Essencial

### Sobre o usuário (Alessander)
- Trabalha na equipe de contabilidade da Efí S.A. (domínio `sejaefi.com.br` / `gerencianet.sharepoint.com`)
- Prefere orientação **passo a passo**, **um passo de cada vez**, em linguagem não técnica
- Trabalha de forma iterativa, enviando prints para validar cada etapa

### Caminhos já descartados (não retomar)
- **Microsoft Lists + Power Apps + Azure AD:** TI da Efí **não libera** registro de aplicativo no Azure AD (erro 401 — sem role "Application Developer"). Qualquer solução que dependa de Azure AD App Registration está bloqueada.
- **Power BI:** Foi construído um relatório funcional (workspace "Fechamento Contábil - Contabilidade"), mas o visual não atingiu a "cara de sistema" desejada. As 5 tabelas separadas dificultam comparações. O usuário preferiu o HTML.
- **Netlify:** Foi avaliado como opção de hospedagem com senha nativa, mas o usuário optou por GitHub Pages.

### Definições de status (REGRA DE NEGÓCIO CRÍTICA)
A coluna S (Situação da entrega — Fechamento) tem estes valores. A distinção é essencial:

| Status | Significado | É pendente? |
|--------|-------------|-------------|
| **No Prazo** | Entregue dentro do prazo (col R preenchida, data ≤ prazo) | Não (concluída) |
| **Concl. c/ atraso** ("Atrasado") | Entregue APÓS o prazo (col R preenchida, data > prazo) | **Não (concluída)** |
| **Em atraso** | Prazo vencido, col R VAZIA → pendente crítica | **SIM** |
| **A iniciar** | Prazo futuro, col R vazia | Não |
| **Fechamento prévio** | Digitar "Prévio" na col R = ação paliativa aplicada | Não |
| **N/A** | Não se aplica | Não |

**REGRA:** Apenas "Em atraso" conta como demanda pendente/crítica (aparece no quadro de críticos do Dashboard). "Atrasado" = concluída com atraso, NÃO é pendente. O usuário corrigiu isso explicitamente.

### Conclusão por empresa
% de conclusão = (demandas com data de entrega preenchida) / (total da empresa). Ou seja, "No Prazo" + "Atrasado" = concluídas. Exibido como `X/Y concl.` por empresa.

---

## 3. Arquivos, Documentos e Fontes

### Sistema Web (arquivos ATUAIS para GitHub Pages)
| Arquivo | Conteúdo |
|---------|----------|
| `index.html` | Estrutura da página (login + load + app). CSS e JS externos. **NÃO tem CSS inline** |
| `styles.css` | Todo o tema escuro (dark theme) |
| `app.js` | Toda a lógica: senha, leitura Excel, render de KPIs/gráficos/tabelas |
| `.nojekyll` | Arquivo vazio que desativa o Jekyll do GitHub Pages (causava CSS vazando como texto) |
| `README.md` | Instruções de publicação e troca de senha |

> **IMPORTANTE:** A versão monolítica `Sistema_Fechamento_Contabil_Efi.html` (CSS+JS inline) está OBSOLETA. O GitHub Pages processava via Jekyll e quebrava o `<style>`. A solução foi separar em 3 arquivos + `.nojekyll`. Sempre usar a versão de 4 arquivos.

### Planilha de dados (fonte)
- Nome padrão: `Acompanhamento de demandas - Fechamento MM-2026.xlsx`
- Local no SharePoint: `https://gerencianet.sharepoint.com/sites/contabilidade2/Shared Documents/GESTÃO CONTÁBIL/Cronogramas e agendas/2026/Fechamento mensal 2026`
- Estrutura: 5 abas de empresa (IP, SCFI, EVA, Lesta, Holding) + aba Feriados + aba Parâmetros
- Mês de referência: célula D1 de cada aba

### Mapa de colunas da planilha (índices base-0 no app.js)
```
0=A Demandas | 1=B Descrição | 2=C Periodicidade | 3=D Data-base | 4=E Responsável
5=F Setor dep | 6=G Demanda dep | 7=H Prazo OS | 8=I Data Prazo OS | 9=J Data Entrega OS
10=K Situação OS | 11=L Prazo CTB | 12=M Data Prazo CTB | 13=N Data Entrega CTB
14=O Situação CTB | 15=P Prazo Total | 16=Q Data Prazo Fch | 17=R Data Entrega Fch
18=S Situação Fch | 19=T Observações
```
O app lê: demanda (0), descrição (1), responsável (4), prazo fch (16=Q), entrega fch (17=R), situação fch (18=S), **observações (19=T, desde a melhoria de ago/2026)**. Dados começam na linha 3 do Excel (índice 2).

> A coluna Observações não é mais lida por índice fixo cego: o app procura nas 3 primeiras linhas da aba um cabeçalho cujo texto (normalizado, sem acento/espaços extras) seja "observacoes"; só usa o índice 19 como *fallback* se não encontrar. Isso torna a leitura tolerante a pequenas variações de formatação do cabeçalho.

---

## 4. Decisões Já Tomadas

| Decisão | Motivo | Definitiva? |
|---------|--------|-------------|
| Hospedar como HTML no GitHub Pages | TI não libera Azure AD; usuário quer visual de sistema | Sim |
| Proteção por senha via SHA-256 no JS | Restringir acesso sem depender de servidor | Sim (ver ressalva na Seção 6) |
| Senha padrão: `Efi@Fechamento2026` | Hash: `c5d6126c1f876f8a2a1ae01486dd6d9578c837f98f8a2cf4a71a985ffa85cb9a` | Alterável |
| CSS/JS em arquivos separados + `.nojekyll` | GitHub Pages/Jekyll quebrava CSS inline | Sim |
| Gráfico de status = barras (não donut) | Donut ficava gigante e distorcia o layout | Sim |
| 7 KPIs no topo | Total, Concluído no prazo, A iniciar, Em atraso, Concl. c/ atraso, Fechamento prévio, N/A | Sim |
| Filtro por empresa no Dashboard | Permite ver cada empresa individualmente, não só o consolidado | Sim |
| Só "Em atraso" é pendente/crítico | "Atrasado" = concluído com atraso | Sim (regra de negócio) |
| Coluna "Descrição" na tabela | Entre Demanda e Responsável | Sim |
| Tema escuro (#0f1117) hardcoded | Visual moderno pedido pelo usuário | Sim |
| Quadro de pendentes vencidas do Dashboard mostra **todos** os registros (sem limite fixo) | O `.slice(0,10)` anterior escondia demandas críticas; usuário exigiu visão completa | Sim |
| Quadro de pendentes vencidas ordenado por prazo mais antigo primeiro | Prioriza visualmente o que está vencido há mais tempo | Sim |
| Coluna **Observações** exibida na aba Operacional (col. T = índice 19 da planilha) | Informação já existia na planilha mas não era importada/exibida | Sim |
| Observações editável em memória seguindo o mesmo padrão da Data de entrega (não grava no Excel) | Consistência com o padrão de edição já existente; sem mudança estrutural | Sim |

---

## 5. Pendências e Próximos Passos

### Imediato (em andamento)
1. **Resolver cache do navegador:** O usuário via CSS como texto porque o navegador cacheou a versão antiga. Solução: Ctrl+Shift+R ou aba anônima. Confirmar que os 4 arquivos novos renderizam corretamente.

### Curto prazo
2. **Validar auto-fetch do SharePoint:** O auto-carregamento (`autoLoadSharePoint()`) só funciona se o HTML for aberto DENTRO do domínio SharePoint (por CORS). No GitHub Pages, o auto-fetch vai falhar e cair no fallback manual (URL ou upload). Avaliar se o usuário aceita usar o fallback ou se quer hospedar dentro do próprio SharePoint.
3. **Testar login + carregamento completo** com a planilha real de julho.

### Médio prazo
4. **Replicar para novos meses:** Processo mensal = salvar novo Excel no SharePoint. O sistema busca o mais recente automaticamente (quando auto-fetch funcionar) ou o usuário cola a URL.
5. **Considerar edição bidirecional:** Hoje o "Salvar" na aba Operacional só altera em memória (não grava no Excel). Se quiser gravar de volta, precisaria de integração server-side (fora do escopo atual sem Azure AD).

---

## 6. Riscos, Dúvidas e Pontos de Atenção

### Segurança da senha (limitação conhecida)
A senha é verificada via JavaScript no navegador (hash SHA-256 no `app.js`). **Qualquer pessoa com o link pode ver o código-fonte e o hash.** Isso NÃO é segurança forte — protege contra acesso casual, não contra alguém determinado. O repositório GitHub é **público** (Pages gratuito exige repo público, ou GitHub Pro pago para privado). Para segurança real, seria necessário: repo privado + GitHub Pro, ou Netlify com HTTP Basic Auth (senha no servidor), ou hospedar no SharePoint (autenticação Microsoft nativa).

### CORS no auto-fetch
`autoLoadSharePoint()` e `loadFromUrl()` usam `fetch(..., {credentials:'include'})`. Isso só funciona quando a página está no mesmo domínio do SharePoint OU quando o SharePoint permite CORS para a origem. No GitHub Pages (domínio diferente), o navegador bloqueia. Por isso o fallback de upload manual existe. **Não prometer auto-fetch funcionando no GitHub Pages sem testar.**

### Jekyll / GitHub Pages
Se voltar a aparecer CSS/JS como texto: confirmar que `.nojekyll` está na raiz do repo e que `index.html` referencia `styles.css` e `app.js` externos (não inline). Nunca voltar para HTML monolítico com `<style>` inline no GitHub Pages.

### Repositório atual
- GitHub user: `Ctbefi`
- **Dois repositórios mantidos em paralelo** (decisão do usuário, ago/2026):
  - `fechamento-contabil` (público) — `https://ctbefi.github.io/fechamento-contabil/` — repositório original
  - `fechamento-contabil2` (público) — `https://ctbefi.github.io/fechamento-contabil2/` — repositório novo, GitHub Pages já ativado (Deploy from a branch, main, /root)
- Os 4 arquivos (`index.html`, `styles.css`, `app.js`, `.nojekyll`) precisam ser enviados manualmente para **cada** repositório a cada atualização — não há sincronização automática entre eles. Confirmar com o usuário se ele quer manter os dois sempre idênticos ou se vão divergir com o tempo.

---

## 7. Preferências de Execução

- **Um passo de cada vez.** Nunca dar duas orientações simultâneas quando guiando por interface. O usuário pediu isso explicitamente.
- **Não ficar esperando o usuário pedir o próximo passo** — encadear a sequência, mas um comando por vez.
- Linguagem **não técnica**, orientação por prints.
- Ao gerar arquivos, sempre entregar via `present_files` e copiar para `/mnt/user-data/outputs`.
- **Ser honesto sobre limitações** (o usuário valoriza isso — já apontou quando promessas não se concretizaram, ex: Embed web part do SharePoint bloqueando JS).
- Ao editar o HTML/CSS/JS: **cuidado com edições incrementais** que corromperam o arquivo antes. Se o arquivo ficar grande e frágil, recriar do zero é mais seguro.
- Aplicar correções **aditivamente** — não reverter decisões já validadas.

---

## 8. Melhorias Implementadas (Ago/2026)

### 8.1 Dashboard — fim do truncamento do quadro de pendentes vencidas

**Causa raiz identificada:** a função `critTable()` em `app.js` aplicava `.slice(0,10)` sobre a lista já filtrada por `situacaoFch==='Em atraso'`, limitando a exibição a 10 linhas mesmo quando havia mais demandas pendentes vencidas. Não havia limitação por CSS (sem `max-height`/`overflow` no `.card` ou `.tbl`) — o corte era 100% no JavaScript.

**Ajuste realizado:**
- Removido o `.slice(0,10)`.
- Adicionada ordenação pelo prazo mais antigo primeiro (`dataPrazoFch` convertida via nova função `parseBR()`, que interpreta o formato `DD/MM/AAAA`). Registros sem prazo reconhecível vão para o final, sem quebrar a ordenação.
- Filtro por empresa (`dashData()`), regra "só Em atraso conta como pendente" e o layout do quadro (colunas, cabeçalho, estado vazio) permanecem exatamente como antes.

### 8.2 Aba Operacional — coluna Observações

**Como foi localizada e incorporada:**
- Nova função `acharColObservacoes(rows)`: varre as 3 primeiras linhas da aba em busca de um cabeçalho cujo texto — normalizado via `normalizeHeader()` (remove acentos, colapsa espaços extras, minúsculas) — seja igual a `"observacoes"`. Se não encontrar, cai no índice fixo 19 (coluna T), que é a posição documentada na Seção 2.
- O valor é lido em `processWorkbook()` e incorporado a cada objeto de dado como `observacoes` (string; `undefined`/`null`/vazio viram `''`, nunca aparecem como texto literal "undefined"/"null").
- Exibida na aba Operacional como nova coluna, ao final da tabela (antes do botão de ação), sem alterar nenhuma coluna existente.
- Texto longo: célula com `white-space:normal; word-break:break-word` (classe `.obs-cell`), preservando o layout da tabela.
- Segurança de renderização: nova função `escHtml()` escapa `&`, `<`, `>`, `"` para que aspas ou caracteres especiais na observação não quebrem a tabela nem o campo de edição.
- Edição: seguindo o padrão já existente (editar/salvar em memória), a Observações ganhou um campo de texto editável junto com a Data de entrega; `saveR()` agora também persiste esse valor em memória. Nenhuma gravação de volta no Excel foi implementada (fora do escopo desta etapa).

### 8.3 Arquivos alterados

| Arquivo | Alteração |
|---|---|
| `app.js` | Adicionadas as funções `parseBR`, `normalizeHeader`, `acharColObservacoes`, `escHtml`; `processWorkbook()` passa a extrair `observacoes`; `critTable()` sem `.slice(0,10)` e com ordenação; `renderOp()` e `saveR()` com suporte à coluna Observações |
| `styles.css` | Adicionadas as classes `.obs-cell` e `.inp-obs` (2 linhas) |
| `index.html` | Sem alterações (cabeçalhos e tabelas são gerados dinamicamente pelo `app.js`) |

Login, KPIs, gráfico de status, aba Alertas e as regras de status ("Em atraso" vs. "Atrasado") não foram tocados — confirmado por diff linha a linha contra os arquivos originais.

### 8.4 Testes realizados

Executados via simulação de dados (Node.js), reproduzindo a lógica de `app.js`:

| Cenário pedido | Resultado |
|---|---|
| Empresa sem demandas pendentes vencidas → estado vazio | ✅ 0 registros, mensagem de vazio mantida |
| Empresa com poucas demandas vencidas → todas aparecem | ✅ 3/3, ordenadas pelo prazo mais antigo |
| Empresa com muitas demandas vencidas → sem truncamento | ✅ 25/25 exibidas |
| Troca do filtro de empresa → recalcula corretamente | ✅ |
| Atividade "Atrasado" (concluída com atraso) não entra no quadro | ✅ 0 registros |
| Atividade "Em atraso" (pendente vencida) entra no quadro | ✅ 1 registro |
| Cabeçalho "Observações" com espaços extras/maiúsculas é reconhecido | ✅ índice 19 detectado corretamente |
| Observação preenchida aparece na linha correta, com aspas/tags escapados | ✅ |
| Observação vazia ou coluna ausente na linha → célula vazia, sem "undefined"/"null" | ✅ |
| Sintaxe geral do `app.js` | ✅ `node -c app.js` sem erros |

Teste manual no navegador (upload da planilha real) ainda deve ser feito por Alessander antes da publicação definitiva no GitHub Pages, como confirmação final.

### 8.5 Dashboard — card "Indicadores de performance" (3º card + gráfico de pizza)

**Pedido:** adicionar um 3º card ao lado de "Distribuição de status" (ficando: Conclusão por empresa | Distribuição de status | Indicadores de performance), respeitando o filtro de empresa (Todas/IP/SCFI/EVA/Lesta/Holding), com Total de demandas, Total concluído no prazo e Total entregue fora do prazo, mais um gráfico de pizza com os %.

**Decisão de negócio confirmada com o usuário:** durante o pergunta feita, ficou definido que o gráfico de pizza tem **3 fatias**: `No Prazo`, `Atrasado` e `Em andamento` (esta última soma `A iniciar` + `Em atraso` + `Fechamento previo`) — as demandas em andamento entram no cálculo do %, não ficam de fora.

**Ajuste adicional assumido (não perguntado explicitamente, sinalizado ao usuário):** a categoria `N/A` também foi agrupada dentro de "Em andamento" no gráfico, para que as 3 fatias sempre somem 100% do Total de demandas da empresa/filtro selecionado. Sem isso, o círculo do gráfico não fecharia quando houvesse demandas N/A.

**Implementação:**
- `.section-row` (CSS) passou de `grid-template-columns:1fr 1fr` para `1fr 1fr 1fr` — os 3 cards dividem a largura da página igualmente, sem exceder a largura atual.
- Novo card no `index.html`: `<div id="perfStats">` (3 linhas de texto) + `<canvas id="perfChart">` (gráfico de pizza via Chart.js, já carregado no projeto).
- Nova função `perfCard()` em `app.js`: usa `dashData()` (o mesmo filtro de empresa já usado pelos outros elementos do Dashboard), calcula `tot`, `noPrazo` (`situacaoFch==='No Prazo'`), `atrasado` (`situacaoFch==='Atrasado'`) e `andamento = tot - noPrazo - atrasado` (por diferença, garante que a soma sempre bate com o total).
- Percentuais exibidos no texto e no gráfico são sempre em relação ao **Total de demandas** da empresa/filtro selecionado (não só das já concluídas) — consistente com o gráfico de pizza fechando em 100%.
- `perfCard()` é chamada junto com `statusChart()` (no `renderAll()` inicial e no `setDE()` ao trocar o filtro de empresa do Dashboard), então reage a todos os filtros: Todas, IP, SCFI, EVA, Lesta, Holding.
- Empresa/filtro sem nenhuma demanda: texto mostra zeros e o gráfico fica oculto (sem tentar desenhar um pie vazio).
- Cores seguem a paleta já usada no sistema: verde `#4ade80`/`#276221` (No Prazo), vermelho `#f87171`/`#9C0006` (Atrasado), azul `#3d7dd4` (Em andamento — cor de destaque do tema).

**Arquivos alterados nesta melhoria:** `index.html` (novo card), `styles.css` (grid 3 colunas + `.perf-row`/`.perf-label`/`.perf-val`), `app.js` (`perfCard()` + chamadas em `renderAll()`/`setDE()`). Login, KPIs do topo, gráfico de barras "Distribuição de status", aba Operacional, aba Alertas e regras de status não foram tocados (confirmado por diff).

**Testes realizados (simulação em Node, replicando a lógica de `perfCard`):**

| Cenário | Resultado |
|---|---|
| Exemplo do usuário (138 total, 46 no prazo, 92 atrasado, 0 andamento) | ✅ 33% / 67%, soma bate com o total |
| Cenário com demandas em andamento misturadas (A iniciar + Em atraso + Fechamento prévio + N/A) | ✅ `andamento` calculado corretamente, 3 fatias somam 100% do total |
| Empresa/filtro sem nenhuma demanda | ✅ não quebra (sem divisão por zero) |
| Filtro por empresa (ex.: só IP, só SCFI) | ✅ indicadores calculados isoladamente por empresa |
| Sintaxe geral do `app.js` | ✅ `node -c app.js` sem erros |

Teste visual no navegador (conferir alinhamento dos 3 cards e o gráfico de pizza renderizando) ainda precisa ser feito por Alessander.

### 8.6 Ajustes: remoção da aba Alertas, remoção da edição de data e Observações no Dashboard

Três ajustes pontuais, aplicados sobre o estado descrito nas Seções 8.1–8.5:

**1. Remoção da aba "Alertas"**
- Removidos do `index.html`: o botão da aba (`switchTab('alertas',...)` + badge `alertBadge`) e todo o bloco `<div id="alertasView">`.
- Removidos do `app.js`: a função `renderAlertas()` inteira, a chamada a ela em `renderAll()`, e as referências a `alertasView` em `switchTab()` (que agora só alterna entre `dashView` e `opView`).
- Removidas do `styles.css`: as classes exclusivas dessa aba (`.alert-count`, `.alert-row`, `.alert-icon`, `.alert-title`, `.alert-sub`), que ficaram órfãs.
- O quadro "Demandas pendentes com prazo vencido" continua existindo — só que agora **exclusivamente no Dashboard** (era duplicado entre Dashboard e Alertas antes).

**2. Aba Operacional — remoção da edição de data**
- A coluna "Data entrega" deixou de ser editável: `entCell` agora é sempre texto estático (`r.dataEntregaFch||'-'`), nunca mais vira `<input>`.
- **A edição de Observações foi preservada** (não foi pedido remover, só a de data). O botão "Ed"/"Salvar" (à direita da tabela) continua existindo, mas agora só controla a edição do campo Observações.
- `saveR()` foi simplificado: não lê mais nenhum input de data nem recalcula `situacaoFch` a partir de uma data digitada — só grava o valor de Observações em memória.
- Classe CSS `.inp-date`, agora órfã, foi removida do `styles.css`.

**3. Dashboard — coluna Observações no quadro de pendentes vencidas**
- Nova coluna "Observações" adicionada à direita da tabela `critTable` (mesmo quadro "Demandas pendentes com prazo vencido" do Dashboard), usando a mesma classe `.obs-cell` (quebra de linha, sem `undefined`/`null`) e a mesma função `escHtml()` já usadas na aba Operacional.
- `colspan` do estado vazio ("Nenhuma demanda...") atualizado de 7 para 8 para acompanhar a nova coluna.
- É somente leitura (igual às demais colunas desse quadro) — não há edição de Observações nesse card, só na aba Operacional.

**Testes realizados:**

| Teste | Resultado |
|---|---|
| `node -c app.js` (sintaxe) | ✅ sem erros |
| Nenhuma referência órfã a `alertas`/`alertBadge`/`renderAlertas` restante | ✅ confirmado por grep |
| Nenhuma referência órfã a `inp-date` restante | ✅ confirmado por grep |
| `saveR()` sem o input de data não quebra e continua salvando Observações | ✅ simulado |
| Linha do quadro de pendentes vencidas com Observações preenchida (com aspas/tags) — escapada corretamente | ✅ |
| Linha do quadro de pendentes vencidas com Observações vazia — célula vazia, sem "undefined" | ✅ |
| Diff completo contra a versão anterior — só as 3 mudanças pedidas, sem regressão em login/KPIs/gráficos/Operacional/perfCard | ✅ |

Teste visual no navegador (conferir que a aba Alertas sumiu do menu, que a coluna Data entrega não abre mais input, e que a nova coluna aparece corretamente no Dashboard) ainda precisa ser feito por Alessander.

## 9. Prompt de Retomada

```
Você está continuando um projeto a partir do handoff abaixo. Leia-o integralmente antes de responder.

[COLE AQUI O CONTEÚDO COMPLETO DESTE HANDOFF]

Contexto rápido: sou Alessander, da contabilidade da Efí S.A. Estou construindo um sistema web (dashboard HTML dark) para acompanhar o fechamento contábil de 5 empresas, lendo de uma planilha Excel. O sistema está hospedado no GitHub Pages, em DOIS repositórios mantidos em paralelo (user "Ctbefi"): "fechamento-contabil" (https://ctbefi.github.io/fechamento-contabil/) e "fechamento-contabil2" (https://ctbefi.github.io/fechamento-contabil2/), ambos protegidos por senha.

Os arquivos atuais do sistema são 4: index.html, styles.css, app.js e .nojekyll (te envio os que forem necessários).

REGRAS OBRIGATÓRIAS:
1. Quando me guiar por interface, dê UM passo de cada vez. Nunca dois comandos simultâneos.
2. Regra de negócio crítica: só "Em atraso" (prazo vencido + coluna R vazia) é demanda pendente. "Atrasado" significa concluída com atraso — NÃO é pendente.
3. Não proponha soluções que dependam de Azure AD App Registration — a TI não libera.
4. Seja honesto sobre limitações técnicas (CORS, segurança da senha no client-side, etc).
5. Ao editar o HTML/CSS/JS, se o arquivo ficar frágil, recrie do zero em vez de fazer muitas edições incrementais.

Me diga que leu o handoff e pergunte em que ponto estou para continuarmos.
```

---

*Handoff gerado para continuidade do projeto Sistema Web de Fechamento Contábil — Efí.*
