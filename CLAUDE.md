# CLAUDE.md — App do Apartamento (Despesas, Contas fixas e Vencimentos)

Web App em **Google Apps Script** do Bruno para controlar as despesas do **apartamento dele**.
Uma planilha do Google é o banco de dados. Responda sempre em **português (pt-BR)**.

**Separado de propósito** do `app-lancamentos` (imóveis e carros dos pais do Bruno): planilha, projeto do
Apps Script e link próprios. Este app **nunca** lê nem grava nas planilhas do outro. Nada de carros aqui.

## Como o Bruno trabalha
- Tarefas com várias etapas: **um passo por vez**, só avançar quando ele disser **"OK"**.
- Depois de cada mudança no código, dizer **exatamente** o que colar e como publicar (ver "Publicar").
- Este repositório deve ficar igual ao que está publicado: ao terminar uma mudança, oferecer commit + push.

## Arquivos
- `Code.gs` — servidor: leitura/gravação, acessos, backup, alertas.
- `Index.html` — tela (HTML/CSS/JS num arquivo só). Módulos `D` (despesas), `F` (contas fixas, aba dentro de Despesas), `V` (vencimentos).
  Fora do Google usa o objeto `MOCK` (fim do arquivo) com dados de exemplo.
- Base: copiado do `app-lancamentos` (nov/2026), sem os módulos de abastecimento e manutenção.

## Onde está publicado
- Projeto do Apps Script **da planilha do apartamento** (Extensões → Apps Script), arquivos `Código.gs` e `Index`.
- App da Web: *Executar como: Eu* · *Quem pode acessar: Qualquer pessoa*. Acesso por link pessoal `?k=CHAVE`.
- Usuários: o Bruno e mais uma pessoa, cada um com seu link (aba `APP_ACESSOS`).

## Publicar uma alteração (instruções para o Bruno)
1. Colar `Code.gs` em `Código.gs` e/ou `Index.html` em `Index`; salvar (Ctrl+S).
2. Implantar → Gerenciar implantações → ✏ Editar → Versão: **Nova versão** → Implantar (o link não muda).
Funções novas que mexem em dados: sempre oferecer **simulação** antes e backup (`fazerBackup`).

## Planilha
- Nome: **Apartamento - Neori**. ID: `1fkBAKWoAUAMBrWXnEH0vpj_pYVEVVhquRmB1imf0xNM` (criada e publicada em 16/11/2026). O script é ligado a ela, então `ID_PLANILHA` no `Code.gs` pode ficar vazio.
- `configurar()` (rodar pelo editor) cria as abas que faltam e gera chave/link de quem está em `APP_ACESSOS`. Não apaga nada.

| Aba | Colunas |
|---|---|
| `DESPESAS` | A carimbo · B vencimento · C categoria · D valor · E descrição · F anexo · G data pagamento · H lançado por |
| `CONTAS FIXAS` | A ID (F0001) · B descrição · C categoria · D dia do vencimento · E valor previsto · F situação (ATIVA/INATIVA) · G criado por |
| `VENCIMENTOS` | A ID (V0001) · B criado em · C tipo · D descrição · E fornecedor · F início · G vence · H repetir (meses) · I anexo · J situação · K concluído em · L observação · M anterior · N criado por |
| `APP_ACESSOS` | A nome · B chave · C link (`=$E$2&"?k="&B2`) · E2 URL `/exec` |

- Todos em `APP_ACESSOS` veem e lançam tudo (não há restrição por local/categoria).
- **Sem campo Local** (é um apartamento só). Categorias iniciais (`CATS_DESP` no `Index.html`):
  MÓVEIS E DECORAÇÃO · ELETRO · UTILIDADES · OUTROS; outras são criadas no "+ nova".
- **Contas fixas** ficam dentro do módulo Despesas (aba "Contas fixas"): cadastro + "Lançar selecionadas" das que ainda
  não têm lançamento com a mesma descrição no mês. Lançar usa `despSalvar`. Em nov/2026 ainda não há nenhuma (mudança prevista para ~nov/2027).
- **Vencimentos**: sem carro e sem km. Tipos fixos Garantia e Seguro (`TIPOS_VENC`) + criados. Ao lançar despesa de
  MÓVEIS E DECORAÇÃO ou ELETRO, a tela oferece "Cadastrar a garantia".
- Alerta semanal (`ativarAlertas`, segunda 7h) e backup semanal (`ativarBackupSemanal`, domingo 3h) são opcionais.

## Regras de negócio (herdadas do app-lancamentos)
- **Parcelas**: uma linha por mês, sufixo `(1/3)` na descrição; centavos de arredondamento na última; só a 1ª recebe data de pagamento.
- **Pago x a pagar**: pago = data de pagamento preenchida. "A pagar" tem "Pago hoje" e 📅 (data escolhida).
- **Resumos** agrupam pela **data de pagamento** (só o que foi pago).
- **Vencimentos**: guarda só SITUAÇÃO = ATIVO/CONCLUÍDO/CANCELADO; o status (em dia/vence em breve/vencido) é calculado.
  Garantia não se renova. **Renovar** cria o próximo e marca o atual CONCLUÍDO.

## Armadilhas de Apps Script / Sheets
- Planilhas em **pt-BR**: `setFormula` com separador de argumentos deu `#ERROR!` — preferir fórmulas sem separador ou gravar valores.
- `google.script.run` não aceita `Date` como retorno: o servidor devolve datas como `'yyyy-MM-dd'`.
- Ao gravar datas use `Utilities.parseDate(iso, tz, 'yyyy-MM-dd')` com o fuso da planilha.
- Para checar data num valor de célula use `Object.prototype.toString.call(v) === '[object Date]'`.
- **Nunca escrever `/*` dentro do `<script>` do `Index.html`, nem em string** (ex.: `'image/*'`): publicado, travou a tela inteira.
  Em HTML (atributo `accept="image/*"`) não tem problema.
- **Nem `//` dentro de texto no `<script>`** (ex.: `'https://...'`): publicado, deu `SyntaxError: Invalid or unexpected token`
  e a tela não carregou (16/11/2026). Montar em partes (`'https:' + '/' + '/...'`, ver `URL_DRIVE`). Comentário `//` normal pode.
  Conferir antes de publicar: nenhuma linha do script com aspas seguidas de `//`.

## Testar antes de entregar
- Sintaxe do servidor: `node --check` numa cópia `.js` do `Code.gs`.
- Tela: `python -m http.server 8769` nesta pasta e abrir `/Index.html` (usa o `MOCK`); conferir no tamanho de celular.
  (No app do Claude, a configuração "apartamento" está no `.claude/launch.json` do `app-lancamentos`.)
- Funções de servidor que dependem do Google só rodam lá: pedir ao Bruno o print do **Registro de execução**.
