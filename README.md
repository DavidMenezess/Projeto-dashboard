# API — Dashboard Directus Advocacia e Consultoria

API que alimenta o dashboard com os dados de produção e processos parados,
com login autenticado. Esta é a primeira etapa do projeto: roda com um
arquivo Excel local. A sincronização automática com o OneDrive é a
próxima etapa (veja "Próximos passos" no final).

## O que este projeto faz

- Login com e-mail/senha, protegido por token JWT
- Duas rotas de dados, que só respondem para quem estiver logado:
  - `GET /api/tarefas` — dados da seção "Produção 2026"
  - `GET /api/processos-parados` — dados da seção "Processos Parados"
- Documentação automática da API em `/docs` (gerada pelo FastAPI)

## Estrutura do projeto

```
projeto-dashboard/
├── app/
│   ├── main.py            # rotas da API
│   ├── auth.py             # login, hash de senha, token JWT
│   ├── config.py            # leitura das variáveis de ambiente
│   ├── database.py          # banco de dados de usuários (SQLite)
│   ├── data_processor.py    # transforma a planilha em dados do dashboard
│   └── cache.py             # guarda em memória o último dado processado
├── scripts/
│   └── criar_usuario.py     # cria/atualiza um login (linha de comando)
├── data/                    # planilhas e banco SQLite ficam aqui (não versionado)
├── Dockerfile
├── docker-compose.yml
├── requirements.txt
├── .env.example              # copie para .env e preencha
└── README.md
```

## Como rodar localmente (sem Docker)

Pré-requisito: Python 3.12+.

```bash
# 1. Ambiente virtual e dependências
python3 -m venv venv
source venv/bin/activate          # Windows: venv\Scripts\activate
pip install -r requirements.txt

# 2. Configuração
cp .env.example .env
# Abra o .env e gere uma chave secreta com:
python -c "import secrets; print(secrets.token_hex(32))"
# Cole o resultado em SECRET_KEY dentro do .env

# 3. Coloque as planilhas na pasta data/
#    (mesmo nome configurado no .env: CAMINHO_PLANILHA_TAREFAS e
#    CAMINHO_PLANILHA_PROCESSOS_PARADOS)

# 4. Crie o primeiro usuário (responda "admin" quando perguntar o tipo de acesso)
python -m scripts.criar_usuario

# 5. Suba a API
uvicorn app.main:app --reload
```

Acesse `http://localhost:8000/docs` para ver e testar todas as rotas
pelo navegador, sem precisar de Postman ou curl.

## Como rodar com Docker (igual à produção)

```bash
cp .env.example .env   # preencha como no passo acima
docker compose up --build
```

Para criar o usuário dentro do container:

```bash
docker compose exec api python -m scripts.criar_usuario
```

## Testando o login pela linha de comando

```bash
# 1. Fazer login e guardar o token
TOKEN=$(curl -s -X POST http://localhost:8000/auth/login \
  -d "username=seu@email.com&password=suasenha" \
  | python3 -c "import sys,json; print(json.load(sys.stdin)['access_token'])")

# 2. Usar o token para pegar os dados
curl http://localhost:8000/api/tarefas -H "Authorization: Bearer $TOKEN"
```

## Usuários e administradores

O login é feito com **Nome + Sobrenome** (não e-mail) — mais fácil de
lembrar e digitar no dia a dia. Internamente, isso vira uma "chave de
login" única (ex: David + Menezes → `david.menezes`), sem acento, sem
espaço, tudo minúsculo — então "DAVID" + "menezes" funciona igual a
"David" + "Menezes".

**Atenção:** se duas pessoas tiverem exatamente o mesmo nome e sobrenome,
a segunda pessoa precisa ser cadastrada com um sobrenome mais completo
(ex: "Silva Neto" em vez de só "Silva") para não colidir com a primeira.

Existem dois tipos de acesso:

- **Normal**: loga e vê o dashboard normalmente.
- **Administrador**: além de ver o dashboard, tem acesso à aba **"Usuários"**,
  onde consegue ver todos os usuários cadastrados (login, cargo, tipo,
  quando foi criado, último login, e se está com sessão ativa agora) e
  cadastrar novas pessoas — sem precisar mexer em terminal ou no servidor.

O **primeiro** administrador precisa ser criado pelo script de linha de
comando (`scripts/criar_usuario.py`), porque a rota de cadastro pela API
exige que já exista um admin logado. Depois desse primeiro, os próximos
usuários (admin ou normais) podem ser criados direto pela tela.

O campo "sessão ativa" (aparece como uma etiqueta "online" na tabela) é uma
aproximação: considera a pessoa "online" se o último login dela foi dentro
da janela de validade do token (`ACCESS_TOKEN_EXPIRE_MINUTES`, 60 minutos
por padrão). Não é uma contagem em tempo real — para isso seria necessário
um mecanismo à parte (WebSocket), que não foi incluído por ora para manter
o projeto simples.

## Aba "Apresentação"

Modo pensado para exibir em TV ou projetor numa reunião — em vez dos
painéis normais (feitos pra tela de computador, com bastante informação
lado a lado), essa aba mostra os mesmos dados em formato de slide: um
destaque por tela, letras bem grandes, fácil de ler à distância.

- Navega com os botões de seta, os pontinhos embaixo, ou as setas do
  teclado (←/→) quando essa aba está aberta.
- O botão "Tela cheia" usa o modo de tela cheia do navegador — ideal na
  hora de projetar.
- Os slides são montados a partir dos mesmos dados das outras abas (não
  busca nada novo na API) e acompanham o ano selecionado no seletor de
  "Produção".
- São 15 slides prontos por padrão (capa, números do ano, situação
  detalhada, evolução mensal com variação mês a mês, quem mais produziu,
  qualidade por responsável, tipos de tarefa, processual x administrativa,
  tempo de resposta, polo no processo, área do direito, delegação,
  processos parados recentes e crônicos, e um resumo final) — cada um com
  números reais ao lado do gráfico e frases explicando o que aquele dado
  significa, não só o gráfico "pelado".
- Botão **"Montar minha apresentação"**: escolhe quais desses slides
  entram e em que ordem, com um clique em "Salvar". Fica guardado no
  navegador — abrir de novo já carrega a versão personalizada. O botão
  "Usar apresentação completa" volta pra versão com tudo, a qualquer hora.
- Botão **"Criar slide do zero"**: monta um slide totalmente seu, escolhendo
  livremente entre qualquer número/estatística do dashboard (total de
  tarefas, taxa de sucesso, tarefas de uma pessoa específica, uma situação
  específica, etc.), qualquer gráfico já existente, e/ou um bloco de texto
  livre escrito na hora — quantos blocos quiser, na ordem e no tamanho
  (pequeno/médio/grande) que escolher, com pré-visualização ao vivo antes
  de salvar. O slide salvo aparece automaticamente na lista do "Montar
  minha apresentação", com opção de editar ou excluir depois. Também fica
  guardado no navegador.

**Nota sobre "clientes atendidos"**: esse número é calculado a partir de
quem aparece marcado como cliente do escritório nas partes do processo
(ativas e passivas) de cada tarefa — não da coluna "Envolvidos do
atendimento", que só é preenchida em tarefas administrativas (uma fração
pequena da planilha) e sozinha subestimava bastante o número real.

## Aba "Indicadores"

Análises mais profundas, calculadas a partir de colunas da planilha que
antes não eram usadas — mas só as que têm dado suficiente pra serem
confiáveis (ver checagem de preenchimento abaixo):

- **Tempo médio de ciclo**: quantos dias, em média, uma tarefa leva da
  criação até a conclusão — geral e por responsável.
- **Polo no processo (produção geral)**: mesma classificação autor/réu que
  já existia em "Processos Parados", agora aplicada a todas as tarefas do
  ano, não só às paradas.
- **Área do direito mais frequente**: extraída do campo "Assunto" da
  planilha.
- **Quem mais cria tarefas / Autoatribuídas x Delegadas**: mostra se quem
  cria a tarefa é quem a executa, ou se ela é repassada para outra pessoa.

**Campos que ficaram de fora por enquanto**: horas trabalhadas (timesheet),
instância, fase e vara do processo. Na planilha atual, esses campos estão
0-16% preenchidos — não davam um gráfico confiável. Se o escritório passar
a preencher esses dados no Projuris, `app/data_processor.py` é o lugar
certo para adicioná-los (a função `_indicadores_do_ano` já está pronta para
receber mais campos).

## Processos Parados: recentes x crônicos

Antes, só existia um grupo ("90 a 120 dias"). Um processo parado há 300 ou
400 dias simplesmente não aparecia em lugar nenhum do dashboard. Agora a
aba mostra dois grupos separados:

- **Recentes** — entre 90 e 120 dias sem movimentação (os que acabaram de
  cruzar o limite).
- **Crônicos** — mais de 120 dias (o grupo que estava escondido antes).

Na planilha usada para os testes, isso revelou uma diferença enorme: 64
processos no grupo "recentes" contra quase 2 mil no grupo "crônicos" — a
grande maioria dos processos parados do escritório é, na real, crônica.

## Histórico de anos e comparativo

A planilha de tarefas cresce com o tempo — em 2027 ela vai ter linhas de
2026 e 2027 juntas, em 2028 os três anos, e assim por diante (o Projuris
exporta tudo, não só o ano corrente). O dashboard já foi pensado pra isso:

- A API separa os dados por ano automaticamente, sem precisar de nenhuma
  configuração — ela olha a coluna "Data prevista" de cada tarefa.
- A aba "Produção" tem um seletor no topo ("Exibindo: Produção 2026 ▾")
  que troca entre qualquer ano encontrado na planilha, mostrando sempre o
  ano mais recente por padrão ao abrir.
- O botão "Comparar anos/meses" abre um comparativo lado a lado entre dois
  anos (ou o mesmo ano, dois recortes de meses diferentes) — útil pra
  perguntas do tipo "como estamos em relação ao ano passado nesse mesmo
  período".
- As pendências (prazo vencido / a vencer) **não** mudam com o seletor de
  ano — elas mostram sempre a situação atual, então uma tarefa vencida de
  2026 continua aparecendo ali em 2027 até ser resolvida.

Nenhuma configuração é necessária para isso — assim que a planilha tiver
linhas de um novo ano, ele aparece sozinho no seletor.

## Testando com planilhas antigas (data de referência fixa)

A API sempre calcula prazos e o filtro de "90 a 120 dias parado" com base na
data real de hoje — é o comportamento correto em produção. Isso significa
que, se você testar localmente com as planilhas de exemplo (de julho de
2026), a seção "Processos Parados" pode aparecer vazia depois de algum
tempo, porque a amostra "envelhece" e sai da janela de 90–120 dias.

Para testar sem esse efeito, defina no `.env`:

```
DATA_REFERENCIA_TESTE=30/07/2026
```

Isso faz a API calcular tudo como se "hoje" fosse essa data. **Deixe essa
variável em branco em produção** — sem ela, a API sempre usa a data real,
que é o que você quer quando a planilha estiver sendo atualizada de verdade.

## Segurança e LGPD — decisões tomadas e por quê

- **Senhas**: nunca são guardadas em texto puro, apenas o hash gerado com
  `bcrypt` (algoritmo lento de propósito, dificulta ataque de força bruta).
- **Login sem pista**: uma senha ou e-mail errados devolvem sempre a mesma
  mensagem genérica, para não revelar se um e-mail existe no sistema.
- **Sem cadastro público**: os logins só são criados por quem tem acesso
  ao servidor (`scripts/criar_usuario.py`), reduzindo a superfície de ataque.
- **Minimização de dados**: os dados processuais (tarefas, processos) não
  ficam guardados permanentemente em banco — são recalculados a partir da
  planilha e mantidos em memória. Menos um lugar com dado sensível parado.
- **CORS restrito**: só os endereços listados em `ORIGENS_PERMITIDAS`
  podem consumir a API a partir de um navegador.
- **Logs de acesso**: todo login (bem ou malsucedido) é registrado com
  data e hora — importante para auditoria, já que há dados processuais
  de terceiros envolvidos.
- **HTTPS**: obrigatório em produção. O Render (hospedagem recomendada)
  fornece certificado grátis automaticamente.

Pontos que ainda precisam de decisão do escritório (não são técnicos,
são de política interna, e valem uma conversa com quem cuida do
compliance/LGPD):
- Por quanto tempo os logs de acesso devem ser guardados.
- Quem, dentro do escritório, deve ter uma conta no dashboard.
- Qual a base legal declarada para o tratamento dos dados processuais
  exibidos (sugestão: legítimo interesse / execução de contrato).

## O dashboard (frontend)

O arquivo `frontend/dashboard.html` é o mesmo painel visual de antes, agora com
tela de login e busca de dados na API em vez de vir tudo fixo dentro do arquivo.

### Como rodar localmente

O dashboard precisa ser aberto por um servidor web (não funciona clicando
duas vezes no arquivo — o navegador bloqueia por segurança). Com a API já
rodando (veja acima), em outro terminal:

```bash
cd frontend
python3 -m http.server 5500
```

Acesse `http://localhost:5500/dashboard.html`. Faça login com o usuário
criado em `scripts/criar_usuario.py`.

### Ajustando o endereço da API

No topo do bloco `<script>` final do `dashboard.html`, existe esta linha:

```javascript
const API_BASE_URL = 'http://localhost:8000';
```

Troque para o endereço real quando a API for publicada (ex:
`https://api.seudominio.com.br`).

### CORS

A API só aceita chamadas de navegador vindas dos endereços listados em
`ORIGENS_PERMITIDAS` no `.env`. Para testar localmente, inclua o endereço
onde o dashboard está sendo servido (ex: `http://localhost:5500`).

## Próximos passos (ainda não implementados)

1. **Sincronização com OneDrive**: substituir a leitura do arquivo local
   por um job agendado que baixa o Excel mais recente via Microsoft Graph
   API antes de chamar `data_processor.py` — o processamento em si não
   muda nada.
2. **Deploy no Render**: conectar este repositório, configurar as
   variáveis de ambiente do `.env` no painel do Render, e apontar o
   domínio comprado para o endereço gerado. O `frontend/dashboard.html`
   também precisa de um lugar para ficar hospedado com HTTPS (pode ser o
   próprio Render como um segundo serviço, ou GitHub Pages/Netlify).
