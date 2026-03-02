# MSGtoOutlook

Uma aplicação macOS que corre silenciosamente no background, convertendo ficheiros `.msg` para `.eml` usando `extract-msg` e abrindo-os no Microsoft Outlook de forma automática.

## Requisitos Iniciais
- Python 3 instalado no macOS.
- Xcode Command Line Tools (o `py2app` precisa disto para compilar). Nas versões recentes do macOS, pode simplesmente abrir o terminal e correr `xcode-select --install`.

## Compilação usando o GitHub Actions (Recomendado se não tiver um Mac à mão)

Criámos um *workflow* para o GitHub Actions que tratará de compilar tudo num Mac na nuvem!

1. Crie um repositório vazio no seu GitHub.
2. Coloque toda esta pasta (incluindo a pasta `.github` que acabámos de criar) nesse repositório via `git push` ou submetendo os ficheiros na interface web.
3. Vá ao separador **Actions** no seu repositório do GitHub. O workflow irá correr automaticamente caso tenha feito commit para o `main` ou `master`. 
4. Se quiser, também pode clicar em "Build macOS App" no separador Actions e de seguida em "Run workflow" para correr manualmente.
5. Quando o processo terminar, no fundo da página inicial do Workflow aparecerá um **Artifact** chamado `MSGtoOutlook-macOS`.
6. Faça download do ficheiro ZIP, descomprima-o e instale a sua App no seu Mac.

## Compilação Local (Se tiver o seu Mac e Terminal abertos)

1. Abra o Terminal e navegue até à pasta onde estão guardados estes ficheiros.

2. Instale as dependências executando:
   ```bash
   pip3 install -r requirements.txt py2app
   ```

3. Geração do executável (App do macOS):
   ```bash
   python3 setup.py py2app
   ```

4. Após a geração concluir com sucesso, será criada uma pasta chamada `dist`. Lá dentro, encontrará a aplicação `MSGtoOutlook.app` (ou apenas `MSGtoOutlook`).  
   **Recomendação importantíssima**: Mova o ficheiro `MSGtoOutlook.app` da pasta `dist` para a pasta `/Aplicações/` do seu Mac.

## Como Associar Ficheiros .msg no Mac à nova App

Agora que tem a sua aplicação "MSGtoOutlook" compilada e na pasta de Applicaçōes, queremos que o Mac abra todos os ficheiros `.msg` sempre lá:

1. Clique com o botão direito (ou Control+Clique) em cima de **qualquer ficheiro `.msg`** que tenha no seu Mac.
2. Escolha **"Obter Informações"** (Get Info).
3. Na janela que abre, procure a secção **"Abrir com:"** (Open with:).
4. No menu suspenso (dropdown menu), escolha a sua recém-criada aplicação **`MSGtoOutlook`**.
   *(Se não aparecer logo, clique em "Outra...", vá à pasta Aplicações e selecione-a)*
5. Assim que a app estiver selecionada, clique no botão logo abaixo: **"Alterar tudo..."** (Change All...).
6. Confirme que quer alterar, e pronto!

## O que vai acontecer na Prática

Ao dar dois cliques (ou um "Enter" selecionando) qualquer ficheiro `.msg` no seu Mac:
1. O macOS abre *apenas virtualmente e no background* (não aparece nenhuma doca ou janela preta) a app `MSGtoOutlook`.
2. A conversão e extração de dados + anexos acontece num milissegundo.
3. É gerado o `.eml` convertido de forma temporária na sua pasta oculta `/tmp/`.
4. O Microsoft Outlook para o Mac é aberto já com o e-mail completo, formatado e fidedigno à sua frente no ecrã.
