# Chamada Mestre

PWA para realizar chamadas no iPhone usando a planilha XLSX da extensão Mestre.

O aplicativo publicado não contém a extensão Mestre nem planilhas com dados de
alunos. A planilha é selecionada pelo próprio usuário no aparelho ou no Drive.

## Regras preservadas

- A planilha mantém a estrutura e as datas originais.
- Datas repetidas continuam sendo datas iguais no Excel.
- O aplicativo mostra `1º horário`, `2º horário` apenas na tela.
- Vazio ou `0` significa presente.
- `F` significa falta.
- `I` significa ausência justificada/enfermidade.
- Ao salvar, somente as células da aula selecionada são alteradas no XML da primeira aba.

## Executar localmente

Na pasta do projeto:

```bash
python3 -m http.server 8080
```

Abra `http://localhost:8080`.

Para instalar no iPhone, publique a pasta em um endereço HTTPS e use
`Safari > Compartilhar > Adicionar à Tela de Início`.

## Google Drive

1. Crie um projeto no Google Cloud.
2. Ative a Google Drive API.
3. Crie um OAuth Client ID do tipo aplicativo Web.
4. Adicione o domínio HTTPS do aplicativo às origens JavaScript autorizadas.
5. No aplicativo, abra a engrenagem e informe:
   - OAuth Client ID;
   - ID do arquivo XLSX no Google Drive.

O aplicativo usa o escopo `drive` para abrir e atualizar o arquivo existente
informado pelo ID. Para uso pessoal, a conta deve ser adicionada como usuária
de teste na tela de consentimento OAuth enquanto o projeto estiver em teste.
