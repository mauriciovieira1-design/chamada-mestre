# Chamada Mestre

PWA para realizar chamadas no iPhone usando a planilha XLSX da extensão Mestre.

O aplicativo publicado não contém a extensão Mestre nem planilhas com dados de
alunos. A planilha é selecionada pelo próprio usuário no aparelho ou no Drive.

## Regras preservadas

- A planilha mantém a estrutura e as datas originais.
- Datas repetidas continuam sendo datas iguais no Excel.
- O aplicativo mostra uma única chamada por data.
- Ao salvar, a mesma frequência é gravada em todos os horários repetidos daquela data.
- Vazio ou `0` significa presente.
- `F` significa falta.
- `I` significa ausência justificada/enfermidade.
- Ao salvar, somente as células da aula selecionada são alteradas no XML da primeira aba.
- Antes de enviar, o aplicativo baixa a versão mais recente do Drive e reaplica
  somente as chamadas pendentes.
- Se a mesma turma/data ou a lista de alunos mudou no Drive, o envio é bloqueado
  para impedir a perda de chamadas.

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
2. Ative a Google Drive API e a Google Picker API.
3. Crie um OAuth Client ID do tipo aplicativo Web.
4. Adicione o domínio HTTPS do aplicativo às origens JavaScript autorizadas.
5. Crie uma chave de API restrita ao domínio do aplicativo e à Google Picker API.
6. No aplicativo, abra a engrenagem e informe:
   - OAuth Client ID;
   - chave de API;
   - número do projeto Google Cloud.
7. Toque em `Escolher planilha no Google Drive` e selecione o arquivo XLSX.

O aplicativo usa o escopo limitado `drive.file`. Ele pode abrir e atualizar
somente a planilha escolhida pelo usuário no seletor oficial do Google. Para
uso pessoal, a conta deve ser adicionada como usuária de teste na tela de
consentimento OAuth enquanto o projeto estiver em teste.

A planilha escolhida fica memorizada no aparelho. Uma cópia privada também é
mantida no armazenamento local do PWA, permitindo abrir e fazer chamadas sem
selecionar o arquivo novamente. Quando o acesso temporário do Google expirar,
o aplicativo solicitará uma nova autorização somente ao sincronizar.

Uma chamada feita sem conexão fica armazenada como uma operação pendente
(turma, data e alunos), e não como um arquivo inteiro pronto para substituir o
Drive. Na sincronização, ela é aplicada sobre a versão mais recente da planilha.
