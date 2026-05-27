# Jornadas sem segundos

Projeto web estático para tratar uma planilha mensal de jornada de motoristas. Ele lê uma planilha Excel, localiza a aba `Jornadas` e remove os segundos de células que contenham horários ou durações, mantendo a estrutura da planilha.

## Como abrir

Abra o arquivo `index.html` diretamente no navegador. Não é necessário backend, servidor local, React, Vue, Angular ou instalação de pacotes.

O projeto usa bibliotecas via CDN:

- ExcelJS para ler e gerar arquivos `.xlsx` preservando melhor abas, estilos básicos, larguras e células mescladas.
- SheetJS como fallback e para leitura de arquivos `.xls`.

## Como usar

1. Abra `index.html`.
2. Arraste a planilha para a área de upload ou clique para selecionar um arquivo `.xlsx` ou `.xls`.
3. Confira o nome do arquivo carregado.
4. Clique em `Processar planilha`.
5. Veja o resumo, os alertas e a prévia das primeiras 20 linhas.
6. Clique em `Baixar planilha final`.

O arquivo final será gerado com o nome:

```text
Relatorio_Jornadas_Sem_Segundos.xlsx
```

## O que é alterado

Somente horários e durações com segundos na aba processada.

Exemplos:

```text
122:36:34 -> 122:36
07:02:41  -> 07:02
08:23:22  -> 08:23
00:00:00  -> 00:00
197:43:51 -> 197:43
5:04:09   -> 5:04
```

Também são tratados valores numéricos do Excel quando a célula possui formatação de hora/duração ou quando a coluna tem nome relacionado a jornada, tempo, hora, serviço, direção, descanso, refeição, espera ou interjornada.

## O que não é alterado

O sistema não remove colunas, não remove linhas, não reorganiza motoristas e não cria cálculos novos.

São mantidos:

- Todas as abas da planilha original.
- Todas as colunas da aba processada.
- Todas as linhas.
- Textos comuns.
- Nomes de motoristas.
- CPF.
- Código de RH.
- Títulos e cabeçalhos.
- Período.
- Datas comuns.
- Valores monetários.
- Números comuns fora de colunas de tempo.

## Aba processada

O sistema procura uma aba chamada `Jornadas`.

Se encontrar, processa essa aba.

Se não encontrar, processa a primeira aba da planilha e mostra um alerta informando que a aba `Jornadas` não foi localizada.

## Alertas

A tela pode exibir alertas quando:

- A aba `Jornadas` não for encontrada.
- Nenhum horário com segundos for encontrado.
- Alguma célula parecer horário, mas estiver em formato inválido.
- O arquivo não puder ser lido.
- Um arquivo `.xls` for usado e alguns estilos não puderem ser preservados integralmente.

Esses alertas não impedem o download quando a planilha foi processada com sucesso.

## Regra principal

A aba processada deve continuar igual à planilha original, com todas as colunas, linhas e dados. A única mudança permitida é transformar horários ou durações com segundos em horários ou durações sem segundos.
