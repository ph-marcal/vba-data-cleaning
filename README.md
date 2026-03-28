# 📊 Padronização de Dados Contábeis com VBA (Alta Performance)

![VBA](https://img.shields.io/badge/Language-VBA-blue.svg)
![Excel](https://img.shields.io/badge/Tool-Microsoft%20Excel-green.svg)
![Data Analysis](https://img.shields.io/badge/Field-Data%20Audit-orange.svg)

## 📝 O Problema de Negócio
Em departamentos contábeis e de auditoria, a inconsistência nos nomes de fornecedores (espaços extras, acentos variados, letras maiúsculas/minúsculas misturadas) impede a conciliação automática de dados e gera erros em fórmulas como `VLOOKUP` (PROCV) e `SUMIFS` (SOMASE). 

Processar manualmente milhares de linhas é inviável e propenso a erros humanos.

## 💡 A Solução
Desenvolvi um script em **VBA Senior** que utiliza o conceito de **Arrays (processamento em memória)** para padronizar colunas inteiras de dados em segundos.

### Principais Funcionalidades:
- **Processamento em Memória:** Diferente de macros comuns que alteram célula por célula, este script carrega os dados para a memória RAM, processa e devolve tudo de uma vez, sendo até 50x mais rápido.
- **Sanitização Robusta:** 
  - Remove espaços duplos internos e espaços nas extremidades (`Application.Trim`).
  - Converte todo o texto para Maiúsculas (`UCase`).
  - **Remoção de Acentos:** Transforma "CONSTRUÇÃO" em "CONSTRUCAO", garantindo integridade na conciliação.
- **Tratamento de Erros:** O código ignora células com erro (`#N/D`, `#VALOR!`) e trata automaticamente colunas com apenas uma linha de dados.
- **Otimização de Recursos:** Desativa o motor de cálculo e atualização de tela do Excel durante a execução.

## 🚀 Performance
| Volume de Dados | Tempo Estimado (Tradicional) | Tempo com este Script |
|-----------------|------------------------------|-----------------------|
| 10.000 linhas   | ~15 segundos                 | < 1 segundo           |
| 100.000 linhas  | ~2 minutos                   | ~3 segundos           |

## 🛠️ Como utilizar
1. Abra o arquivo `.xlsm` disponível na pasta `bin/`.
2. Selecione qualquer célula da coluna que deseja padronizar.
3. Pressione `ALT + F8` e execute a macro `PadronizarColunaAtiva_Senior`.

## 📂 Estrutura do Repositório
- `/src`: Contém o código fonte em formato `.bas` (texto puro).
- `/bin`: Contém o arquivo Excel de exemplo para testes.

---
**Contato:**
[Paulo Marçal] - [in/pmarcal]
