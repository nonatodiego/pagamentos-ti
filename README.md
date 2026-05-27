# Mini sistema de controle de pagamentos (Excel)

Este repositório gera uma planilha `.xlsx` para controle de pagamentos de faturas, no formato solicitado.

## Funcionalidades

- Cadastro de faturas/contratos com campos:
  - Fornecedor, B.U, Serviço, Natureza, C/C, Dia de vencimento, Nº da fatura, Observações, Status
- Lançamentos mensais (Jan a Dez)
- Cálculo automático:
  - **Total Anual**
  - **Média Mensal (apenas meses com valor)**
- Fluxo para:
  - Contratos fixos
  - Faturas avulsas
- Alertas:
  - Vencimento nos próximos 7 dias
- Relatórios:
  - Pagos x Pendentes
  - Previsão do mês
- Importação de dados (aba `Importacao`)
- Exportação nativa para Excel (`.xlsx`)

## Como usar

1. Instale dependências:

```bash
python3 -m pip install -r requirements.txt
```

2. Gere a planilha:

```bash
python3 gerar_planilha.py
```

3. Abra o arquivo criado: `controle_pagamentos.xlsx`.

## Estrutura da planilha

- `Lancamentos`: base principal de controle
- `Dashboard`: indicadores resumidos
- `Importacao`: área para colar/importar dados
- `Listas`: apoio para validações

