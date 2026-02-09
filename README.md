# 📊 Apurador de Investimentos

Aplicação web para apuração de investimentos em promoções de produtos.

## 🚀 Funcionalidades

- Upload de planilha de Preço Final
- Upload de múltiplas planilhas de Orçamento
- Cálculo automático de investimentos e valores de pedido
- Geração de relatório Excel formatado com:
  - Resumo geral com totais
  - Cores personalizadas
  - Formatação de moeda (R$) e percentual (%)
  - Análise por orçamento

## 📋 Como Usar

1. **Carregue a Planilha de Preço Final**
   - Arquivo Excel com colunas: EAN/COD BARRAS e Valor Negociado

2. **Informe o Nome da Rede**
   - Digite o nome da rede para identificação no relatório

3. **Carregue as Planilhas de Orçamento**
   - Arquivos Excel com cabeçalhos na linha 8
   - Colunas obrigatórias: EAN, VALOR SKU PAGO, QUANTIDADE

4. **Processar Dados**
   - Clique em "Processar Dados" para gerar a análise

5. **Baixar Resultado**
   - Faça download do arquivo Excel com a apuração completa

## 🛠️ Tecnologias

- Python 3.9+
- Streamlit
- Pandas
- OpenPyXL

## 📦 Instalação Local

```bash
pip install -r requirements.txt
streamlit run app_apurador.py
```

## 📄 Licença

Uso interno - Projeto Nivea
