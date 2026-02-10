# 📊 Apurador de Investimentos

Aplicação web para apuração de investimentos em promoções de produtos.

## 🚀 Funcionalidades

- 📋 Download de planilha modelo para Preço Final
- Upload de planilha de Preço Final
- Upload de múltiplas planilhas de Orçamento
- Cálculo automático de investimentos e valores de pedido
- Geração de relatório Excel formatado com:
  - Resumo geral com totais
  - Cores personalizadas
  - Formatação de moeda (R$) e percentual (%)
  - Análise por orçamento

## 📋 Como Usar

1. **Baixe a Planilha Modelo (Opcional)**
   - Clique no botão "📋 Download Planilha Modelo" no topo da página
   - Use como referência para o formato esperado de Preço Final

2. **Carregue a Planilha de Preço Final**
   - Arquivo Excel com colunas: EAN/COD BARRAS e Valor Negociado

3. **Informe o Nome da Rede**
   - Digite o nome da rede para identificação no relatório

4. **Carregue as Planilhas de Orçamento**
   - Arquivos Excel com cabeçalhos na linha 8
   - Colunas obrigatórias: EAN, VALOR SKU PAGO, QUANTIDADE

5. **Processar Dados**
   - Clique em "Processar Dados" para gerar a análise

6. **Baixar Resultado**
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
