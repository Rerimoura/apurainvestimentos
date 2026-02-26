# Apurador de Investimentos Reppox

Aplicação web para apuração de investimentos para pedidos feitos no site Reppos.

## Funcionalidades

- Download de planilha modelo para preencher o preço da rede
- Upload de múltiplas planilhas de orçamento reppos
- Cálculo automático de investimentos e valores de pedido
- Geração de relatório Excel formatado com:
  - Resumo geral com totais
  - Cores personalizadas
  - Formatação de moeda (R$) e percentual (%)
  - Análise por orçamento

## Como Usar

1. **Baixe a planilha modelo (opcional)**
   - Clique no botão "📋 Download Planilha Modelo", de acordo com a indústria, no topo da página.
   - Preencha a planilha com os preços do cliente.

2. **Carregue a planilha de preço final**
   - Arquivo Excel com colunas: EAN/Cod barras e valor negociado

3. **Informe o nome da rede**
   - Digite o nome da rede para identificação no relatório

4. **Carregue as planilhas de orçamento Reppos**
   - Arquivos Excel obtido ao exportar o carrinho no site Reppos

5. **Processar Dados**
   - Clique em "processar dados" para gerar a planilha final de investimentos

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
