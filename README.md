# Dashboard de Vendas Financeiro - Serviço de Assinaturas

## 📊 Visão Geral do Projeto
Este projeto consiste em um Dashboard de Vendas desenvolvido em Excel, focado em indicadores financeiros (KPIs) para um serviço de assinatura de jogos (semelhante ao Game Pass). O objetivo é transformar dados brutos de assinantes em insights visuais claros sobre faturamento, perfil de consumo e eficiência de descontos.

## 📁 Estrutura do Repositório
- `Dashboard_Vendas.xlsx`: O arquivo Excel final contendo os dados processados e o Dashboard visual.
- `generate_dashboard.py`: Script Python utilizado para processar os dados e automatizar a geração do dashboard e dos gráficos.
- `README.md`: Este arquivo com as informações do projeto.

## 📈 Indicadores (KPIs) e Visualizações
O dashboard destaca os seguintes pontos:
1. **KPIs de Cartão**:
   - **Receita Total**: Faturamento bruto consolidado.
   - **Total de Inscritos**: Volume da base de clientes.
   - **Ticket Médio (ARPU)**: Receita média gerada por usuário.
   - **Total de Descontos**: Impacto financeiro dos cupons aplicados.
2. **Gráficos**:
   - **Distribuição de Receita por Plano**: Comparativo entre planos Ultimate, Core e Standard.
   - **Share de Receita por Tipo**: Distribuição percentual entre pagamentos Mensais, Trimestrais e Anuais.

## 🛠️ Tecnologias Utilizadas
- **Excel**: Para a visualização final e dashboard.
- **Python (Pandas & XlsxWriter)**: Para processamento de dados e automação da estrutura do Excel.

## 🚀 Como Reproduzir
Caso deseje gerar o dashboard novamente ou aplicar em novos dados:
1. Certifique-se de ter o Python instalado.
2. Instale as dependências: `pip install pandas xlsxwriter openpyxl`.
3. Execute o script: `python generate_dashboard.py`.
4. O arquivo `Dashboard_Vendas.xlsx` será gerado automaticamente.

---
*Este projeto foi desenvolvido como parte de um desafio de criação de Dashboard de Vendas.*
