

Este script em Python realiza uma análise comparativa de preços entre dois mercados, Cerrano e Telentos, para uma lista de produtos alimentícios. Ele verifica qual mercado oferece os preços mais baixos e calcula a economia potencial para o usuário. Além disso, gera um arquivo Excel com os dados analisados e adiciona um gráfico de barras para visualização dos preços.

Instruções de Uso:

1. Certifique-se de ter o Python instalado em seu sistema. Você também precisará instalar as bibliotecas pandas e openpyxl. Você pode instalá-las usando o pip:
   
   ```
   pip install pandas openpyxl
   ```

2. Execute o script em um ambiente Python. Ele solicitará que você insira o valor disponível para compras.

3. O script comparará os preços dos produtos nos mercados Cerrano e Telentos e calculará a economia potencial ao optar pelo mercado mais barato.

4. Um arquivo Excel chamado `analise_precos.xlsx` será gerado, contendo os dados analisados e um gráfico de barras para visualização.

Observações:

- Certifique-se de ter os produtos e seus respectivos preços listados nos dicionários `cerrano` e `telentos`, São nomes de mercados fictício.
- Os preços devem ser fornecidos como strings, com a vírgula como separador decimal (por exemplo, "5,39").
- Este script utiliza as bibliotecas pandas e openpyxl para manipulação e exportação de dados em Excel. Certifique-se de tê-las instaladas no seu ambiente Python.

Para quaisquer dúvidas ou sugestões de melhoria, sinta-se à vontade para entrar em contato.
