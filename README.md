# Projeto de Automação de Coleta de Dados com Selenium e Openpyxl


Este projeto foi desenvolvido para automatizar a coleta de informações de produtos em um site de ofertas. Utilizando Python, Selenium e Openpyxl, o script pesquisa por um produto específico e extrai dados como nome, valor, vendedor e data de publicação dos 5 primeiros resultados.

## Como Usar

1. **Pré-requisitos**

   Certifique-se de ter os seguintes componentes instalados:

   - Python (versão X.X.X)
   - Selenium WebDriver
   - Chrome WebDriver

2. **Instalação de Dependências**

   Execute o seguinte comando para instalar as bibliotecas necessárias:

   ```bash
   pip install selenium openpyxl
   ```

3. **Execução do Script**

   Execute o script `coleta_dados.py` para iniciar a automação. O script pesquisará o produto definido na variável `produto_item` e salvará os resultados em um arquivo Excel.

## Estrutura do Projeto

- `coleta_dados.py`: O script principal que automatiza a coleta de dados.
- `dados.xlsx`: O arquivo Excel onde os dados serão armazenados.
- `chromedriver.exe`: O driver necessário para o Selenium interagir com o navegador Chrome.

## Contribuindo

Sinta-se à vontade para abrir issues, propor melhorias ou contribuir com código para este projeto. Toda ajuda é bem-vinda!

1. Faça um fork do projeto
2. Crie um branch para sua feature (`git checkout -b feature/NomeDaFeature`)
3. Faça commit das mudanças (`git commit -m 'Adiciona nova feature'`)
4. Faça push para o branch (`git push origin feature/NomeDaFeature`)
5. Abra um pull request
