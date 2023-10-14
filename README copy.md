# Automatize

## Descrição
O projeto "Automatize" é uma ferramenta de automação que utiliza a biblioteca Selenium para interagir com páginas da web e automatizar tarefas repetitivas. Ele foi projetado para facilitar o preenchimento de formulários da web com base em dados de uma planilha do Excel, fornecendo uma maneira eficiente de enviar informações para um sistema web.

## Funcionalidades Principais
- Preenchimento automático de formulários web.
- Integração com planilhas do Excel para obter dados.
- Registro de chamados enviados com sucesso.
- Atualização de uma planilha de chamados com novos registros.

## Como Usar
Para utilizar o projeto "Automatize", siga as etapas abaixo:

1. Instale as bibliotecas necessárias listadas no arquivo `requirements.txt` usando o comando:

```Bash
pip install -r requirements.txt
```
2. Configure o projeto com os detalhes específicos da sua aplicação, como a URL do formulário web, mapeamento de clientes, mapeamento de observações, etc., no código-fonte.

3. Certifique-se de que o driver do Selenium para o navegador correto esteja instalado e configurado.

4. Execute o script Python para iniciar a automação. O script preencherá os formulários da web com base nos dados da planilha do Excel e registrará os chamados enviados com sucesso.

5. O histórico dos chamados enviados será salvo em um arquivo de texto com carimbo de data e hora.

## Contribuições
Contribuições para o projeto são bem-vindas. Sinta-se à vontade para abrir problemas (issues) e enviar solicitações de pull (pull requests) para melhorias, correções de bugs ou novos recursos.

## Autor
Este projeto foi desenvolvido por Charllys Fernandes.

## Licença
Este projeto está licenciado sob a MIT. Consulte o arquivo LICENSE para obter mais detalhes.