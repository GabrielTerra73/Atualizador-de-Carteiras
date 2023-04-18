
# Python
Software para união de tabelas

Atualizador de carteira
Este documento tem o proposito de auxiliar na preparação das planilhas do software “Atualizador”, onde seu principal objetivo é cruzar as informações das tabelas bases e retornar com o Status dos clientes.
Abaixo esta o print do modelo para o cabeçalho das planilhas que precisam de alteração para que o software funcione:

No ERP Sênior no diretório “” é possível exportar as informações que são necessárias para o funcionamento, o arquivo precisa ser salvo em Excel e necessita que seja salvo com o nome: “StatusCliente” onde abrindo o arquivo será necessário fazer as seguintes alterações:
Alterar o cabeçalho das colunas conforme o modelo abaixo:
Coluna ‘A’ = “Cliente”
Coluna ‘B’ = “Nome”
Coluna ‘N’ = “Situação”
Coluna ‘O’ = “Data Última Compra”
 
Seguindo exatamente este formato e sintaxe este arquivo já estará pronto para uso.
O segundo arquivo pode ser obtido no diretório do ERP “”, ao exportar é necessário ser salvo como “BaseCarteira”, abrindo o arquivo será necessário criar filtro para todas as colunas, com isso, na primeira coluna(A) selecione o filtro e marque apenas a opção “Cliente”, após aplicação, selecione tudo e copie, crie uma nova aba no relatório e cole a seleção, volte na primeira aba do relatório e selecione novamente a coluna(A), desta vez marcando a opção “Contato”, aplicando selecione tudo e copie, passando para a segunda aba, cole ao lado das informações do ‘Cliente”(necessário colar exatamente na mesma linha ao lado da primeira informação), agora é necessário acrescentar o cabeçalho das colunas, crie uma linha no topo da planilha e em seguida preencha com as informações(Não esqueça de apagar a primeira aba desta planilha ao concluir as alterações citadas):

Coluna ‘B’ = “ID”
Coluna ‘D’ = “Nome”
Coluna ‘I’ = “CNPJ/CPF”
Coluna ‘P’ = “Cod”
Coluna ‘Q’ = “Repr/Vend”
Destaco que o mais importe é estar conforme o modelo abaixo:
 
Seguindo exatamente este formato e sintaxe este arquivo já estará pronto para uso.
O terceiro arquivo pode ser obtido “”, este deve ser salvo como “Base Clientes” onde será necessário fazer alteração na formatação de dados da primeira coluna, ou coluna “Cliente”, para o tipo de dado “Geral” conforme o modelo abaixo:
 

  

Feito isso todos os arquivos estão prontos para execução do Atualizador


Por Gabriel Terra
