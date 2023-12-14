# Atualizador de carteira.
---

Este software foi desenvolvido com o proprosito de unir informações distintas entre planilhas, com o proposito de mesclar informações estrategicas da empresa, esta verçao disponibilizada foi desenvolvida com o proposito de compartilhar conhecimento, visando a confidencialidade dos dados originas, este codigo está sendo levemente alterado, contudo não afetara seu proposito original.

Este software está em constante uso e atualizações, está versão recebe atualizações periodicas, este codigo pode ser copiado apenas para fins educacionais.

Este documento tem o proposito de auxiliar na preparação das planilhas do software “Atualizador”, onde seu principal objetivo é cruzar as informações das tabelas bases e retornar com o Status dos clientes, baseado na data de sua ultima compra.
Como não há os arquivos bases este readme também recebera alterações para auxiliar na compreenção.

Abaixo esta o modelo para o cabeçalho das planilhas que precisam de alteração para que o software funcione:(imagem)

No ERP ou SGBD, após selecionar a empresa e em seguida busque pelo relatório(planilha), salve com o nome desejado e subistitua onde estiver escrito `Base Clientes` neste codigo, sem seguida identifique o campo do (`ID`) e `Data do cadastro` do cliente salve o nome da coluna para realizar a alteração no codigo, em seguida busue pelo retatorio de data da ultima compra dos clientes, gere o arquivo, identidique o campo do (`ID`) do cliente e a `data da ultima compra`, o qual neste codigo a coluna estará salva com este nome e salve em excel, necessario que seja salvo um nome intuitivo para a filial se for o caso, após isso refaça a operação para as demais filiais guarde o nome dos arquivos para alteração no codigo subistua onde estiver escrito "StatusCFM" para Matriz e "StatusCF2" para as demais filias, este cogigo foi construido inicialmente para 5 filiais, contudo facilmente pode ser aumentado para mais filiais.

O arquivo relação 'Cliente por vendedor' apos ser obtido no seu ERP ou SGBD, salve preferencialmente como `Base Carteira` como esta no modelo, apois isso precisara que seja identificao a coluna de `ID` do cliente o qual neste codigo esta escrito como `Cliente` e a coluna do seu vendedor o qual esta escrito como `Cod` neste software. apos esses processos os arquivos estão prontos para execução do software.
  

Feito isso execute o **Atualizador**


Por **Gabriel Terra**
