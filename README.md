# Automatização Python WebScrapping
WebScrapping do preço do Alumínio e do Dólar da LME (London Metal Exchange), via site da ShockMetais

# Funcionamento
O script utiliza principalmente Selenium para acessar o site da ShockMetais, buscar a tabela em que estão os preços dos insumos, procura o dia atual e extrai as informações relativas ao preço do Alumínio e do Dólar, e depois converte para Reais, gerando um arquivo em Excel com tais informações. Ele foi configurado para, todos os dias, salvar os históricos anteriores no mesmo arquivo (para que se conserve o histórico) e atualizar o do dia atual na linha abaixo. Ele também envia por email todos os dias, e utilizei o Schedule para essa função. Há também a extração da variação semanal do preço desses insumos, que também é enviado semanalmente em um determinado horário por email para os emails que defini.
