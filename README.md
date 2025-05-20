# lei-perimetro-urbano

Dado a entrada de uma tabela, do tipo:

|    | OBJECTID_1 | OBJECTID  |Shape_Leng|         Nome|  Direction| Distance|  ORIG_FID|             N|            E|
| :------- | :------- | :------- | :------- | :------- | :------- | :------- | :------- | :------- | :------- |
|0   |           1|         7 | 106.137654|  Bairro Alto|   92-44-51|  106,138|         0|  7.206350e+06|  726831.2672|
|1   |           2|         8 | 42.093888|  Bairro Alto|  156-13-49|   42,094|         1|  7.206344e+06|  726937.2828|
|2   |           3|         9 | 49.615773|  Bairro Alto|   85-51-24|   49,616|         2|  7.206306e+06|  726954.2492|


Cria-se uma descrição para a lei com uma saida do tipo: 

"...vértice P00,de coordenadas N 7.346.174,33000 m e E 360.535,67000 m, que segue confrontando por linha seca em um azimute de 91°44'22" a uma distância de 1.492,74 m até o vértice P01...."
