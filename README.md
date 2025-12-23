**Script: gerar_planilhas_por_media_source.py**

**O script gera uma duas pasta: uma contendo todos os media sources extraídos da planilha bruta que estejam no app analisado e que possuam media source e publisher correspondentes na planilha de "pid e media source".
A segunda pasta contem arquivos zips de todos os publishers, que pertencam ao app da planilha bruta, contendo as planilhas com os dados dos media sources correspondentes, extraidos da planilha bruta.**
1. Pega o nome do aplicativo na coluna App Name da planilha bruta
2. Verifica a coluna App da planilha "Pid e Publisher" e seleciona os publishers e seus media sources correspondentes
3. Extrai as informações dos media sourcers dos publishers selecionados na planilha bruta 
4. Gera uma pasta contendo as planilhas com as informações extraidas de cada media source do App
5. Gera uma pasta contendo os arquivos zips de cada publisher com as planilhas extraidas correspondentes   

==================================================

**Pré-requesitos:**

1. Instalar o python. **Link para o download:** https://www.python.org/downloads/

2. Instalar a extensão do python no vs code

<img width="250" height="250" alt="image" src="https://github.com/user-attachments/assets/ea3f83e0-ac73-4f2c-ab14-f17312d4107c" />

3. Instalar o pandas pelo terminal do vs code. **Comando:** python -m pip install pandas

<img width="719" height="86" alt="image" src="https://github.com/user-attachments/assets/fe66fb15-7708-4243-ae10-3d83a9ae6342" />

4. Instalar o openpyxl. **Comando:** python -m pip install openpyxl
 
<img width="696" height="75" alt="image" src="https://github.com/user-attachments/assets/d1399be6-c3c0-496c-ae43-09d98f92d637" />
