import openpyxl
from PIL import Image
from PIL import ImageDraw
from PIL import ImageFont


 # Abrindo o arquivo excel
planilha = openpyxl.load_workbook('./Formandos.xlsx')

 # Escolhendo a página
pagina = planilha['Plan1']


 #Escolhendo a fonte
fonte = ImageFont.truetype("./CAMBRIAI.TTF",200)

 # Para cada linha na tabela, coletar dados
for linhas in (pagina.iter_rows(min_row=2,max_row=4)):

  # Escolhendo imagem para cada arquivo
  imagem = Image.open('./Certificado.png')
  #Função para escrever na imagem
  escrever = ImageDraw.Draw(imagem)

  #Salvando os dados da tabela em variáveis
  nome = linhas[0].value
  curso = linhas[1].value
  nome_Facul = linhas[2].value
  horas = linhas[3].value
  horas = str(horas)


  #Escrendo na imagem
  escrever.text((1561,393),nome_Facul,fill='black',font=fonte)
  escrever.text((1957,1200),nome,fill='black',font=fonte)
  escrever.text((1956,1890),curso,fill='black',font=fonte)
  escrever.text((1957, 2485),horas,fill='black',font=fonte)

  #Salvando o arquivo
  imagem.save(f'Certificado_{nome}.png')