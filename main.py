from PIL import Image
from PIL import ImageChops
import requests
import io
import pandas as pd
import itertools

#abre a planilha plugg.to
df = pd.read_excel('testecentauro2.xlsx')

#declara as listas
duplicados = []
resolucao = []
n = 1

for y in range(len(df['sku (*)']) - 1):
  print(y)

  links = [df['link_image_1'][n], df['link_image_2'][n], df['link_image_3'][n], df['link_image_4'][n], df['link_image_5'][n], df['link_image_6'][n]]

  for a, b in itertools.combinations(links, 2):
      #get todas as imagens
      link1 = requests.get(a)
      link2 = requests.get(b)

      #carrega as imagens na memória
      hash1 = io.BytesIO(link1.content)
      hash2 = io.BytesIO(link2.content)

      #abre as imagens no pillow
      img1 = Image.open(hash1)
      img2 = Image.open(hash2)

      #check de resolução da imagem
      if img1.size != (1000, 1000) and img1.size != (900, 900):
        #converte imagem em 1000x1000 via api da vtex
        resolucao_mil = a[:53] + "-1000-1000" + a[53:]

        #adiciona alteração como dicionário
        resolucao.append({a: resolucao_mil})

      if img2.size != (1000, 1000) and img2.size != (900, 900):
        #converte imagem em 1000x1000 via api da vtex
        resolucao_mil = a[:53] + "-1000-1000" + a[53:]

        #adiciona alteração como dicionário
        resolucao.append({a: resolucao_mil})

      #calcula a diferença
      diff = ImageChops.difference(img1, img2)

      if diff.getbbox() == None:
        #print(a, "e", b)
        duplicados.append(b)

  n = n + 1

df.replace(duplicados, "", inplace=True)

#loop de alteração dos produtos com dimensões inválidas
for produto in resolucao:
  print(produto)
  df.replace(to_replace = produto, inplace=True)

#exporta pra o excel
df.to_excel('duplicadas.xlsx', index=False) 

print('Acabou!')
