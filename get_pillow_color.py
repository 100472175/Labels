from PIL import Image
import pprint

img = Image.open("base/base.png")
list = (img.getcolors(100000))
lista2 = []
for i in list:
    if i[0] > 100:
        lista2.append(i)

for i in lista2:
    print(i)