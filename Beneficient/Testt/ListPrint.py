TestName=['A','B']
TestName2=['C','D']
B=""
C=""
for io in range(len(TestName)):
    B = B+" \n\n"+str(io+1)+") "+"".join(TestName[io])+"=>"+"".join(TestName2[io])
   # C = C + " " + "".join(TestName2[io])
print(B)