
def remplir_liste(nombre):
    liste=[]
    for i in range(nombre):
        element=int(input(f"Saisir la note d'etudiant {i+1} :"))
        if element > 0 and element < 20 :
            liste.append(element)
         

    for i in range(len(liste)) :
        for j in range(i+1,len(liste)):
            if liste[j] < liste[i]:
                temp=liste[i]
                liste[i]=liste[j]
                liste[j]=temp     
    return liste

def main():
    nombre = nombre_element()   
    liste  = remplir_liste(nombre)
    print(f"la liste des notes = {liste}")

def nombre_element():
   nombre=int(input("saisir le nombre des etudiant:"))
   return nombre

 


main()
