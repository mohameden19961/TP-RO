from ortools.linear_solver import pywraplp
import openpyxl

def TP3_Exercice1():
    file = "exercice1.xlsx"
    wb = openpyxl.load_workbook(file)
    
    ws = wb["Données"]
    n = ws.cell(1, 2).value  
    capacity = ws.cell(2, 2).value  
    benefice = [ws.cell(4, 2+i).value for i in range(n)]  
    poids = [ws.cell(5, 2+i).value for i in range(n)]  
    
    print(f"Nombre d'objets: {n}")
    print(f"Capacité de la camionnette: {capacity} kg")
    
    solver = pywraplp.Solver('TP3_Exercice1', pywraplp.Solver.CBC_MIXED_INTEGER_PROGRAMMING)
    infinity = solver.infinity()
    
    
    X = []
    for i in range(n):
        X.append(solver.IntVar(0, 1, f'X[{i}]'))
    

    objective = solver.Objective()
    for i in range(n):
        objective.SetCoefficient(X[i], benefice[i])
    objective.SetMaximization()
    
    contrainte_capacite = solver.Constraint(-infinity, capacity, 'capacite')

    for i in range(n):
        contrainte_capacite.SetCoefficient(X[i], poids[i])
    
    status = solver.Solve()

    if status == pywraplp.Solver.OPTIMAL:
        print("Solution optimale trouvée!")
        print(f"Bénéfice optimal: {solver.Objective().Value()}")
    
        ws = wb["Résultats"]
        ws['B1'] = solver.Objective().Value()  
        
        for i in range(n):
            ws.cell(3, 2+i).value = X[i].solution_value()
        
        wb.save(file)
        print(f"\nRésultats sauvegardés dans {file}")
        
    elif status == pywraplp.Solver.FEASIBLE:
        print("Une solution réalisable a été trouvée (mais peut-être pas optimale)")
        print(f"Bénéfice: {solver.Objective().Value()}")
    else:
        print("Aucune solution trouvée")


    wb.close()

TP3_Exercice1()

