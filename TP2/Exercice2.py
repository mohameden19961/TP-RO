from ortools.linear_solver import pywraplp

def TP2_Exo2():
    solver = pywraplp.Solver('Acier_Compact', pywraplp.Solver.GLOP_LINEAR_PROGRAMMING)
    n = 7
    
    Couts = [1.2, 1.5, 0.9, 1.3, 1.45, 1.2, 1.0]
    Stocks = [4000, 3000, 6000, 5000, 2000, 3000, 2500]
    
    Carbone = [2.5, 3.0, 0.0, 0.0, 0.0, 0.0, 0.0]
    Cuivre = [0.0, 0.0, 0.3, 90.0, 96.0, 0.4, 0.6]
    Aluminium = [1.3, 0.8, 0.0, 0.0, 4.0, 1.2, 0.0]
    
    D = 3000
    
    Carbone_min= 2.0
    Carbone_max = 3.0
    Cuivre_min= 0.4
    Cuivre_max = 0.6
    Aluminium_min = 1.2
    Aluminium_max = 1.65
    
    X = [solver.NumVar(0, Stocks[i], f'X{i+1}') for i in range(n)]
    
    solver.Minimize(sum(Couts[i] * X[i] for i in range(n)))
    
    solver.Add(sum(X[i] for i in range(n)) == D)
    
    solver.Add(sum(Carbone[i] * X[i] for i in range(n)) >= Carbone_min * D)
    solver.Add(sum(Carbone[i] * X[i] for i in range(n)) <= Carbone_max * D)
    
    solver.Add(sum(Cuivre[i] * X[i] for i in range(n)) >= Cuivre_min * D)
    solver.Add(sum(Cuivre[i] * X[i] for i in range(n)) <= Cuivre_max * D)
    
    solver.Add(sum(Aluminium[i] * X[i] for i in range(n)) >= Aluminium_min * D)
    solver.Add(sum(Aluminium[i] * X[i] for i in range(n)) <= Aluminium_max * D)
    
    status = solver.Solve()
    
    if status == pywraplp.Solver.OPTIMAL:
        print('Solution optimale trouvée :')
        print(f'Valeur optimale du coût = {solver.Objective().Value()} Euro')
        print('\nQuantités de matières premières :')
        noms = ['Fer1', 'Fer2', 'Fer3', 'Cuivre1', 'Cuivre2', 'Aluminium1', 'Aluminium2']
        for i in range(n):
            print(f'{noms[i]} = {X[i].solution_value()} kg')
        
    else:
        print('Le problème n\'a pas de solution optimale.')

TP2_Exo2()