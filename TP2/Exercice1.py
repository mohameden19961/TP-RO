from ortools.linear_solver import pywraplp

def TP2_Exo1():
    solver = pywraplp.Solver('Acier_Explicite', pywraplp.Solver.GLOP_LINEAR_PROGRAMMING)
    
    x1 = solver.NumVar(0, 4000, 'Fer1')
    x2 = solver.NumVar(0, 3000, 'Fer2')
    x3 = solver.NumVar(0, 6000, 'Fer3')
    x4 = solver.NumVar(0, 5000, 'Cuivre1')
    x5 = solver.NumVar(0, 2000, 'Cuivre2')
    x6 = solver.NumVar(0, 3000, 'Aluminium1')
    x7 = solver.NumVar(0, 2500, 'Aluminium2')

    solver.Minimize(1.2*x1 + 1.5*x2 + 0.9*x3 + 1.3*x4 + 1.45*x5 + 1.2*x6 + 1.0*x7)

    
    solver.Add(x1 + x2 + x3 + x4 + x5 + x6 + x7 == 3000)
    solver.Add(2.5*x1 + 3.0*x2 >= 2 * 3000)
    solver.Add(2.5*x1 + 3.0*x2 <= 3 * 3000)

    solver.Add(0.3*x3 + 90.0*x4 + 96.0*x5 + 0.4*x6 + 0.6*x7 >= 0.4 * 3000)
    solver.Add(0.3*x3 + 90.0*x4 + 96.0*x5 + 0.4*x6 + 0.6*x7 <= 0.6 * 3000)

    solver.Add(1.3*x1 + 0.8*x2 + 4.0*x5 + 1.2*x6 >= 1.2 * 3000)
    solver.Add(1.3*x1 + 0.8*x2 + 4.0*x5 + 1.2*x6 <= 1.65 * 3000)

    solver.Solve()

    if solver.Solve() == pywraplp.Solver.OPTIMAL:
        #pywrapp
        print('Solution optimale trouvée :')
        print(f'Valeur optimale du coût = {solver.Objective().Value()} Euro') 
        print(f'Fer1 = {x1.solution_value()} kg')
        print(f'Fer2 = {x2.solution_value()} kg')
        print(f'Fer3 = {x3.solution_value()} kg')
        print(f'Cuivre1 = {x4.solution_value()} kg')
        print(f'Cuivre2 = {x5.solution_value()} kg')
        print(f'Aluminium1 = {x6.solution_value()} kg')
        print(f'Aluminium2 = {x7.solution_value()} kg')
    else:
        print('Le problème n\'a pas de solution optimale.')

TP2_Exo1()