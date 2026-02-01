from ortools.linear_solver import pywraplp
def TP1_Exo1():
    # Appel du solver GLOP des problèmes linéaires en nombres réels
    solver = pywraplp.Solver('TP1_Exo1', pywraplp.Solver.GLOP_LINEAR_PROGRAMMING)
    infinity = solver.infinity() #+infini 
    # Création des variables réelles x et y.
    x = solver.NumVar(0, infinity, 'x')
    y = solver.NumVar(0, infinity, 'y')
    print('Nombre des variables =', solver.NumVariables())
    # Minimize 50*x + 70*y.
    solver.Minimize(50*x + 70*y)
    # Création des contraintes.
    # 20*x + 30*y <= 360.
    solver.Add(20*x + 30*y <= 360)
    # 40*x + 35*y <= 480.
    solver.Add(40*x + 35*y <= 480)
    print('Nombre des contraintes =', solver.NumConstraints())
    solver.Solve()

    print('Solution:')
    print('Valeur optimale =', solver.Objective().Value())# Zmax
    print('x =', x.solution_value())
    print('y =', y.solution_value())
TP1_Exo1()


# 
#
#
#
#
#
#
#