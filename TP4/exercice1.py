from ortools.linear_solver import pywraplp
import openpyxl

def TP4_Exercice1():
    file = "exercice1.xlsx"
    wb = openpyxl.load_workbook(file)
    ws = wb["Données"]
    
    # Lecture des paramètres
    n = ws.cell(1, 2).value  # Nombre de types d'objets
    capacity = ws.cell(2, 2).value  # Capacité par avion (tonnes)
    nb_max_bins = ws.cell(3, 2).value  # Nombre max de vols
    
    # Lecture des données des objets
    poids = [ws.cell(6, 2+i).value for i in range(n)]  
    nombre = [ws.cell(7, 2+i).value for i in range(n)]  
    
    print(f"Nombre de types d'objets: {n}")
    print(f"Capacité par avion: {capacity} tonnes")
    print(f"Nombre max de vols: {nb_max_bins}")
    
    # Création du solveur CBC
    solver = pywraplp.Solver('TP4_Exercice1', pywraplp.Solver.CBC_MIXED_INTEGER_PROGRAMMING)
    infinity = solver.infinity()
    
    # Variables de décision
    # x[i][j] = nombre d'objets de type i dans le vol j
    x = {}
    for i in range(n):
        for j in range(nb_max_bins):
            x[i, j] = solver.IntVar(0, nombre[i], f'x[{i},{j}]')
    
    # y[j] = 1 si le vol j est utilisé, 0 sinon
    y = {}
    for j in range(nb_max_bins):
        y[j] = solver.IntVar(0, 1, f'y[{j}]')
    
    print(f"Nombre de variables: {solver.NumVariables()}")
    
    # Fonction objectif: Minimiser le nombre de vols utilisés
    objective = solver.Objective()
    for j in range(nb_max_bins):
        objective.SetCoefficient(y[j], 1)
    objective.SetMinimization()
    
    # Contraintes
    # 1. Tous les objets doivent être transportés
    for i in range(n):
        contrainte_demande = solver.Constraint(nombre[i], nombre[i], f'demande_{i}')
        for j in range(nb_max_bins):
            contrainte_demande.SetCoefficient(x[i, j], 1)
    
    # 2. Capacité de chaque vol
    for j in range(nb_max_bins):
        contrainte_capacite = solver.Constraint(-infinity, capacity, f'capacite_{j}')
        for i in range(n):
            contrainte_capacite.SetCoefficient(x[i, j], poids[i])
    
    # 3. Lien entre x et y (si un vol contient des objets, il est utilisé)
    for i in range(n):
        for j in range(nb_max_bins):
            contrainte_lien = solver.Constraint(-infinity, 0, f'lien_{i}_{j}')
            contrainte_lien.SetCoefficient(x[i, j], 1)
            contrainte_lien.SetCoefficient(y[j], -nombre[i])
    
    print(f"Nombre de contraintes: {solver.NumConstraints()}")
    
    # Résolution
    print("\nRésolution en cours...")
    status = solver.Solve()
    
    if status == pywraplp.Solver.OPTIMAL:
        print("\nSolution optimale trouvée!")
        print(f"Nombre minimum de vols: {int(solver.Objective().Value())}")
        print(f"Temps de calcul: {solver.wall_time()} ms")
        print(f"Itérations: {solver.iterations()}")
        
        # Écriture des résultats dans Excel
        ws_res = wb["Résultats"]
        ws_res['B1'] = int(solver.Objective().Value())
        
        # Détail de chaque vol utilisé
        row = 4
        vols_utilises = 0
        
        for j in range(nb_max_bins):
            if y[j].solution_value() > 0.5:  # Vol utilisé
                vols_utilises += 1
                
                # Numéro du vol
                ws_res.cell(row, 1).value = vols_utilises
                
                # Calculer la charge totale
                charge_totale = sum(x[i, j].solution_value() * poids[i] for i in range(n))
                ws_res.cell(row, 2).value = charge_totale
                
                # Détail des objets (afficher 0 au lieu de cellule vide)
                for i in range(n):
                    nb_objets = int(x[i, j].solution_value())
                    ws_res.cell(row, 3+i).value = nb_objets
                
                row += 1
        
        wb.save(file)
        print(f"\nRésultats sauvegardés dans {file}")
        
    elif status == pywraplp.Solver.FEASIBLE:
        print("Une solution réalisable a été trouvée (mais peut-être pas optimale)")
        print(f"Nombre de vols: {int(solver.Objective().Value())}")
    else:
        print("Aucune solution trouvée")
        if status == pywraplp.Solver.INFEASIBLE:
            print("Le problème est infaisable.")
        elif status == pywraplp.Solver.UNBOUNDED:
            print("Le problème est non borné.")
    
    wb.close()

TP4_Exercice1()