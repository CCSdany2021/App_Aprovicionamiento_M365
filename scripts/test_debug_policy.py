
import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))
from scripts.sincronizador_politicas_teams import SincronizadorPoliticasTeams

def run_test():
    csv_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'test_policy.csv'))
    print(f"Testing with file: {csv_path}")
    
    if not os.path.exists(csv_path):
        print("Test file not found!")
        return

    sinc = SincronizadorPoliticasTeams()
    # Ejecutar en modo manual pasando el archivo
    for result in sinc.ejecutar(filepath=csv_path):
        print(result)

if __name__ == "__main__":
    run_test()
