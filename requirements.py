import subprocess

# Obtener lista de paquetes instalados
pip_freeze = subprocess.check_output(['pip', 'freeze']).decode('utf-8').split('\n')

# Crear el archivo requirements.txt
with open('requirements.txt', 'w') as f:
    for package in pip_freeze:
        f.write(package.split('==')[0] + '\n')
