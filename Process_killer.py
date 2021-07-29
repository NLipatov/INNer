import psutil

def process_kill(process_name):
    # В случае удачного завершения всех процессов с указанным в аргументе "process_name" названием, верёнт "True"
    for processes in psutil.process_iter():
        try:
            pinfo = processes.as_dict(attrs=['pid', 'name'])
        except psutil.NoSuchProcess:
            pass
        if pinfo['name'] == f'{process_name}':
            processes.kill()
            return True