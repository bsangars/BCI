from datetime import datetime
def Welcome(name):
    now =datetime.now()
    content = "Welcome "+name+ "The time now is " + str(now)
    return content
print (Welcome('Bhargav'))
