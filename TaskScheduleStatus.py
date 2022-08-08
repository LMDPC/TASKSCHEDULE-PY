import win32com.client as w32
import pyodbc
import smtplib
import ssl


TASK_STATE = {0: 'Unknown', 1: 'Disabled', 2: 'Queued', 3: 'Ready', 4: 'Running'}   
                                    
scheduler = w32.Dispatch("Schedule.Service")
scheduler.Connect()
n = 0
objTaskFolders = [scheduler.GetFolder("\\")]

while objTaskFolders:
    
      objTaskFolder = objTaskFolders.pop(0)
      objTaskFolders += list(objTaskFolder.GetFolders(0))
      colTasks = objTaskFolder.GetTasks(1)
      n += len(colTasks)

      for task in colTasks:           
            colTasks =str(task.Path)[1:]
            if task.Name in ['Actualizador TARIFA','EDISERVER','INVENTARIOXMES','soiaq','NOTIFICACION_INVENTARIO_DIARIO']:                       
                print("Name: ", task.Name)
                #print(type(task.Name))
                print("State: ",TASK_STATE[task.State])
                #print(type(TASK_STATE[task.State]))
                print("Enabled: ",task.Enabled)
                #print(type(task.Enabled))
                print("Last Run Time: ",task.LastRunTime)
                #print(type(task.LastRunTime))
                print("Last Task Result:",task.LastTaskResult)
                #print(type(task.LastTaskResult))
                print("Next Run Time: ",task.NextRunTime)
                #print(type(task.NextRunTime))
                print("Number of Missed Runs: ",task.NumberOfMissedRuns)
                #print(type(task.NumberOfMissedRuns)) 
                
                #print(type(task.State))
                #print("Path:")
                #print(task.Path)
                #print(type(task.Path))
                #print("XMl:")
                #print(task.XML)
                
                server = 'SERVER' 
                database = 'DATABASE' 
                username = 'user' 
                password = 'password'                                               
                cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
                cursor = cnxn.cursor()         
                sql = "INSERT INTO Servicios.ServerLTX(Name,Hidden,State,LastRun,LastResult,NextRun,MissedRun,Comment) VALUES ('v1', 'v2', 'v3', 'v4', v5,'v6',v7,'v8')"
                sql2= "DECLARE @DEF nvarchar(250) SELECT @DEF = Definition FROM Servicios.LastRunCode  where Code = 'hexa' SELECT ISNULL(@DEF,'Definition not found') as Definition"
                
                sql = sql.replace('v1',task.Name)  
                sql = sql.replace('v2',str(task.Enabled))
                sql = sql.replace('v3',TASK_STATE[task.State] )              
                sql = sql.replace('v4',str(task.LastRunTime)[0:19])
                sql2 = sql2.replace('hexa', str(hex(task.LastTaskResult)))  
                Comments = cursor.execute(sql2).fetchone()[0]
                print(Comments)
                sql = sql.replace('v5',str(task.LastTaskResult))
                sql = sql.replace('v6',str(task.NextRunTime)[0:19])
                sql = sql.replace('v7',str(task.NumberOfMissedRuns))
                sql = sql.replace('v8',str(Comments) )
                print(sql)  
                                     
                cursor.execute(sql)
                cursor.commit() 
                print('Listed %d tasks.' % n)
                
smtp_server = "Servidor de correo"
port = 587  # For starttls
sender_email = "test@outlook.com"
receiver_email = "test@outlook.com"
password = "PASSWORD"
message = """Subject: servServicePY(LTX) it was executed successfully!!"""

context = ssl.create_default_context()
with smtplib.SMTP(smtp_server, port) as server:
    server.ehlo()  # Can be omitted
    server.starttls(context=context)
    server.ehlo()  # Can be omitted
    server.login(sender_email, password)
    server.sendmail(sender_email, receiver_email, message)
