import socket
# import subprocess
# import os
sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
result = sock.connect_ex(('127.0.0.1',5000)) # 52.163.121.219
hostname = socket.gethostname()
local_ip = socket.gethostbyname(hostname)
print(local_ip)

if result == 0:
   print ("Port is open")
   #output = subprocess.getoutput("sudo /etc/init.d/asterisk status", shell=True)
   # os.system("sudo /etc/init.d/asterisk status")
   # mes=local_ip+" Port is open"
   #line_notify(token, mes)
else:
   print ("Port is not open")
   # os.system("sudo /etc/init.d/asterisk restart")
   #output = subprocess.getoutput("sudo /etc/init.d/asterisk restart", shell=True)
   # mes=local_ip+"Port is open"
   # print(mes)
   # line_notify(token, mes)
sock.close()
