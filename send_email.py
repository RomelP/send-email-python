#---------------------------------------------
# Este script realiza la colocación y envio 
# por mail de documento en fprmarto pdf # 
# @utor: Romel Principe A # 
#---------------------------------------------
import pandas as pd
import yagmail
#---------------------------------------------
#      Envia los correos de manera ciclica 
#---------------------------------------------
# intento1 
# intento2
# intento3
# Cargando los datos a partir de la lista 
hoja = pd.read_excel("Lista.xlsx", sheet_name="nombres",header= 0) 
# Ingresando mail y passw
sender_email= "rprincipe@gmail.com" #input(f"please, enter the mail: ") 
sender_password = "xxxx-yyy-zzz"    #input(f"please, enter the password {sender_email}: ") 
# Enviando a todos
for i in range(156, len(hoja) ): 
    # Envia los correos de manera ciclica
    receiver_email = hoja["MAIL"][i]
    yag = yagmail.SMTP(user= sender_email, password = sender_password)
    subject = "Constancia Catedra IGP"
    contents = ["Buen dia, reciba un cordial saludo de parte del Instituto Geofisico del Peru.",
                "Se le está enviando la constancia de participación del evento del IGP", 
                "/.../certificados/" + # path de carpeta cetificados
                str("constancia"+ str(i+1) + ".pdf") ]
    # Destinatario de correo
    yag.send(to = receiver_email, subject = subject, contents = contents )
    print("Envio número: ", i+1 ,"-->", receiver_email)
#---------------------------------------------
