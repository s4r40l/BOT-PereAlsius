from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException
from telegram.ext import Updater, CommandHandler, MessageHandler, CallbackContext
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium import webdriver
from datetime import datetime
from telegram import Update
from telebot import types
from time import sleep
import pandas as pd
import subprocess
import telebot
import ctypes
import json
import csv
import os

file = open('configu.json')
params = json.load(file)



#Establim la variable del token de telegram que el fatherBot ens ha donat
TOKEN = '7502615219:AAG254MaDc01ATpEQ4mZ7BIlPDrBFZLt-8I'
bot = telebot.TeleBot(TOKEN) #Establim que aquest és el token del nostre bot


#Si fem la comanda que ens recomana al missatge que ens dona al començar, ens apareixeran 3 botons
#que ens donaran opcions de bots per executar.
@bot.message_handler(commands=['escollir']) 
def escollir(message):
	# Creem el teclat de botons amb les opcions per escollir
    markup = types.ReplyKeyboardMarkup(row_width=2, one_time_keyboard=True)
    markup.add(
        types.KeyboardButton("ComptadorImpressora"),
        types.KeyboardButton("AfegirUsuariImpressora"),
        types.KeyboardButton("ProfeSubstitut"),
        types.KeyboardButton("BorrarUsuari"),
        types.KeyboardButton("renovarCredencials")
    )
    resposta = bot.send_message(

        message.chat.id,
        "Quin bot vols executar?", #Aquest és el missatge que ens donarà abans dels botons
        reply_markup=markup
    )
    # Registrar el següent pas
    bot.register_next_step_handler(resposta, gestionar_opcio) #La resposta del usuari farà que se executin les comandes "gestionar_opcio"


#BOT A EXECUTAR
def gestionar_opcio(missatge): # Farem diversos if, un per cada opció del teclat de botons. Si no és cap de aquests ha de donar error (else)

#COMTPADOR IMPRESSORA
    if missatge.text == "ComptadorImpressora":
        try:
            #os.system('./executacomptimpr.sh')
            #crida a la funció
            comptadorImpressora(missatge)
        except Exception as e: #si dona error
            bot.send_message(missatge.chat.id, f"Hi ha hagut un problema amb l'escript: {str(e)}")

#AFEGRI USUARI A L'IMPRESSORA
    elif missatge.text == "AfegirUsuariImpressora":
        # per emmagatzemar les dades (llista)
        resposta = bot.send_message(missatge.chat.id,'Escriu les dades en el següet format: usuari,nom, cognoms,email,ID. No oblidi separar amb comes.')
        bot.register_next_step_handler(resposta, gestionarcsvimpr)

#PROFESSOR SUBSTITUT
    elif missatge.text == "ProfeSubstitut":
        # per emmagatzemar les dades (llista)
        resposta = bot.send_message(missatge.chat.id,'Escriu les dades en el següet format: usuari,nom,cognom1,cognom2,contrassenya,correu,usuari professor que substitueix. No oblidi separar amb comes.')
        bot.register_next_step_handler(resposta, gestionarcsv)

#BORRAR USUARI IMPRESSORA
    elif missatge.text == "BorrarUsuari":
        # per emmagatzemar les dades (llista)
        resposta = bot.send_message(missatge.chat.id,'Escriu els usuaris que vol eliminar en el següent format; usuari,nom,cognom. No oblidi separar amb comes.')
        bot.register_next_step_handler(resposta, csvborrarusu)
        
#RENOVAR CREDENCIALS
    elif missatge.text == "renovarCredencials":
        resposta = bot.send_message(missatge.chat.id,'Utilitzarà ID o Nom? (En minuscula)')
        bot.register_next_step_handler(resposta, gestionarcredencials)
        
    else:
        bot.send_message(missatge.chat.id, "Opció no vàlida. Torna a intentar-ho.")




#GESTIONAR BÚSQUEDA DE USUARI PER REGENERAR CREDENCIALS
def gestionarcredencials(missatge):
    try:
        if missatge.text == "id":
            resposta = bot.send_message(missatge.chat.id, "Escriu el número de identificació de l'alumne.")
            bot.register_next_step_handler(resposta, credencialsID) 
        elif missatge.text == "nom":
            resposta = bot.send_message(missatge.chat.id, "Escriu la informació de l'alumne en el següent format: nom,cognom1,cognom2,ensenyament,nivell. Recorda com funciona l'ensenyament: Batxiller Tecno, escriu 'BT', social: 'BS', eso: 'ESO', Cicle: 'FP'. Si no està exactament així, serà erroni.")
            bot.register_next_step_handler(resposta, credencialsNom)
    except Exception as e:
        bot.send_message(missatge.chat.id, f"Hi ha hagut un problema amb l'escript: {str(e)}")
		 
	
def credencialsID(missatge):
    try:
		
        file = open('configIdi.json')
        IdAlu = missatge.text
        bot.send_message(missatge.chat.id,'ID guardat!')
        driver = webdriver.Firefox()
        #obrirIdi(driver,missatge)
        params = json.load(file)
        sleep(2)
        driver.maximize_window()
        driver.get(params["Urlindic"])
        sleep(2)
        search_field = driver.find_element("id","user")
        search_field.send_keys(params["user"])
        sleep(2)
        search_field = driver.find_element("id","password")
        search_field.send_keys(params["password"])
        sleep(2)
        search_field = driver.find_element("xpath","//input[@type='submit']").submit()
        sleep(5)
        search_field = driver.find_element("xpath","//div[@class='t-Card u-color']").click()
        sleep(3)
        wait = WebDriverWait(driver,600)
        #Filtrar
        search_field = driver.find_element("name","P10_IDENTIFICADOR")
        search_field.clear()
        search_field.send_keys(IdAlu)   
        driver.find_element("xpath","//button[@title='Cerca']").click()
        sleep(5)
        
        #Selecciona la casella per regenerar credencials
        alumnes = driver.find_elements("xpath","//input[@name='f01']")
        if (len(alumnes) == 1):
            driver.find_element("xpath","//input[@name='f01']").click()
            #botó regenerar credencials
            driver.find_element("id","B162063473064565067").click()
            sleep(5)
            driver.find_element("xpath","//span[@class='fa fa-pencil edit-link-pencil']").click()
        else:
            bot.send_message(missatge.chat.id, "No hi ha cap coincidencia. Tancant bot...")
            driver.quit()
            bot.send_message(missatge.chat.id, "Bot tancat.")      
           
        #Baixem credencials
        nom=driver.find_element("name","P13_NOM").get_attribute("value")
        cognom=driver.find_element("name","P13_COGNOM1").get_attribute("value")
        bot.send_message(missatge.chat.id, "S'ha trobat l'alumne "+ nom + " "+ cognom + ". Regenerant contrassenya...")
        driver.find_element("id","B162306647501880342").click()   
        #canviem nom fitxer os.getcwd()+"/
        #TODO Posar directori baixades al fitxer de configuració
        bot.send_message(missatge.chat.id, "Baixant credencials...")
        while not os.path.exists("/home/super/Baixades/IDI_CREDENCIALS_IDI_CREDENCIALS.pdf"):
            sleep(1)
        if os.path.isfile("/home/super/Baixades/IDI_CREDENCIALS_IDI_CREDENCIALS.pdf"):
			
           os.rename("/home/super/Baixades/IDI_CREDENCIALS_IDI_CREDENCIALS.pdf", os.getcwd()+"/"+nom+cognom+".pdf")        
           bot.send_document(missatge.chat.id,open(os.getcwd()+"/"+nom+cognom+".pdf", 'rb'))     
        driver.find_element("id","B161869714135804447").click()
        bot.send_message(missatge.chat.id, "Tot correcte! Tantant bot...")
        #Tanca la finestra del navegador
        driver.quit()
        bot.send_message(missatge.chat.id, "Bot finalitzat. Finestra tancada.")
        
            
    except Exception as e: #si dona error
        bot.send_message(missatge.chat.id, f"Hi ha hagut un problema amb l'escript: {str(e)}")


def credencialsNom(missatge):
    file = open('configIdi.json')
    params = json.load(file)
    
    #csvFile = open("usuari.csv",mode='w') #obrim en mode w per sobreescriure
    data = []    
    user_data = missatge.text.split(',') #que les dades del csv es separin per comes
    nom, cognom1, cognom2, ensenyament, nivell = [item.strip() for item in user_data]
    bot.send_message(missatge.chat.id,'Dades guardades!')

    #obrir el navegador
    driver = webdriver.Firefox()
    sleep(2)
    driver.maximize_window()
    #Entrar a la url de indic
    driver.get(params["Urlindic"])
    sleep(2)

    #Validar usuari
    try:
        search_field = driver.find_element("id","user")
        search_field.send_keys(params["user"])
        sleep(2)
        search_field = driver.find_element("id","password")
        search_field.send_keys(params["password"])
        sleep(2)
        search_field = driver.find_element("xpath","//input[@type='submit']").submit()
        sleep(5)
    
        #entrar a la configuració d'usuaris
        search_field = driver.find_element("xpath","//div[@class='t-Card u-color']").click()
        sleep(3)
    except Exception as e:
        bot.send_message(missatge.chat.id, f"Hi ha hagut un problema amb l'escript: {str(e)}")
    
    #per cada usuari del csv
    wait = WebDriverWait(driver,600)
    #Filtrar
    search_field = driver.find_element("name","P10_IDENTIFICADOR")   
    search_field.clear()
    search_field = driver.find_element("name","P10_NOM")   
    search_field.clear()
    search_field.send_keys(nom)
    search_field = driver.find_element("name","P10_COGNOM1")   
    search_field.clear()
    search_field.send_keys(cognom1)
    search_field = driver.find_element("name","P10_COGNOM2")   
    search_field.clear()
    search_field.send_keys(cognom2)
    driver.find_element("id","P10_ENSENYAMENT_lov_btn").click()
    sleep(2)
    
    if (ensenyament == "BT"):
        driver.find_element("xpath","//li[starts-with(text(),'BATXLOE 2000')]").click()
    elif (ensenyament == "BS"):
        driver.find_element("xpath","//li[starts-with(text(),'BATXLOE 3000')]").click()
    elif (ensenyament == "ESO"):
        driver.find_element("xpath","//li[starts-with(text(),'ESO LOE')]").click()
    elif (ensenyament == "FP"):
        driver.find_element("xpath","//li[starts-with(text(),'CFPM AE10')]").click()
    else:
        bot.send_message(missatge.chat.id, "Opció no trobada.")
        driver.quit()
    sleep(2)    
    search_field = driver.find_element("name","P10_NIVELL")   
    search_field.send_keys(nivell)   
    
    driver.find_element("xpath","//button[@title='Cerca']").click()
    sleep(5)
    
    try: 
        a = driver.find_elements("xpath","//input[@name='f01']")
        if (len(a) == 1):
            driver.find_element("xpath","//input[@name='f01']").click()
            #botó regenerar credencials
            driver.find_element("id","B162063473064565067").click()
            sleep(5)
            driver.find_element("xpath","//span[@class='fa fa-pencil edit-link-pencil']").click()
            nom=driver.find_element("name","P13_NOM").get_attribute("value")
            cognom=driver.find_element("name","P13_COGNOM1").get_attribute("value")
            bot.send_message(missatge.chat.id, "S'ha trobat l'alumne "+ nom + " "+ cognom + ". Regenerant contrassenya...")
            driver.find_element("id","B162306647501880342").click()   
            #canviem nom fitxer os.getcwd()+"/
            #TODO Posar directori baixades al fitxer de configuració
            bot.send_message(missatge.chat.id, "Baixant credencials...")
            while not os.path.exists("/home/super/Baixades/IDI_CREDENCIALS_IDI_CREDENCIALS.pdf"):
                sleep(1)
            if os.path.isfile("/home/super/Baixades/IDI_CREDENCIALS_IDI_CREDENCIALS.pdf"):
               os.rename("/home/super/Baixades/IDI_CREDENCIALS_IDI_CREDENCIALS.pdf", os.getcwd()+"/"+nom+cognom+".pdf")        
               bot.send_document(missatge.chat.id,open(os.getcwd()+"/"+nom+cognom+".pdf", 'rb'))    
            driver.find_element("id","B161869714135804447").click()
            bot.send_message(missatge.chat.id, "Tot correcte! Tantant bot...")
            #Tanca la finestra del navegador
            driver.quit()
            bot.send_message(missatge.chat.id, "Bot finalitzat. Finestra tancada.")
 
            
        elif (len(a) == 0):
            bot.send_message(missatge.chat.id, "No hi ha cap coincidencia. Tancant bot")
            driver.quit()
        else:
            bot.send_message(missatge.chat.id, "N'hi ha més d'un. Prova amb un dels ID que veuràs a continuació:  ")
            nomalumnes = driver.find_elements("xpath",'//*[@headers="NOM"]')
            idralc = driver.find_elements("xpath",'//*[@headers="ID_ALUMNE_RALC"]')
            cognoms = driver.find_elements("xpath",'//*[@headers="COGNOM1"]')
            cognoms2 = driver.find_elements("xpath",'//*[@headers="COGNOM2"]')
            index =0
            opcions=types.ReplyKeyboardMarkup(row_width=2, one_time_keyboard=True)
            for nom in nomalumnes:
                opcions.add(types.KeyboardButton(idralc[index].text+","+nom.text+","+cognoms[index].text+","+cognoms2[index].text))
                index=index+1
                   
            resposta = bot.send_message(missatge.chat.id,"Quin alumne vols?", reply_markup=opcions)
            bot.register_next_step_handler(resposta, gestionar_alumne)
            driver.quit()
    except Exception as e:
        bot.send_message(missatge.chat.id, f"Hi ha hagut un problema amb l'escript: {str(e)}")
            
def gestionar_alumne(missatge): 

    try:
        missatgestring = missatge.text
        a = missatgestring.split(",")
        idalumneescollit = a[0]
        bot.send_message(missatge.chat.id, "Aquest és el Id: " + a[0])
        bot.send_message(missatge.chat.id, "Regenerant les credencials...")
		
        file = open('configIdi.json')
        
        driver = webdriver.Firefox()
        #obrirIdi(driver,missatge)
        params = json.load(file)
        sleep(2)
        driver.maximize_window()
        driver.get(params["Urlindic"])
        sleep(2)
        search_field = driver.find_element("id","user")
        search_field.send_keys(params["user"])
        sleep(2)
        search_field = driver.find_element("id","password")
        search_field.send_keys(params["password"])
        sleep(2)
        search_field = driver.find_element("xpath","//input[@type='submit']").submit()
        sleep(5)
        search_field = driver.find_element("xpath","//div[@class='t-Card u-color']").click()
        sleep(3)
        wait = WebDriverWait(driver,600)
        #Filtrar
        search_field = driver.find_element("name","P10_IDENTIFICADOR")
        search_field.clear()
        search_field.send_keys(idalumneescollit)
        driver.find_element("xpath","//button[@title='Cerca']").click()
        sleep(5)
        
        #Selecciona la casella per regenerar credencials
        alumnes = driver.find_elements("xpath","//input[@name='f01']")
        if (len(alumnes) == 1):
            driver.find_element("xpath","//input[@name='f01']").click()
            #botó regenerar credencials
            driver.find_element("id","B162063473064565067").click()
            sleep(5)
            driver.find_element("xpath","//span[@class='fa fa-pencil edit-link-pencil']").click()
        else:
            bot.send_message(missatge.chat.id, "No hi ha cap coincidencia. Tancant bot...")
            driver.quit()
            bot.send_message(missatge.chat.id, "Bot tancat.")      
           
        #Baixem credencials
        nom=driver.find_element("name","P13_NOM").get_attribute("value")
        cognom=driver.find_element("name","P13_COGNOM1").get_attribute("value")
        bot.send_message(missatge.chat.id, "Regenerant contrassenya...")
        driver.find_element("id","B162306647501880342").click()   
        #canviem nom fitxer os.getcwd()+"/
        #TODO Posar directori baixades al fitxer de configuració
        bot.send_message(missatge.chat.id, "Baixant credencials...")
        while not os.path.exists("/home/super/Baixades/IDI_CREDENCIALS_IDI_CREDENCIALS.pdf"):
            sleep(1)
        if os.path.isfile("/home/super/Baixades/IDI_CREDENCIALS_IDI_CREDENCIALS.pdf"):
           os.rename("/home/super/Baixades/IDI_CREDENCIALS_IDI_CREDENCIALS.pdf", os.getcwd()+"/"+nom+cognom+".pdf")        
           bot.send_document(missatge.chat.id,open(os.getcwd()+"/"+nom+cognom+".pdf", 'rb'))    
        driver.find_element("id","B161869714135804447").click()
        bot.send_message(missatge.chat.id, "Tot correcte! Tancant bot...")
        #Tanca la finestra del navegador
        driver.quit()
        bot.send_message(missatge.chat.id, "Bot finalitzat. Finestra tancada.")
        
            
    except Exception as e: #si dona error
        bot.send_message(missatge.chat.id, f"Hi ha hagut un problema amb l'escript: {str(e)}")
	

#EXECUCIÓ COMPTADOR IMPRESSORA
def comptadorImpressora(missatge):    
    file = open('configu.json')
    params = json.load(file)
    bot.send_message(missatge.chat.id, "Obrint el comptador de l'impresora...")
    #Carregar drive del selenium per treballar amb Chrome
    #from webdriver_manager.chrome import ChromeDriverManager
    #driver = webdriver.Chrome(service=Service(ChromeDriverManager().install())) 
    driver = webdriver.Firefox()
    driver.implicitly_wait(30)
    driver.maximize_window()

    # Acceder a la primer URL
    #URL IMPRESSORA
    driver.get(params["baseUrl"])
    main_page = driver.current_window_handle

    # Localizar cuadro de texto
    search_field = driver.find_element("id","inputUsername")
    search_field.send_keys(params["user"])
    search_field = driver.find_element("id","inputPassword")
    search_field.send_keys(params["password"])
    search_field = driver.find_element("name","$Submit$0")
    search_field.submit()
    sleep(1)

    #Revisar saldo
    search_field = driver.find_element("xpath","//a[@href='/app?service=direct/1/UserList/filter.toggle']").click()
    search_field = driver.find_element("id","maxBalance")
    search_field.send_keys("100")
    search_field = driver.find_element("name","apply")
    search_field.click()
    sleep(1)
    
    try:
        #busca els enllaços dels usuaris
        search_field = driver.find_elements("xpath","//a[starts-with(@href,'/app?service=direct/1/UserList/user.link')]")

        #definir llista d'usuaris
        enllacos = []
        #recorrer search_field i omplier array enllacos
        for s in search_field:
            #afegir a l'array ennlacos el s.get_attribute("href")
            enllacos.append(s.get_attribute("href"))

        print(enllacos)
        for h in enllacos:
            driver.get(h)
            sleep(1)
            fullName = driver.find_element("xpath","//input[@id='fullName']").get_attribute("value")
            bot.send_message(missatge.chat.id, "Actualitzant el comptador de l'impressora de " + fullName)
            sleep(1)
            #actualitzes sumbmit
            search_field = driver.find_element("xpath","//a[@href='/app?service=direct/1/UserDetails/$DirectLink&sp=1']").click()
            sleep(1)
            search_field = driver.find_element("id","adjustmentValue")        
            search_field.clear()
            search_field.send_keys("1000,00")
            sleep(1)
            search_field = driver.find_element("name","$Submit").click()
            #search_field = driver.find_element("xpath","//a[@href='/app?service=page/UserList']").click()
            sleep(2)
        bot.send_message(missatge.chat.id,'Comptador Fet')
        driver.quit()


    except NoSuchElementException:
        bot.send_message(missatge.chat.id,'No hi ha cap usuari. Comptador Fet')
        driver.quit()
    except ElementClickInterceptedException:
        bot.send_message(missatge.chat.id,"No s'ha pogut fer clic a l'usuari.")
        driver.quit()
    finally:
        driver.quit()




#GUARDA LES DADES QUE L'USUARI ENTRA DEL USUARI A AFEGIR DE LA IMPRESSORA
def gestionarcsvimpr(missatge):
    csvFile = open("usuari.csv",mode='w') #obrim en mode w per sobreescriure
    data = []    
    user_data = missatge.text.split(',') #que les dades del csv es separin per comes
    try:
        usuari, nom, cognoms, email, ID = [item.strip() for item in user_data]
        data.append({'usuari': usuari, 'nom': nom, 'cognoms': cognoms, 'email': email, 'ID': ID})
        df = pd.DataFrame(data)
        df.to_csv('usuari.csv', index=False) #guardem las dades a usuari.csv
        bot.send_message(missatge.chat.id,'Dades guardades!')
    except Exception as e:
        bot.send_message(missatge.chat.id, f"Hi ha hagut un problema amb l'escript: {str(e)}")

    #EXECUCIÓ AFEGIR USUARI IMPRESSORA
    try:
        bot.send_message(missatge.chat.id, "Afegint els usuaris a la impressora...")
        #os.system('./executamoodlesubs.sh')
        #Obrir firefox
        driver = webdriver.Firefox()
        driver.implicitly_wait(30)
        driver.maximize_window()

        # Acceder a la aplicación web
        #URL IMPRESSORA
        driver.get(params["URLusuari"])
        main_page = driver.current_window_handle
        sleep(2)

        # Localizar cuadro de texto
        search_field = driver.find_element("id","inputUsername")
        search_field.send_keys(params["user"])
        sleep(2)
        search_field = driver.find_element("id","inputPassword")
        search_field.send_keys(params["password"])
        sleep(2)
        search_field = driver.find_element("name","$Submit$0")
        search_field.submit()
        #csv file
        csvFile = open("usuari.csv",mode='r')
        users = csv.DictReader(csvFile)
        for user in users:
            search_field = driver.find_element("id","chosen-username")
            search_field.clear()
            sleep(1)
            search_field.send_keys(user["usuari"])
            search_field = driver.find_element("name","adminInputFullName")
            search_field.clear()
            search_field.send_keys(user["nom"]+params["espai"]+user["cognoms"])
            search_field = driver.find_element("name","adminInputEmail")
            search_field.clear()
            search_field.send_keys(user["email"])
            search_field = driver.find_element("name","adminInputPassword")
            search_field.clear()
            search_field.send_keys(user["ID"])
            search_field = driver.find_element("name","adminInputPasswordVerify")
            search_field.clear()
            search_field.send_keys(user["ID"])
            search_field = driver.find_element("name","adminInputUserID")
            search_field.clear()
            search_field.send_keys(user["ID"])
            search_field = driver.find_element("name","adminInputIDPIN")
            search_field.clear()
            search_field.send_keys(user["ID"])
            search_field = driver.find_element("name","adminInputIDPINVerify")
            search_field.clear()
            search_field.send_keys(user["ID"])
            sleep(1)
            search_field = driver.find_element("name","adminInputEmailConfirmation").click()
            #Enviar
            sleep(1)    
            search_field = driver.find_element("name","$Submit$1").click()
            #Modificar usuari
            sleep(1)
            search_field = driver.find_element("xpath","//a[@href='/app?service=page/UserList']").click()
            search_field = driver.find_element("id","quickFindAuto")
            search_field.clear()
            search_field.send_keys(user["usuari"])
            sleep(1)
            search_field = driver.find_element("name","$Submit").click() 

            #search_field = driver.find_element("xpath","//a[@href='/app?service=direct/1/UserDetails/$DirectLink&sp=1']").click()
            search_field = driver.find_element("id","userBalance")
            search_field.clear()
            search_field.send_keys("1000,00")
            search_field = driver.find_element("id","cardNumber2")
            search_field.clear()
            search_field.send_keys(user["ID"])
            sleep(1)
            search_field = driver.find_element("name","$Submit$1").click()
            sleep(1)
            search_field = driver.find_element("xpath","//a[@href='/app?service=page/UserList']").click()
            sleep(2)
            search_field = driver.find_element("id","pageactions").click()
            sleep(1)
            search_field = driver.find_element("xpath","//a[@href='/app?service=page/CreateInternalUser']").click() 
            bot.send_message(missatge.chat.id, "Usuari creat.")  
        # Cerrar la ventana del navegador
        csvFile.close()
        bot.send_message(missatge.chat.id, "Tancant finestra...")
        driver.quit()
        bot.send_message(missatge.chat.id, "Finestra tancada.")
    except Exception as e:
        bot.send_message(missatge.chat.id, f"Hi ha hagut un problema amb l'escript: {str(e)}")   


#EXECUCIÓ BORRAR USUARI 
def csvborrarusu(missatge):

    csvFile = open("csvborrarusu.csv",mode='w')
    data = []    
    user_data = missatge.text.split(',')
    usuari, nom, cognoms = [item.strip() for item in user_data]
    data.append({'usuari': usuari, 'nom': nom, 'cognoms': cognoms})
    df = pd.DataFrame(data)
    df.to_csv('csvborrarusu.csv', index=False)
    bot.send_message(missatge.chat.id,'Dades guardades!')
    bot.send_message(missatge.chat.id, "Esborrant usuaris...")

    #Obrir firefox
    driver = webdriver.Firefox()
    driver.implicitly_wait(30)
    driver.maximize_window()

    # Acceder a la aplicación web
    #URL IMPRESSORA
    driver.get(params["borrarurl"])
    main_page = driver.current_window_handle
    sleep(2)

    #Localizar cuadro de texto
    search_field = driver.find_element("id","inputUsername")
    search_field.send_keys(params["user"])
    sleep(2)
    search_field = driver.find_element("id","inputPassword")
    search_field.send_keys(params["password"])
    sleep(2)
    search_field = driver.find_element("name","$Submit$0")
    search_field.submit()

    search_field = driver.find_element("xpath","//a[@class='btn secondary']").click()
    search_field = driver.find_element("id","minBalance")
    search_field.send_keys(1000)
    search_field = driver.find_element("id","maxBalance")
    search_field.send_keys(1000)
    search_field = driver.find_element("name","apply").click()

    Filecsv = open("csvborrarusu.csv",mode='r')
    usuaris = csv.DictReader(Filecsv)
    for usuari in usuaris:
        search_field = driver.find_element("id","quickFindAuto")
        search_field.send_keys(usuari["usuari"])    
        search_field = driver.find_element("name","$Submit$0")
        search_field.submit()
        sleep(2)
        search_field = driver.find_element("xpath","//span[@class='title']").click()
        sleep(1)
        try:
            driver.find_element("xpath","//a[contains(@href, '/app?service=direct/1/UserDetails/deleteUser')]").click() 
            alert = driver.switch_to.alert
            alert.accept()  # Para aceptar
            bot.send_message(missatge.chat.id, "Usuari eliminat. Tancant finestra.")
            driver.quit()
        except NoSuchElementException:
            print("Usuari no trobat.")
            bot.send_message(missatge.chat.id, "Usuari no trobat. Tancant finestra.")
            driver.quit()
            
        except Exception as e:
            print("Usuari no trobat. Tancant finestra.")
            bot.send_message(missatge.chat.id, "Usuari no trobat. Tancant finestra.")
            driver.quit()





#GUARDA LES DADES QUE L'USUARI ENTRA DEL PROFESSOR SUBSTITUT I DESPRÉS EXECUTA EL BOT
def gestionarcsv(missatge):
    csvFile = open("usuarisubstitut.csv",mode='w')
    data = []    
    user_data = missatge.text.split(',')
    usuari, nom, cognom1, cognom2, contra, correu, usuprof = [item.strip() for item in user_data]
    data.append({'Usuari': usuari, 'Nom': nom, 'Primer Cognom': cognom1, 'Segon Cognom': cognom2, 'Contrassenya': contra, 'Correu': correu, 'Usuari professor original': usuprof})
    df = pd.DataFrame(data)
    df.to_csv('usuarisubstitut.csv', index=False)
    bot.send_message(missatge.chat.id,'Dades guardades!')
    try:
        bot.send_message(missatge.chat.id, "Obrint el arxiu per afegir l'usuari substitut...")
        #os.system('./executamoodlesubs.sh')

        file = open('configsubs.json')
        params = json.load(file)

        driver = webdriver.Firefox()

        # Acceder al URL del login del moodle
        driver.get(params["baseUrl"])
        main_page = driver.current_window_handle

        #INICIAR SESSIÓ AMB L'ADMINISTRADOR DEL MOODLE
        # Localizar input per al usuari
        search_field = driver.find_element("id","username")
        search_field.send_keys(params["user"])
        # Localitzar input per a la contrasenya
        search_field = driver.find_element("id","password")
        search_field.send_keys(params["password"])
        # Enviar
        search_field = driver.find_element("id","loginbtn")
        search_field.submit()
        sleep(2)


        # Entrar al Url per crear un usuari
        driver.get(params["urlmkusr"])
        sleep(1)

        #AGAFAR LES DADES DEL CSV PER CREAR L'USUARI
        csvFile = open("usuarisubstitut.csv",mode='r')
        users = csv.DictReader(csvFile)
        for user in users:
            search_field = driver.find_element("id","id_firstname")
            search_field.send_keys(user["Nom"])
            search_field = driver.find_element("id","id_lastname")
            search_field.send_keys(user["Primer Cognom"])
            search_field.send_keys(params["espai"])
            search_field.send_keys(user["Segon Cognom"])
            driver.find_element("xpath","//a[contains(@class, 'form-control')]").click() 
            search_field = driver.find_element("xpath","//input[contains(@id, 'id_newpassword')]")
            search_field.send_keys(user["Contrassenya"])
            search_field = driver.find_element("id","id_email") 
            search_field.send_keys(user["Correu"])
            search_field = driver.find_element("id","id_username") 
            search_field.send_keys(user["Usuari"])
        sleep(2)

        # Envia formulari
        search_field = driver.find_element("id","id_submitbutton")
        search_field.submit()
        sleep(1)

        # BUSCAR ELS CURSOS ALS QUALS EL PROFESSOR SUBSTITUT S'HAURÀ D'APUNTAR
        # Entra al menú
        driver.get(params["urlcurs"])
        sleep(1)
        # Entrar als paràmetres per afegir el filtre del nom d'usuari
        driver.find_element("xpath","//a[@class='moreless-toggler']").click()
        # Buscar i emplenar el nom d'usuari
        search_field = driver.find_element("id","id_username")
        search_field.send_keys(user["Usuari professor original"])
        sleep(1)
        # Clicar botó per enviar formulari
        search_field = driver.find_element("id","id_addfilter")
        search_field.submit()
        sleep(1)

        # obrir ajustos de l'usuari

        try: 
            driver.find_element("xpath","//a[contains(@href, '../user/view.php?id=')]").click()
            sleep (2)
        except NoSuchElementException:
            print("No se encontró el enlace.")
            driver.quit()
        except ElementClickInterceptedException:
            print("No se pudo hacer clic en el enlace.")    
            driver.quit()

        try:
            # Fer clic a la opció de veure tots els enllaços si és que hi ha l'opció
            driver.find_element("xpath","//a[contains(@href, 'showallcourses=1')]").click()
            sleep(3)
        except NoSuchElementException:
            print("No se encontró el enlace ""view more"", continuando con la siguiente instrucción.")
        except ElementClickInterceptedException:
            print("No se pudo hacer clic en el enlace, continuando con la siguiente instrucción.")


        #APUNTAR AL PROFESSOR SUBSTITUT ALS CURSOS
        #guardem tots els enllaços que continguin showallcourses en una llista
        enllacos = driver.find_elements("xpath","//a[contains(@href, 'user/view.php?id=')]")
        #fem que aquestos enllaços que eran string tinguin l'atribut d'un enllaç (cada un)
        hrefs = [enllac.get_attribute("href") for enllac in enllacos]

        # Farà el seguent amb cada element de la llista (un per un)
        for href in hrefs:
            #Entra a l'enllaç
            driver.get(href)
            #Busca un enllaç que contingui "/user/index.php" i fa clic
            driver.find_element("xpath","//a[contains(@href, '/user/index.php?')]").click()
            sleep(3)
            #agafa el valor del h2 que és el nom del curs i el estableix a una variable que es diu nomcurs
                #nomcurs = driver.find_element("xpath","//h1[@class='h2']").get_attribute("value")
                #informacio = "El usuari" + (user["Usuari"]) + "està inscrit al curs " + nomcurs
            #Busca un element que el seu valor contingui "usu" i fa clic
            driver.find_element("xpath","//input[contains(@value, 'usu')]").click()
            sleep(3)
            #Busca un camp per escriure que el seu rol contingui "combobox" i emplena amb l'usuari que hem creat
            #al principi
            search_field = driver.find_element("xpath","//*[contains(@role, 'combobox')]")
            search_field.send_keys(user["Nom"])
            search_field.send_keys(params["espai"])
            search_field.send_keys(user["Primer Cognom"])
            search_field.send_keys(params["espai"])
            search_field.send_keys(user["Segon Cognom"])
            sleep(2)
            try:
                # Fer clic a la opció de veure tots els enllaços si és que hi ha l'opció
                driver.find_element("xpath","//*[contains(@role, 'option')]").click()
                sleep(3)
                bot.send_message(missatge.chat.id, "El usuari" + (user["Usuari"]) + "està inscrit al curs. ")
            except NoSuchElementException:
                print("Usuari no trobat, seguint amb les instruccions...")
            except ElementClickInterceptedException:
                print("No s'ha pogut fer clic a l'element. Continuant amb les instruccions...")    
            sleep(3)
            #fa clic al botó per afegir l'usuari que acabem d'entrar els paràmetres
            driver.find_element("xpath","//*[contains(@data-action, 'save')]")
            search_field.submit()
            sleep(3)

        driver.quit()

        
    except Exception as e:
        bot.send_message(missatge.chat.id, f"Hi ha hagut un problema amb l'escript: {str(e)}") 
    
    
    
    
    
#COMANDA /START
@bot.message_handler(commands=['start'])
def start_command(message):
    bot.send_message(message.chat.id,'Hola, aquest és un canal de execució de bots, si vols continuar escriu "/escollir"') #ens retornarà aquest missatge en escriure la comanda

bot.polling()


