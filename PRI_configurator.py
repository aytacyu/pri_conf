import re,time
from xlrd import open_workbook
#import get_data_from_router as rtr
#import netmiko

filename=input("Dosya Adı:")
wb = open_workbook(filename+".xls")

sh1=wb.sheet_by_index(1)

for row in range(sh1.nrows):
    if re.search("SubeKodu",str(sh1.cell(row,0).value),re.I):        subeKodu=int(sh1.cell(row,1).value)
    if re.search("İlKodu",str(sh1.cell(row,0).value),re.I):        ilKodu=int((sh1.cell(row,1).value))
    if re.search("Santral No",str(sh1.cell(row,0).value),re.I):        santralNo=int(sh1.cell(row,1).value)
    if re.search("Fax No",str(sh1.cell(row,0).value),re.I):         faxNo=int(sh1.cell(row,1).value)
    if re.search("FaxSunucuIP",str(sh1.cell(row,0).value),re.I):        faxSunucuIP=str(sh1.cell(row,1).value)
    if re.search("RouterIP",str(sh1.cell(row,0).value),re.I):        routerIP=str(sh1.cell(row,1).value)
    if re.search("VoiceIP",str(sh1.cell(row,0).value),re.I):        voiceIP=str(sh1.cell(row,1).value)
    if re.search("Konferans Bridge",str(sh1.cell(row,0).value),re.I):        confBridge=str(sh1.cell(row,1).value)
    if re.search("CM Group",str(sh1.cell(row,0).value),re.I):
        cmGroup=str(sh1.cell(row,1).value)
        cm1=str(int(cmGroup[0])-1)
        cm2=str(int(cmGroup[1])-1)
    if re.search("SubeTipi",str(sh1.cell(row,0).value),re.I):
        if sh1.cell(row,1).value == "SCMT":
            scmtKodu=2591
        elif sh1.cell(row,1).value == "Verimlilik":
            scmtKodu=2577
        elif sh1.cell(row,1).value == "4+1":
            scmtKodu=2578

sh0=wb.sheet_by_index(0)
elements=[]
for row in range(sh0.nrows):
    dahili=0
    direkhat=0
    try:    dahili=int(sh0.cell(row,2).value)
    except ValueError:  pass
    if dahili != 0:
        element={}
        element["Dahili"]=subeKodu*10000+dahili
        try:    direkhat=int(sh0.cell(row,9).value)
        except ValueError:  pass
        if direkhat !=0:
            element["DirekHat"]=direkhat
            elements.append(element)

f=open("conf template.txt","r")
f2=open(filename+".txt","w")
for line in f:
    s1=re.search("^(.*)(#ilKodu)(.*)",line,re.I)
    s2=re.search("^(.*)(#faxNo)(.*)",line,re.I)
    s3=re.search("^(.*)(#faxSunucuIP)(.*)",line,re.I)
    s4=re.search("^(.*)(#scmtKodu)(.*)",line,re.I)
    s5=re.search("^(.*)(#santralNo)(.*)",line,re.I)
    s6=re.search("^(.*)(#confBridge)(.*)",line,re.I)
    s7=re.search("^(.*)(#routerIP)(.*)",line,re.I)
    s8=re.search("^(.*)(#subeKodu)(.*)",line,re.I)
    s9=re.search("^(.*)(#voiceIP)(.*)",line,re.I)
    s10=re.search("^(.*)(#cm1)(.*)",line,re.I)
    s11=re.search("^(.*)(#cm2)(.*)",line,re.I)
    if s1:       line=line.replace("#ilKodu",str(ilKodu))    #newline=s1.group(1)+str(ilKodu)+s1.group(3)+"\n"
    if s2:       line=line.replace("#faxNo",str(faxNo))  #newline=s2.group(1)+str(faxNo)+s2.group(3)+"\n"
    if s3:       line=line.replace("#faxSunucuIP",str(faxSunucuIP))
    if s4:       line=line.replace("#scmtKodu",str(scmtKodu))
    if s5:       line=line.replace("#santralNo",str(santralNo))
    if s6:       line=line.replace("#confBridge",str(confBridge))   #newline=s6.group(1)+confBridge+s6.group(3)
    if s7:       line=line.replace("#routerIP",str(routerIP))    #newline=s7.group(1)+routerIP+s7.group(3)
    if s8:       line=line.replace("#subeKodu",str(subeKodu))   #newline=s8.group(1)+str(subeKodu)+s8.group(3)
    if s9:       line=line.replace("#voiceIP",str(voiceIP))
    if s10:       line=line.replace("#cm1",str(cm1))
    if s11:       line=line.replace("#cm2",str(cm2))
    f2.write(line)
f.close()


f2.write("voice translation-rule 1\n")
f2.write("rule 1 /^.*\({}\)$/ /{}1200/\n".format(santralNo,subeKodu))
i=2
for element in elements:
    f2.write("rule {} /^.*\({}\)$/ /{}/\n".format(i,element["DirekHat"],element["Dahili"]))
    i+=1
f2.write("voice translation-rule 10\n")
i=1
for element in elements:
    f2.write("rule {} /^{}$/ /{}/\n".format(i,element["Dahili"],element["DirekHat"]))
    i+=1
f2.close


"""
enable_secret="secret"
client=rtr.connect(routerIP)
remote_conn = client.invoke_shell()
remote_conn.send("enable\n")
time.sleep(.5)
remote_conn.send("{}\n".format(enable_secret))
time.sleep(.5)
output=remote_conn.recv(65535)
time.sleep(.5)
f3=open("lastconf.txt","r")
f4=open("logs.txt","w")
for line in f3:
    print("*")
    output=rtr.sendCommand(remote_conn,line)
    f4.write(output)
    print(output)
client.close()
f3.close()
f4.close()
"""
