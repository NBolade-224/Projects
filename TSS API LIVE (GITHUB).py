import requests, json, time
from base64 import b64encode

usern ="Username goes here"
passw = "Password goes here"
userAndPass = b64encode(b"%s:%s" % (usern,passw)).decode("ascii")
ses = requests.Session()
hrs = {
'Accept':'application/json',
'Content-Type':'application/json',
'Request':'application/json',
'Authorization' : 'Basic %s' %  userAndPass
}
Gu = "https://api.tsstestenv.co.uk/api/x_fhmrc_tss_api/v1/tss_api/goods"
#############################################################
## FINDS ALL DRAFT SUPS AND TURNS THEM INTO A LIST OF IDS ###
#############################################################
DraftSUPs = ses.get("https://api.tsstestenv.co.uk/api/x_fhmrc_tss_api/v1/tss_api/supplementary_declarations?filter=status=draft",headers=hrs)                     ##### THIS IS A LIST OF ALL SUP REFERENCES
SUPDecList = DraftSUPs.json()['result']
Decs = []  
for y in SUPDecList:
    Decs.append(y['number'])
##############################################################
#### FINDS ALL GOODS ID FOR EACH DRAFT SUP REFERENCE (-1) ####
##############################################################
l = 1
nn = len(Decs)
for z in Decs:
    st = time.time()
    Gd = ses.get("https://api.tsstestenv.co.uk/api/x_fhmrc_tss_api/v1/tss_api/goods?sup_dec_number=%s" % z,headers=hrs)
    ls = Gd.json()['result']['goods'] ## LIST OF ALL GOODS IN A SUP DEC
    if len(ls) > 1: 
        k = 1
        jj = len(ls) - 1        
        for x in ls[:-1]:
            tp = {
            "op_type":"delete",
            "goods_id":"%s" % x['goods_id']
            }
            Td = ses.post(Gu,json=tp,headers=hrs)        
            print(str(Td)+"Sup %d/%d - Good %d/%d" % (l, nn, k, jj))
            if str(Td) == "<Response [200]>":
                pass
            else:
                print("Error - Check Code")
                break
            k += 1
    else:
        pass
    l += 1
    ed = time.time()
    print(ed - st)
    time.sleep(1)
#############################################################
################# the end and fully working #################
#############################################################

##############################################################
######### UPLOAD GOODS TO A SPECIFIED CONSIGNMENT (SFD)#######
##############################################################
# for x in range(99):
#     ToPostGoodCreate = {
#     "op_type":"create",
#     "consignment_number":"DEC000000001007622",
#     "goods_id":"",
#     "equipment_number":"",
#     "un_dangerous_goods_code":"",
#     "type_of_packages":"boxes",
#     "number_of_packages":"1",
#     "number_of_individual_pieces":"",
#     "package_marks":"ADDR",
#     "gross_mass_kg":"2",
#     "net_mass_kg":"",
#     "goods_description":"TEST %s" % x,
#     "invoice_number":"",
#     "controlled_goods":"no",
#     "controlled_goods_type":"",
#     "commodity_code":"",
#     "country_of_origin":"",
#     "item_invoice_amount":"",
#     "item_invoice_currency":"",
#     "procedure_code":"",
#     "additional_procedure_code":"",
#     "taric_code":"",
#     "cus_code":"",
#     }
#     TestPost = ses.post("https://api.tsstestenv.co.uk/api/x_fhmrc_tss_api/v1/tss_api/goods",json=ToPostGoodCreate,headers=hrs)        
#     print(TestPost)
#     print(TestPost.json())

