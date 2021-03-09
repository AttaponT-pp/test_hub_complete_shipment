from win32com import client

rev = '2.9.0.0'
fits_dll = client.Dispatch("FITSDLL.clsDB")

opn101_param = 'Build Type,Mother Lot Qty,PO No.,MFG Part Number (original),Prefix (Serial No),Part No.,Supplier name,' \
               'PID,Test Sampling type,COO on Box'
opn151_param = 'RT,MFG Part Number,PID,Part Number,Supplier Name,Sample Type,SampQty,Mother Lot Qty'

opn601_b_param = 'Packing No,Packing Qty,Box no. Qty in Opn.502'

opn602_param = 'Packing Qty,Shipment Request Qty'

opn1501_param = 'OPERATOR,Invoice No,Packing No,Packing Qty,RT,Pass Kitting Qty,Build Type,' \
                                    'PID,PO No.,Part Number,MFG Part Number,Supplier Name, Mother Lot Qty,' \
                                    'Shipment Request Qty,Test Sampling type,SampQty,Inspection Result'

opn702_param = 'OPERATOR,Invoice No,Packing Qty,Packing No'

def init(opn):
    # Give location of dll
    # fits_dll = client.Dispatch("FITSDLL.clsDB")
    status = fits_dll.fn_InitDB(opn, rev,'')

    print("init=" + status)
    return status


def handshake(fits_dll,opn,inv):
    # fn_handshake(operation,revision,serial/invoice)
    # fits_dll = client.Dispatch("FITSDLL.clsDB")
    status = fits_dll.fn_handshake('*', opn, rev, inv)

    print("handshake=" + status)
    return status


def query(fits_dll, opn, sn, param, fs):
    # fn_query(model,operation,revision,serial,parameters[,fsp]);
    status = fits_dll.fn_query('*', opn, rev, sn,param,fs)
    print("query status: " + str(status))
    return status


def log(opt,param,data,fs):
    # fn_log(model,operation,revision,parameters,values[,fsp]);
    global model, rev

    fits_dll = client.Dispatch("FITSDLL.clsDB")
    status = fits_dll.fn_InitDB(model, opt, rev, "")
    status = fits_dll.fn_log(model,opt,rev,param,data,fs)
    print("log status: " + str(status))
    return status


def valid_inv(opn, inv):

    # fits_dll = client.Dispatch("FITSDLL.clsDB")

    if init(opn) == 'True':
        if handshake(fits_dll,opn,inv) == 'True':
            return {"status": True, "msg": ""}
        else:
            return {"status": False, "msg": "This invoice is not valid at %s operation" % opn}
    else:
        return {"status": False, "msg": "Cannot init FIT DB!"}


def get_necessory_data(opn, rt, param):
    # fn_query(model,operation,revision,serial,parameters[,fsp]);
    # print 'input param: ' + param
    status = fits_dll.fn_query('*', opn, rev, rt, param, ',')
    # print "query status: " + str(status)
    output = status.split(',')
    # print output

    return output


def record2fit(opn, param ,data):

    status = fits_dll.fn_log('*', opn, rev, param, data, ',')
    # print status
    return status


# Check RTV Shipment Blocking Status
def check_block_rtv(etr):
    if init("*") == "True":
        # Get RT from opn.1303 ETR
        rt = fits_dll.fn_query("*", "1303", rev, etr, "RT")
        print("RT of ETR number:"+ etr + " " + "is " + rt)
        # Check RT in Opn.924 RTV Shipment Blocking
        result = fits_dll.fn_query("*", "924", rev, rt, "RTV Shipment Blocking")
        print("RTV Shipment Blocking = " + result)
        return result
    else:
        result = int("*")
        return result


def get_receiving(rt):
    print("Input RT = {}".format(rt))
    fits_dll = client.Dispatch("FITSDLL.clsDB")
    if not fits_dll.fn_InitDB('*', rev, ''):
        return False
    opn101_data = fits_dll.fn_query('*', '101', rev, rt, opn101_param, ',')
    print(opn101_data)
    sn_list = str(fits_dll.fn_query('*', '151', 'RT', '*', rt, ','))
    print(sn_list)
    sn = sn_list.split(',')
    return True


def find_packing_num(rt):
    print("Input RT = {}".format(rt))
    fits_dll = client.Dispatch("FITSDLL.clsDB")
    if not fits_dll.fn_InitDB('*', rev, ''):
        return False
    sn_list = str(fits_dll.fn_query('*', '151', 'RT', '*', rt, ','))
    print(sn_list)
    sn = sn_list.split(',')
    print('Length of RT: {} = {}'.format(rt, len(sn)))
    packing_num_list = []
    for i in range(len(sn)):
        # get packing numbers
        packing_num = fits_dll.fn_query('*', '601_B', rev, sn[i], "Packing No", ',')
        packing_num_list.append(packing_num)
    # find unique packing number
    list_of_unique_num = []
    unique_num = set(packing_num_list)
    for num in unique_num:
        if num == '-':
            pass
        else:
            list_of_unique_num.append(num)
    print('Unique packing number: {}'.format(list_of_unique_num))
    return list_of_unique_num


if __name__ == '__main__':
    RT = '4557917'
    # receiving_data = get_receiving(RT)
    packing_number = find_packing_num(RT)