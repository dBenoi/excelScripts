# from openpyxl import load_workbook

# list of IPs (retrieved from excel)
ip_list = ["192.168.0.15", "192.168.3.8", "192.168.15.86"]
new_ip_list = []

# open and instantiate Excel file
###

###

# pull the data from excel and put it into ip_list
###

###

# function to add Zeros before single or double digit octets
def ip_zero_add(ip_address):
    octets = ip_address.split(".")
    newOctet = []

    for octet in octets:
        octetInt = int(octet)
        if octetInt < 10:
            octet = "00" + octet
        elif octetInt < 100:
            octet = "0" + octet
        elif octetInt > 255 or octetInt < 1:
            print("Please input a number between 1 and 255.")
    
        newOctet.append(octet)

    newIP = ".".join(newOctet)
    return newIP

# loop through IPs and append where needed if applicable
for ip in ip_list:
    modified_ip = ip_zero_add(ip)
    new_ip_list.append(modified_ip)

# loop through new_ip_list and append to excel file
for ip in new_ip_list:
    print(ip)