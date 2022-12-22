import jwt
import openpyxl
import qrcode
import PIL

key = "secret"
# encoded = jwt.encode({"chengqi": ""}, key, algorithm="HS256")
# print(encoded.replace(".", "-"))
# print(encoded.replace(".", "-")[0: 75].lower())

# print(jwt.decode(encoded, key, algorithms="HS256"))

id_dict = {}

wb = openpyxl.load_workbook("namelist.xlsx")
ws = wb.active

col_name = 1
col_id = 2
col_subname = 4
col_website = 3

for row in range(2, ws.max_row + 1):
    # fill in excel sheet and generate id mapping
    encoded = jwt.encode({ws.cell(row, col_name).value: ""}, key, algorithm="HS256")
    ws.cell(row, col_id).value = encoded
    ws.cell(row, col_subname).value = encoded.replace(".", "-")[0: 75].lower()
    id_dict.update({ws.cell(row, col_name).value: encoded.replace(".", "-")[0: 75].lower()})
    ws.cell(row, col_website).value = "https://www.mcu-fc.com/" + encoded.replace(".", "-")[0: 75].lower()
    print("name: " + ws.cell(row, col_name).value)
    print("id: " + ws.cell(row, col_id).value)
    print("subname: " + ws.cell(row, col_subname).value)
    print("website: " + ws.cell(row, col_website).value)
    print("----------------------------")

    # generate qr codes
    input_data = "https://www.mcu-fc.com/" + encoded.replace(".", "-")[0: 75].lower()
    qr = qrcode.QRCode(version=1, box_size=10, border=0)
    qr.add_data(input_data)
    qr.make(fit=True)

    img = qr.make_image(fill="black", back_color="white")
    img.save("qrcodes/" + ws.cell(row, col_name).value + ".png")
    print("saved QR code for " + ws.cell(row, col_name).value)
    print("------------------------")


wb.save("namelist.xlsx")
print("-----------FINISHED----------")
print("TOTAL MAPPING:")
for key in id_dict.keys():
    print("name: " + key + " id: " + id_dict.get(key))
