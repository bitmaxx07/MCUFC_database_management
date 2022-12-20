import jwt

key = "secret"
encoded = jwt.encode({"chengqi": ""}, key, algorithm="HS256")
print(encoded.replace(".", "-"))
print(encoded.replace(".", "-")[0: 75].lower())

print(jwt.decode(encoded, key, algorithms="HS256"))
