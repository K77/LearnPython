print("qiu ping jun!!!")

tolal = 0
count = 0

user_input = input("plz input num, with q end!!")

while user_input!="q":
    num = float(user_input)
    tolal += num
    count += 1
    user_input = input("plz input num, with q end!!")

result = tolal/count

print(result)

