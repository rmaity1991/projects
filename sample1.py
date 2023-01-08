

def prime():

    x=int(input("Kindly enter a number in integer format"))

    isprime=False


    if(x>1):
        for i in range(2,x):
            if(x%i==0):
                isprime=True;
                break
    else:
        print("Enter a value more than 1")



    if(isprime):
        print("The number is a composite number");
        for i in range(2,x):
            if(x%i==0):
                print(f"The factor of {x} is {i}")
    else:
        print("The number is prime");


    


def equation():

    for i in range(-5,6):
        y=8*(i**2)+1
        print(f"For the value of x : {i} , the value of y is {y}")



prime()

equation()


    


