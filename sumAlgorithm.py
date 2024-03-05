def al (n):
    for i in range(len(arr)):
        if (arr[i]==n):
            print(arr[i])
            del arr[i]
        elif (arr[i]>n):
            continue
        else:
            al2()

def al2 ():
    arrLen = len(arr)
    s = arr[0]
    for i in range(arrLen):
        if (s + arr[i] == n):
            return 
            
            
def findMax (arr):
    max = 0
    for i in range(len(arr)):
        if (arr[i]>max):
            max = arr[i]
    return max

def findMin (arr):
    min = 100
    for i in range(len(arr)):
        if (arr[i]<min):
            min = arr[i]
    return min

arr = [8,5,4,6,2,7,2,1,4,5,6,10,8,9,2,3,5,6,8,2,7,4]
print(arr)
n = 10
al(n)