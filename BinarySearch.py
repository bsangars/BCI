import math
linearray=['Jack','Jill','Bhar','Bug','Junior','Alpha','Beta','Kappa','zebra']
linearray.sort()
def BinarySearchalgorithm(array,low,size,x):
    if size >=1:
        mid = int(math.floor(low +(size-low)/2))
        if array[mid] ==x:
            return mid
        elif array[mid]>x and array[low]<=x:
            return BinarySearchalgorithm(array,low,mid-1,x)
        elif array[mid]<x and  array[size]>=x:
            return BinarySearchalgorithm(array,mid+1,size,x)
        else:
            return -1
    else:
        return -1

x='Jack'
print("Linearsorted Array",linearray)
result =BinarySearchalgorithm(linearray,0,len(linearray)-1,x)
if result == -1:
    print("The element in the array is not found")
else:
    print("The array elemnet is found at index ",result)