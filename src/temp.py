

def twoSum(nums: list[int], target: int) -> list[int]:
    for i, num_i in enumerate(nums):
        print(f"循环 i={i}, num={num_i}")
        found = False
        for j, num_j in enumerate(nums):
            print(f"循环 j={i}, num={num_j}")
            if num_i + num_j == target:
                found = True
                print(f"!!![1] found: sum: {nums[i] + nums[j]}; {i}: {nums[i]}; {j}: {nums[j]}; ")
                break
        if found:
            break
           

def twoSum2(nums: list[int], target: int) -> list[int]:
    for i, num_i in enumerate(nums):
        print(f"循环 i={i}, num={num_i}")
        found = False
        for j, num_j in enumerate(nums[i+1:], start=i+1): ###
            print(f"循环 j={i+1}, num={num_j}")
            if num_i + num_j == target:
                found = True
                print(f"!!![2] found: sum: {nums[i] + nums[j]}; {i}: {nums[i]}; {j}: {nums[j]}; ")
                break
        if found:
            break

def twoSum3(nums: list[int], target: int) -> list[int]:
    for i, num_i in enumerate(nums):
        difference = target - num_i
        print(f"循环 i={i}, num={num_i}; 差额: {difference}")
        found = False
        for j, num_j in enumerate(nums[i+1:], start=i+1): ###
            print(f"循环 j={i+1}, num={num_j}")
            if num_j == difference:
                found = True
                print(f"!!![3] found: sum: {nums[i] + nums[j]}; {i}: {nums[i]}; {j}: {nums[j]}; ")
                break
        if found:
            break
    
def twoSum4(nums: list[int], target: int) -> list[int]:
    seen = {}
    for i, num in enumerate(nums):
        print(f"循环 i={i}, num={num}")
        element = target - num
        if element in seen:
            j = seen[element]
            print(f"!!![4] found: sum: {nums[i] + nums[j]}; {i}: {nums[i]}; {j}: {nums[j]}; ")
            break
        seen[num] = i

samples = [
    ([2,7,11,15], 9), 
    # ([3,2,4], 6), ([3,3], 6)
]

for nums, target in samples:
    print(f"检测: {nums}, {target}")
    twoSum4(nums, target)


