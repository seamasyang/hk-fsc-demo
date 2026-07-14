def sum_two(nums:list[int], target:int):
    loop = 0
    for i, num_i in enumerate(nums):
        diff = target - num_i
        for j, num_j in enumerate(nums[i+1:]):
            loop += 1
            print(f"第{loop}次循环: i={i}, j={j}")
            
        #     if num_j==diff:
        #         print(f"===>找到: i({num_i})+j({num_j})={target}")
        #         break
        # if num_j==diff:
        #     break
    



samples = [
    ([2,7,11,15,9], 9), 
    # ([3,2,4], 6), ([3,3], 6)
]

for nums, target in samples:
    print(f"检测: 数字:{nums}, 和:{target}")
    sum_two(nums, target)
