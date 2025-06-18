def find_two_numbers(nums, target):
    box = {}

    for i in range(len(nums)):
        current_number = nums[i]
        needed_number = target - current_number
        print(needed_number, "<- need")

        if needed_number in box:
            return [box[needed_number], i]
        else:
            print(box, "<- box")
            box[current_number] = i
            print(box, "<- box_next")

    return []


nums = [3, 2, 3]
target = 6

print(find_two_numbers(nums, target))
