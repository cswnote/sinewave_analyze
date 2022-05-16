import math

max = 15
min = 12

def get_y_axis_min_max(max, min):
    max_scaled = 0
    min_scaled = 0
    digits = len(str(float(abs(max))).split('.')[0])
    if max > 0:
        max_scaled = math.ceil(max / (10 ** (digits - 1))) * 10 ** (digits - 1)
    else:
        max_scaled = math.ceil(max / (10 ** (digits - 1))) * 10 ** (digits - 1) * 1

    digits = len(str(float(abs(min))).split('.')[0])
    if min > 0 :
        min_scaled = math.ceil(min / (10 ** (digits - 1))) * 10 ** (digits - 1)
    else:
        min_scaled = math.floor(min / (10 ** (digits - 1))) * 10 ** (digits - 1) * 1

    if max_scaled == min_scaled or max_scaled < min_scaled:
        return max, min
    else:
        return max_scaled, min_scaled

print(get_y_axis_min_max(max, min))