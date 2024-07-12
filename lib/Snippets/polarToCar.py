import math

def polar_to_car(radius, angle_radians):
    x = radius * math.cos(angle_radians)
    y = radius * math.sin(angle_radians)
    return x, y

def polarDeg_to_car(radius, angle_degrees):
    angle_radians = math.radians(angle_degrees)
    return polar_to_car(radius, angle_radians)











def polar_to_car(radius, angle_degrees):
    x = radius * math.cos(math.radians(angle_degrees))
    y = radius * math.sin(math.radians(angle_degrees))
    return x, y

