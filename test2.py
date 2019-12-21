def num2columnletters(num, power=0):
    if num <= 26:
        return chr(num % 27 + 64)
    elif num > 26**(power+1):
        power += 1
        # this will return the higher (right most char) first
        char = num2columnletters(num=num-26**power, power=power)
        # then call func again on reminder
        char_next = chr(int(num/(26**(power-1))) + 64)
        char_all = char_next + char
    else:
        return chr(num % 26 + 64)
    return char_all

a=num2columnletters(11685)
pass
