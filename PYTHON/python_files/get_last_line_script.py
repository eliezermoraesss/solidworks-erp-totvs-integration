import inspect

def get_last_line_number(file):
        with open(file, 'r') as file:
            lines = file.readlines()
            return len(lines)

file = __file__
lastline = get_last_line_number(file)

print(lastline)
