fn = input('Enter file name: ')
try:
    file = open(fn, 'r')
except IOError:
    file = open(fn, 'w')
