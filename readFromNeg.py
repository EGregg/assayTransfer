with open('neg.txt') as f:
  credentials = [x.strip().split(':') for x in f.readlines()]

for username,password in credentials:
	print (username, password)

print (credentials);
