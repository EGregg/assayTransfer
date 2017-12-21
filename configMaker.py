#!/usr/bin/env python3
import os


credentialsFile = 'storedCredentials.txt'

#creates credentials file if not present
def config_present():
	try:
	        open(credentialsFile,'r')
	except:
		open(credentialsFile,'w')

#asks for username and password to be written to credentials file
def config_writer():
	username = input('create username: ')
	password = input('create password: ')
	fd = open(credentialsFile,"w")
	fd.write("%s:%s" % (username,password))

#reads credentials file for user info, creates if not present
def config_read():
	with open(credentialsFile) as f:
		if os.stat(credentialsFile).st_size == 0:
			config_writer()
		
		credentials = [x.strip().split(':') for x in f.readlines()]

config_present()
config_read()
