import os, os.path

# Connect to the server

DIR = 'Emails'
print (len([name for name in os.listdir(DIR) if os.path.isfile(os.path.join(DIR, name))]))