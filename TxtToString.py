import os.path

# Connect to the server

DIR = 'Emails'
speak_output = ''
i = 1;
for file in os.listdir(DIR):
    speak_output += f"Email numero {i} \n \n"

    file = open (os.path.join(DIR, file), "r")
    speak_output += file.read(-1) + " \n \n"
    file.close()

    i += 1

