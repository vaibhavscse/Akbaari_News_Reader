# Akbaari_News_Reader
A news reader that fetches latest news based on user input and reads it aloud to the user.
I have used the win32com library to convert text into audio.
Used my newsapi.org apikey to fetch latest news from newsapi servers based on genre.(earlier the website used to provide specific news about the genre. But now I think there is som problem with the website. Anyways just using apikey of some other website would resolve the problem)
Then I pulled request from newsapi.org and loaded it into a json file. 
I used Sapi voice to convert the input text into voice.
