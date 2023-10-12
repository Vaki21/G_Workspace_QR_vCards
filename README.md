# G_Workspace_QR_vCards
Takes user data from Google Workspace, generates QR Code vCards and uploads them to a server.

As Google Workspace Admin I needed to assign QR Code vCards to e-mail signatures. In order to save time I made this code to fetch the needed data from GWS.
Thought many Admins might want to do the same so heres the whole project.

User data it works with: Name, Surname, Position, work e-mail, work number, office address. I also wanted to take user profile picture but couldn't make it work.

Required steps:
You will need an oAuth2 token from Google Cloud connected to Your GWS. (https://developers.google.com/identity/protocols/oauth2)
And for the server upload You need a public SSH key. (https://phoenixnap.com/kb/ssh-with-key)
