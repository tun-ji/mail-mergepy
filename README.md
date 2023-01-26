I wrote this to perform a mail merge operation on Excel sheets to email student's grades to them. This was built to work with Gmail, but to get it started, you need to turn on 2FA on your Google Account, and generate an App Password.

Because I knew I'd be uploading this to Github, I put my login details in a .env file. Feel free to do the same.

If you'd like to try this out, make sure you have your excel data in this format (no headings):


| Column     | Format  |
| ------------ | --------- |
| S/N        | Integer |
| First Name | String  |
| Last Name  | String  |
| Score      | String  |
| Email      | String  |
