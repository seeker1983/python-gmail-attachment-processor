# Python GMail Excel Attachment Processor

Download and send email attachments using gmail API.

# Installation

  pip3 install -r requirements.txt

# Generating Google App Certificate
Scripts use **credentials.json** file to authorizing google project.
Follow this instruction on how to create one.

https://docs.google.com/document/d/1sDe2wiwTWC7MlK7waHtOiLiHiT_2fTQif75zTXkQ9hc/edit?usp=sharing

# Initial run
On the first run of send or load you'll be directed in browser to allow access for you gmail account. 
This creates a token in **token.pickle** file which is used to all sub-sequent calls.
    
Refresh token have to be re-created on password change.
    
# Send example
    
    python send.py

or 

    import gmail
    gmail.create_message_with_attachment('me', 'krupod@gmail.com', 'Test email', "Contents of the body", 'test-attachment.xls')

Will send message from your account to **krupod@gmail.com** with subject **"Test email"**, body **"Contents of the body"**, and will attach local file **"test-attachment.xls"**

# Load example
    
    python load.py

or 

    gmail.load_unread('files', ['krupod@gmail.com'], "attachments_loaded")

Will download attachments from unread emails with "krupod@gmail.com" sender, save them in **"files"** folder and apply label **"attachments_loaded"** to them.

#### File list

  - **gmail.py** - main library
  - **load.py** - download example
  - **send.py** - send example
  - **requirements.txt** - dependencies list
  - **test-attachment.xls** - sample excel file (used by send example)
  - **credentials.json** - Google Project certificate. Not in repo, have to be created separately.
  - **token.pickle** - Refresh token for accessing gmail account is stored here. Will be created on first run after authorizing access to gmail.



